import type { SPFI } from '@pnp/sp';
import type { Logger } from '../logger/Logger';
import type { ITermRef } from '../model';
import { TermStoreClient, termPaths, type ITermSetInfo } from '../taxonomy';

export type TermStrategy = 'sameId' | 'path';

export interface ITermMapping {
  key: string;
  /** The target term; undefined when it was not found. */
  termId?: string;
  /** The target term's own label (what "Label|GUID" is written with). */
  label?: string;
  strategy?: TermStrategy;
}

interface ISetTerms {
  byId: { [id: string]: string };
  byPath: { [path: string]: { id: string; label: string } };
}

/**
 * Template terms → target terms (rendszerterv §7): the same term when the target store has it (same tenant),
 * else the term at the same label path "Group/Set/…/Label" (another tenant, spike 11). A term not found stays
 * out of the values, with one warning per log: an unknown GUID would fail the whole item (spike 11 C).
 */
export class TermMapper {
  private readonly _client: TermStoreClient;
  private readonly _byKey: { [key: string]: ITermRef } = {};
  private readonly _done: { [key: string]: ITermMapping } = {};
  private readonly _sets: { [setId: string]: Promise<ISetTerms | undefined> } = {};
  private readonly _setByPath: { [path: string]: Promise<ITermSetInfo | undefined> } = {};
  private readonly _warned: Map<Logger, string[]> = new Map();

  /** One mapper per loaded template, with its own term store cache (never shared with an earlier load). */
  constructor(sp: SPFI, terms: ITermRef[]) {
    this._client = new TermStoreClient(sp);
    terms.forEach((t) => (this._byKey[t.key] = t));
  }

  public get keys(): string[] {
    return Object.keys(this._byKey);
  }

  /** Maps the given keys (default: every template term); warns once per log about terms not found. */
  public async map(keys: string[] = this.keys, log?: Logger, signal?: AbortSignal): Promise<ITermMapping[]> {
    const wanted = keys.filter((k, i) => keys.indexOf(k) === i);
    for (const key of wanted.filter((k) => !this._done[k])) {
      this._done[key] = await this._map(key, signal);
    }
    if (log) {
      const warned = this._warned.get(log) || [];
      this._warned.set(log, warned);
      wanted
        .filter((k) => !this._done[k].termId && warned.indexOf(k) < 0)
        .forEach((k) => {
          warned.push(k);
          const t = this._byKey[k];
          log.warn(`Term not found on the target: ${t ? t.path : k}. Its values are left empty.`, { code: 'TERM_NOT_FOUND', detail: t ? { key: k, path: t.path } : { key: k } });
        });
    }
    return wanted.map((k) => this._done[k]);
  }

  /** Already mapped terms of these keys, for "Label|GUID" values (unmapped ones are left out). */
  public resolved(keys: string[]): Array<{ termId: string; label: string }> {
    return keys
      .map((k) => this._done[k])
      .filter((m): m is ITermMapping & { termId: string; label: string } => !!m && !!m.termId && !!m.label)
      .map((m) => ({ termId: m.termId, label: m.label }));
  }

  private async _map(key: string, signal?: AbortSignal): Promise<ITermMapping> {
    const t = this._byKey[key];
    if (!t) return { key };
    const same = t.termSetId ? await this._terms(t.termSetId, signal) : undefined;
    const id = t.termId.toLowerCase();
    if (same && same.byId[id] !== undefined) return { key, termId: id, label: same.byId[id], strategy: 'sameId' };
    // By label path: the set by its "Group/Set" prefix, then the term by the whole path.
    const parts = t.path.split('/');
    if (parts.length < 3) return { key };
    const setPath = `${parts[0]}/${parts[1]}`;
    if (!this._setByPath[setPath]) this._setByPath[setPath] = this._client.findSet(parts[0], parts[1]);
    const set = await this._setByPath[setPath];
    const terms = set ? await this._terms(set.id, signal) : undefined;
    const hit = terms ? terms.byPath[t.path.toLowerCase()] : undefined;
    return hit ? { key, termId: hit.id, label: hit.label, strategy: 'path' } : { key };
  }

  private _terms(setId: string, signal?: AbortSignal): Promise<ISetTerms | undefined> {
    const id = setId.toLowerCase();
    if (!this._sets[id]) {
      this._sets[id] = (async () => {
        const set = await this._client.set(id);
        if (!set) return undefined;
        const terms = await this._client.terms(set.id, signal);
        const paths = termPaths(set, terms);
        const out: ISetTerms = { byId: {}, byPath: {} };
        terms.forEach((x) => {
          out.byId[x.id] = x.label;
          out.byPath[paths[x.id].toLowerCase()] = { id: x.id, label: x.label };
        });
        return out;
      })();
    }
    return this._sets[id];
  }
}
