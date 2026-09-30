import { SPQueryable, type SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import { throwIfAborted } from '../errors';
import { isHttpStatus } from '../http/status';

/**
 * The tenant term store through SharePoint's _api/v2.1/termStore (spike 11 A): PnPjs 4 has no taxonomy
 * module, so the requests are made with SPQueryable on the web's behaviors (retry, throttling, client tag).
 * The flat term list has no parents, so term paths come from walking the children (spike 11 B).
 */

interface ILabel {
  name: string;
  languageTag?: string;
  isDefault?: boolean;
}

interface ITermJson {
  id: string;
  labels?: ILabel[];
  childrenCount?: number;
}

export interface ITermSetInfo {
  id: string;
  name: string;
  groupName: string;
}

export interface ITermInfo {
  id: string;
  label: string;
  /** Undefined for a top-level term. */
  parentId?: string;
}

const lower = (id: string): string => id.replace(/[{}]/g, '').toLowerCase();

/** The default label, else the first one. */
export function defaultLabel(labels: ILabel[] | undefined): string {
  const list = labels || [];
  const d = list.filter((l) => l.isDefault)[0] || list[0];
  return d ? d.name : '';
}

/** Collection responses arrive as an array (PnPjs unwraps "value"); tolerate both. */
const asArray = <T>(r: unknown): T[] => (Array.isArray(r) ? (r as T[]) : ((r as { value?: T[] }) || {}).value || []);

export class TermStoreClient {
  private readonly _sp: SPFI;
  private readonly _base: string;
  private _storeId?: Promise<string>;
  private readonly _sets: { [id: string]: Promise<ITermSetInfo | undefined> } = {};
  private readonly _terms: { [setId: string]: Promise<ITermInfo[]> } = {};

  constructor(sp: SPFI) {
    this._sp = sp;
    this._base = `${sp.web.toUrl().replace(/\/_api\/web\/?$/i, '')}/_api/v2.1/termStore`;
  }

  /** The default term store's ID (= a taxonomy column's SspId, spike 11 A). */
  public storeId(): Promise<string> {
    if (!this._storeId) this._storeId = this._get<{ id: string }>('?$select=id').then((s) => lower(s.id));
    return this._storeId;
  }

  /** A term set with its group, or undefined when the store has no such set. */
  public set(setId: string): Promise<ITermSetInfo | undefined> {
    const id = lower(setId);
    if (!this._sets[id]) {
      this._sets[id] = this._get<{ id: string; localizedNames?: ILabel[]; parentGroup?: { name: string } }>(`/sets/${id}?$expand=parentGroup`).then(
        (s) => ({ id: lower(s.id), name: defaultLabel(s.localizedNames), groupName: s.parentGroup ? s.parentGroup.name : '' }),
        (e: unknown) => {
          if (isHttpStatus(e, 404) || isHttpStatus(e, 400)) return undefined;
          throw e;
        }
      );
    }
    return this._sets[id];
  }

  /** Every term of a set with its parent, parents before children. */
  public terms(setId: string, signal?: AbortSignal): Promise<ITermInfo[]> {
    const id = lower(setId);
    if (!this._terms[id]) this._terms[id] = this._walk(id, signal);
    return this._terms[id];
  }

  /** The set with this group and set name (a cross-tenant install finds term sets by path). */
  public async findSet(groupName: string, setName: string): Promise<ITermSetInfo | undefined> {
    const groups = asArray<{ id: string; name: string }>(await this._get('/groups?$select=id,name'));
    const group = groups.filter((g) => g.name === groupName)[0];
    if (!group) return undefined;
    const sets = asArray<{ id: string; localizedNames?: ILabel[] }>(await this._get(`/groups/${group.id}/sets?$select=id,localizedNames`));
    const set = sets.filter((s) => (s.localizedNames || []).some((n) => n.name === setName))[0];
    return set ? { id: lower(set.id), name: setName, groupName } : undefined;
  }

  private async _walk(setId: string, signal?: AbortSignal): Promise<ITermInfo[]> {
    const out: ITermInfo[] = [];
    const queue: Array<{ path: string; parentId?: string }> = [{ path: `/sets/${setId}/children` }];
    while (queue.length) {
      throwIfAborted(signal);
      const next = queue.shift()!;
      const children = asArray<ITermJson>(await this._get(`${next.path}?$select=id,labels,childrenCount`));
      children.forEach((t) => {
        const id = lower(t.id);
        out.push({ id, label: defaultLabel(t.labels), parentId: next.parentId });
        if ((t.childrenCount || 0) > 0) queue.push({ path: `/sets/${setId}/terms/${id}/children`, parentId: id });
      });
    }
    return out;
  }

  private _get<T>(path: string): Promise<T> {
    return SPQueryable([this._sp.web, `${this._base}${path}`])<T>();
  }
}

/** "Group/Set/Parent/Term" for every term (the label path that identifies a term across tenants). */
export function termPaths(set: ITermSetInfo, terms: ITermInfo[]): { [termId: string]: string } {
  const byId: { [id: string]: ITermInfo } = {};
  terms.forEach((t) => (byId[t.id] = t));
  const out: { [id: string]: string } = {};
  const pathOf = (t: ITermInfo): string => {
    if (out[t.id]) return out[t.id];
    const parent = t.parentId ? byId[t.parentId] : undefined;
    out[t.id] = `${parent ? pathOf(parent) : `${set.groupName}/${set.name}`}/${t.label}`;
    return out[t.id];
  };
  terms.forEach(pathOf);
  return out;
}

const clients = new WeakMap<object, TermStoreClient>();

/**
 * One cached client per run: `scope` is the install context (or the preview's), so a later install in the same
 * page never reuses a stale term list – a term deleted in between must be seen as missing (2026-09-30).
 */
export function termStoreFor(sp: SPFI, scope: object): TermStoreClient {
  let client = clients.get(scope);
  if (!client) {
    client = new TermStoreClient(sp);
    clients.set(scope, client);
  }
  return client;
}

export interface ITargetTermSet {
  termStoreId: string;
  termSetId: string;
  /** 'id' in the same term store, 'path' found by group/set name (another tenant). */
  matchedBy: 'id' | 'path';
}

/**
 * The target's term set for a template column: the same set when the store has it (same tenant), else the set
 * at the template's "Group/Set" path (spike 11 B). Undefined when neither exists.
 */
export async function resolveTargetTermSet(client: TermStoreClient, termSet: { termSetId: string; path: string }): Promise<ITargetTermSet | undefined> {
  const termStoreId = await client.storeId();
  const byId = await client.set(termSet.termSetId);
  if (byId) return { termStoreId, termSetId: byId.id, matchedBy: 'id' };
  const i = termSet.path.indexOf('/');
  if (i <= 0) return undefined;
  const found = await client.findSet(termSet.path.slice(0, i), termSet.path.slice(i + 1));
  return found ? { termStoreId, termSetId: found.id, matchedBy: 'path' } : undefined;
}
