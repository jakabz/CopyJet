import { toTemplateKey } from '../keys';
import type { IField, ITermRef } from '../model';
import { termPaths, type TermStoreClient } from './termStore';

/**
 * Collects the terms item values refer to into the template's terms and hands out their keys. Label and
 * path come from the term store (a single-value column's REST "Label" is its WssId, spike 11 A).
 */
export class TermCollector {
  private readonly _terms: ITermRef[];
  private readonly _client: TermStoreClient;
  private readonly _known: { [termId: string]: { termSetId: string; label: string; path: string } } = {};
  private readonly _keys: { [termId: string]: string } = {};

  /** `terms` is the template's collection (appended to). */
  constructor(terms: ITermRef[], client: TermStoreClient) {
    this._terms = terms;
    this._client = client;
    terms.forEach((t) => (this._keys[t.termId.toLowerCase()] = t.key));
  }

  /** Loads the terms of these sets (once each); sets missing from the store are returned. */
  public async prepare(termSetIds: string[], signal?: AbortSignal): Promise<string[]> {
    const missing: string[] = [];
    for (const setId of termSetIds.filter((s, i, all) => all.indexOf(s) === i)) {
      const set = await this._client.set(setId);
      if (!set) {
        missing.push(setId);
        continue;
      }
      const terms = await this._client.terms(set.id, signal);
      const paths = termPaths(set, terms);
      terms.forEach((t) => (this._known[t.id] = { termSetId: set.id, label: t.label, path: paths[t.id] }));
    }
    return missing;
  }

  /** The template key of a term (added to the template on first use), or undefined for an unknown term. */
  public key(termId: string): string | undefined {
    const id = termId.replace(/[{}]/g, '').toLowerCase();
    if (this._keys[id]) return this._keys[id];
    const t = this._known[id];
    if (!t) return undefined;
    const base = toTemplateKey(t.label, 'term');
    const used = (k: string): boolean => this._terms.some((x) => x.key.toLowerCase() === k.toLowerCase());
    let key = base;
    for (let n = 2; used(key); n++) key = `${base}_${n}`;
    this._terms.push({ key, termId: id, termSetId: t.termSetId, label: t.label, path: t.path });
    this._keys[id] = key;
    return key;
  }
}

/**
 * Loads the term sets of the given Managed Metadata columns into the collector; warns once per set that is no
 * longer in the term store (its values are then not carried).
 */
export async function prepareTerms(
  collector: TermCollector,
  fields: Array<{ internalName: string; termSetId?: string }>,
  warn: (message: string, detail: unknown) => void,
  signal?: AbortSignal
): Promise<void> {
  const sets = fields.map((f) => f.termSetId).filter((s): s is string => !!s);
  if (!sets.length) return;
  const missing = await collector.prepare(sets, signal);
  fields
    .filter((f) => f.termSetId && missing.indexOf(f.termSetId) >= 0)
    .forEach((f) => warn(`Column ${f.internalName}: its term set is not in the term store; its values are not copied.`, { column: f.internalName, termSetId: f.termSetId }));
}

/**
 * A Managed Metadata column's term set path ("Group/Set") from the term store, into def.termSet.path – what a
 * cross-tenant install finds the set by (spike 11 B). Warns when the set is gone.
 */
export async function fillTermSetPath(def: IField, client: TermStoreClient, warn: (message: string) => void): Promise<void> {
  const termSet = def.termSet;
  if (!termSet || /^0{8}-/.test(termSet.termSetId)) return;
  const set = await client.set(termSet.termSetId);
  if (set) {
    termSet.path = `${set.groupName}/${set.name}`;
  } else {
    warn(`Column ${def.internalName}: its term set is not in the term store; on install it can only be bound in the same tenant.`);
  }
}
