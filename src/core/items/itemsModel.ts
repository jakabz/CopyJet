import type { ICopyJetTemplate, IList } from '../model';

/** The items of one list (provider input of both the 'items' and the 'itemLookups' step). */
export interface IListItemsDef {
  listKey: string;
  /** Site-relative list URL from the template; a renamed copy is found through the {listurl:K} token. */
  listUrl: string;
  /** Package entry holding the items (items/<listkey>.json). */
  source: string;
  /** Keys of lists the list's lookup columns point to (itself included). */
  lookupTargets: string[];
}

export const itemsKey = (listKey: string): string => `items:${listKey}`;
export const itemLookupsKey = (listKey: string): string => `itemLookups:${listKey}`;
export const itemsEntryPath = (listKey: string): string => `items/${listKey}.json`;

/** Target list keys of the list's lookup columns (from their {listkey:X} lookupList tokens). */
export function lookupTargets(list: IList): string[] {
  const out: string[] = [];
  (list.fields || [])
    .filter((f) => f.type === 'Lookup' || f.type === 'LookupMulti')
    .forEach((f) => {
      const m = f.lookupList ? /\{listkey:([^{}]+)\}/.exec(f.lookupList) : null;
      if (m && out.indexOf(m[1]) < 0) out.push(m[1]);
    });
  return out;
}

/** Lists of the template whose items are in the package. */
export function listItemsDefs(template: ICopyJetTemplate): IListItemsDef[] {
  return template.lists
    .filter((l) => l.content && l.content.mode === 'items' && !!l.content.source)
    .map((l) => ({ listKey: l.key, listUrl: l.url, source: l.content.source!, lookupTargets: lookupTargets(l) }));
}
