import { contentTypeKey } from '../contentTypes/contentTypeModel';
import type { IListItemsDef } from '../items/itemsModel';
import type { IListFieldDef } from '../lists/listFields';
import type { IListViewDef } from '../lists/views';
import type { IArtifactRef, IContentType, IField, IGroup, IList } from '../model';

/**
 * What each artifact needs installed before it. Shared by the Setup (offering missing dependencies) and the
 * planner (install order). Refs may point to artifacts outside the template (built-in content types, base
 * columns …); the caller keeps only the ones it knows.
 */

function listKeyOf(tokenString: string | undefined): string | undefined {
  const m = tokenString ? /\{listkey:([^{}]+)\}/.exec(tokenString) : null;
  return m ? m[1] : undefined;
}

export const artifactKeys = {
  group: (key: string): string => `group:${key}`,
  siteField: (internalName: string): string => `field:${internalName}`,
  contentType: (id: string): string => contentTypeKey(id),
  list: (key: string): string => `list:${key}`,
  listField: (listKey: string, internalName: string): string => `listField:${listKey}/${internalName}`,
  view: (listKey: string, title: string): string => `view:${listKey}/${title}`,
  items: (listKey: string): string => `items:${listKey}`,
  itemLookups: (listKey: string): string => `itemLookups:${listKey}`
};

/** A group owning this one. */
export function groupDependencies(def: IGroup): IArtifactRef[] {
  const m = def.owner ? /^\{groupkey:([^{}]+)\}$/.exec(def.owner) : null;
  return m && m[1] !== def.key ? [{ kind: 'group', key: artifactKeys.group(m[1]) }] : [];
}

/** A lookup site column needs its target list. */
export function siteFieldDependencies(def: IField): IArtifactRef[] {
  const listKey = listKeyOf(def.lookupList);
  return listKey ? [{ kind: 'list', key: artifactKeys.list(listKey) }] : [];
}

/** The parent content type and the site columns the type links. */
export function contentTypeDependencies(def: IContentType): IArtifactRef[] {
  const refs: IArtifactRef[] = def.fieldRefs.map((r) => ({ kind: 'siteField' as const, key: artifactKeys.siteField(r.internalName) }));
  if (def.parentId && def.parentId !== '0x') {
    refs.unshift({ kind: 'contentType', key: artifactKeys.contentType(def.parentId) });
  }
  return refs;
}

/** Content types the list uses. */
export function listDependencies(def: IList): IArtifactRef[] {
  return (def.contentTypes || []).filter((ct) => /^0x/i.test(ct)).map((ct) => ({ kind: 'contentType' as const, key: artifactKeys.contentType(ct) }));
}

/** Its list, a lookup's target list, and the site column it may be an instance of. */
export function listFieldDependencies(def: IListFieldDef): IArtifactRef[] {
  const refs: IArtifactRef[] = [{ kind: 'list', key: artifactKeys.list(def.listKey) }];
  const target = listKeyOf(def.field.lookupList);
  if (target && target !== def.listKey) refs.push({ kind: 'list', key: artifactKeys.list(target) });
  refs.push({ kind: 'siteField', key: artifactKeys.siteField(def.field.internalName) });
  return refs;
}

/** Its list and the columns it shows. */
export function viewDependencies(def: IListViewDef): IArtifactRef[] {
  const refs: IArtifactRef[] = [{ kind: 'list', key: artifactKeys.list(def.listKey) }];
  return refs.concat(def.view.fields.map((f) => ({ kind: 'listField' as const, key: artifactKeys.listField(def.listKey, f) })));
}

/** Its list with every column and content type of it: an item fails when one of its columns is missing. */
export function itemsDependencies(def: IListItemsDef, list: IList): IArtifactRef[] {
  const refs: IArtifactRef[] = [{ kind: 'list', key: artifactKeys.list(def.listKey) }];
  (list.fields || []).forEach((f) => refs.push({ kind: 'listField', key: artifactKeys.listField(def.listKey, f.internalName) }));
  return refs.concat(listDependencies(list));
}

/**
 * The list's own items and the items of every list its lookups point to (their ID maps). Only items steps:
 * lookup rounds never wait for each other, so cyclic lookups (A → B → A) still install.
 */
export function itemLookupsDependencies(def: IListItemsDef): IArtifactRef[] {
  const refs: IArtifactRef[] = [{ kind: 'items', key: artifactKeys.items(def.listKey) }];
  def.lookupTargets.filter((k) => k !== def.listKey).forEach((k) => refs.push({ kind: 'items', key: artifactKeys.items(k) }));
  return refs;
}
