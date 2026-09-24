import type { ICopyJetTemplate, IField } from '../model';
import type { IFieldInfoLike } from '../fields';

/** A list column in the context of its list (the provider's unit of work). */
export interface IListFieldDef {
  listKey: string;
  /** Site-relative list URL from the template; a renamed copy is found through the {listurl:K} token. */
  listUrl: string;
  field: IField;
}

export interface IListFieldInfoLike extends IFieldInfoLike {
  FromBaseType?: boolean;
  CanBeDeleted?: boolean;
}

export const LIST_FIELD_SELECT = ['Id', 'InternalName', 'Title', 'TypeAsString', 'Group', 'Description', 'Required', 'Hidden', 'SchemaXml', 'FromBaseType', 'CanBeDeleted'];

export const listFieldKey = (listKey: string, internalName: string): string => `listField:${listKey}/${internalName}`;

export type ListFieldOrigin = 'list' | 'siteColumn' | 'builtIn';

/**
 * Where a list's field comes from (spike 05):
 * - 'list': created on the list – SourceID is the list's GUID;
 * - 'siteColumn': a site column added to the list (custom or built-in) – not from the base type, deletable;
 * - 'builtIn': the list template's own fields (Title, Created …).
 */
export function listFieldOrigin(info: IListFieldInfoLike, listId: string): ListFieldOrigin {
  const m = /\sSourceID="\{?([0-9a-fA-F-]{36})\}?"/.exec(info.SchemaXml);
  if (m && m[1].toLowerCase() === listId.replace(/[{}]/g, '').toLowerCase()) {
    return 'list';
  }
  return !info.FromBaseType && info.CanBeDeleted ? 'siteColumn' : 'builtIn';
}

/** Fields CopyJet copies: visible list columns and site column instances. */
export function isCopiedListField(info: IListFieldInfoLike, listId: string): boolean {
  return !info.Hidden && listFieldOrigin(info, listId) !== 'builtIn';
}

/** All list columns of a template as provider inputs, in list order. */
export function listFieldDefs(template: ICopyJetTemplate): IListFieldDef[] {
  const out: IListFieldDef[] = [];
  template.lists.forEach((l) => (l.fields || []).forEach((field) => out.push({ listKey: l.key, listUrl: l.url, field })));
  return out;
}
