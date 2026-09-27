import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import { throwIfAborted } from '../errors';
import { readListContentTypes, siteContentTypeIdOf, type IListContentTypes } from '../lists/listStructure';
import type { Logger } from '../logger/Logger';
import type { IArtifactRef, IItemsFile } from '../model';
import type { TokenContext } from '../tokenizer';
import { isNotYetCopied, restPropertyOf, toItemField, toTemplateValue, type IItemField } from './fieldValues';
import type { PrincipalCollector } from './principals';

/** Items are read in ID order, in pages filtered on ID (indexed), so large lists stay under the view threshold. */
export const ITEM_PAGE_SIZE = 2000;

export type RawItem = { [prop: string]: unknown } & {
  ID: number;
  FileDirRef?: string;
  FileLeafRef?: string;
  FSObjType?: number;
  ContentTypeId?: string;
  Attachments?: boolean;
};

export interface ISourceColumns {
  fields: IItemField[];
  contentTypes: IListContentTypes;
}

/** The columns whose values are copied, and the list's content types; warns once per column not copied yet. */
export async function readSourceColumns(sp: SPFI, listUrl: string, ref: IArtifactRef, log: Logger): Promise<ISourceColumns> {
  const [infos, contentTypes] = await Promise.all([
    sp.web.getList(listUrl).fields.select('InternalName', 'TypeAsString', 'Hidden', 'ReadOnlyField')<Array<{ InternalName: string; TypeAsString: string; Hidden: boolean; ReadOnlyField: boolean }>>(),
    readListContentTypes(sp, listUrl)
  ]);
  infos
    .filter(isNotYetCopied)
    .forEach((f) => log.warn(`Column ${f.InternalName} (${f.TypeAsString}): its values are not copied yet.`, { artifact: ref, code: 'ITEM_FIELD_NOT_COPIED', detail: f.TypeAsString }));
  return { fields: infos.map(toItemField).filter((f): f is IItemField => !!f), contentTypes };
}

/** The item properties to select: identity, place, content type, system values and the copied columns. */
export function itemSelect(fields: IItemField[], extra: string[] = []): string[] {
  return ['ID', 'FileDirRef', 'FSObjType', 'ContentTypeId', 'AuthorId', 'EditorId', 'Created', 'Modified'].concat(extra, fields.map(restPropertyOf));
}

/** Calls `onPage` with every page of the list's items (folders included), in ID order. */
export async function forEachItemPage(
  sp: SPFI,
  listUrl: string,
  select: string[],
  expand: string[],
  onPage: (rows: RawItem[]) => void | Promise<void>,
  signal?: AbortSignal
): Promise<void> {
  for (let last = 0; ; ) {
    throwIfAborted(signal);
    let query = sp.web.getList(listUrl).items.select(...select);
    if (expand.length) query = query.expand(...expand);
    const page = await query.filter(`ID gt ${last}`).orderBy('ID', true).top(ITEM_PAGE_SIZE)<RawItem[]>();
    await onPage(page);
    if (page.length < ITEM_PAGE_SIZE) return;
    last = page[page.length - 1].ID;
  }
}

export interface IToTemplateItemContext {
  listUrl: string;
  columns: ISourceColumns;
  tokens: TokenContext;
  principals: PrincipalCollector;
  /** Keep author, editor and dates (default true). */
  preserveAuthors?: boolean;
}

/** A REST item row → the template's item: folder, site content type, values, system values. */
export function toTemplateItem(raw: RawItem, ctx: IToTemplateItemContext): IItemsFile['items'][number] {
  const item: IItemsFile['items'][number] = { sourceId: raw.ID, values: {} };
  const dir = raw.FileDirRef || '';
  if (dir.toLowerCase().indexOf(`${ctx.listUrl.toLowerCase()}/`) === 0) item.folder = dir.slice(ctx.listUrl.length + 1);
  if (raw.ContentTypeId) item.contentType = siteContentTypeIdOf(raw.ContentTypeId, ctx.columns.contentTypes);
  const valueCtx = { tokens: ctx.tokens, principal: (id: number) => ctx.principals.token(id) };
  ctx.columns.fields.forEach((f) => {
    const v = toTemplateValue(f, raw[restPropertyOf(f)], valueCtx);
    if (v !== undefined) item.values[f.internalName] = v;
  });
  const system: NonNullable<typeof item.system> = {};
  const author = typeof raw.AuthorId === 'number' ? ctx.principals.token(raw.AuthorId) : undefined;
  const editor = typeof raw.EditorId === 'number' ? ctx.principals.token(raw.EditorId) : undefined;
  if (author) system.author = author;
  if (editor) system.editor = editor;
  if (typeof raw.Created === 'string') system.created = raw.Created;
  if (typeof raw.Modified === 'string') system.modified = raw.Modified;
  if (ctx.preserveAuthors !== false && Object.keys(system).length) item.system = system;
  return item;
}
