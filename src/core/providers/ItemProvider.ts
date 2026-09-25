import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import { CopyJetError, throwIfAborted } from '../errors';
import { isHttpStatus } from '../http/status';
import {
  dateValuesOf,
  isLookupKind,
  itemsKey,
  logItemFailures,
  peopleValue,
  principalKeysOf,
  principalLogins,
  readTargetFields,
  readWebLocale,
  runItemWrites,
  toTargetValue,
  webLocalTime,
  type IFormValue,
  type IListItemsDef,
  type ITargetField,
  type TemplateItem
} from '../items';
import { formatSpDateTime, type IWebLocale } from '../items/locale';
import type { LocalTimeConverter } from '../items/localTime';
import { createFolder, existingFolderPaths, listContentTypeFor, missingFolders, readListContentTypes, toServerRelativeUrl } from '../lists';
import type { ConflictMode, IApplyResult, IArtifactRef, IContentContext, IDiffResult, IInstallContext, IItemsFile, IProvider } from '../model';

export interface IItemProviderOptions {
  /** Write in $batch requests (default). Off in tests: the mock answers single requests only. */
  batched?: boolean;
}

/** Target list URL of an items step (a renamed copy is found through {listurl:K}). */
export function itemsListUrl(def: IListItemsDef, ctx: IInstallContext): string {
  const web = ctx.tokens.get('siterelative');
  if (web === undefined) {
    throw new CopyJetError('TOKEN_UNRESOLVED', 'The install context has no {siterelative} value.');
  }
  return toServerRelativeUrl(ctx.tokens.get('listurl', def.listKey) || def.listUrl, web);
}

export function contentContext(ctx: IInstallContext): IContentContext {
  if (!ctx.content) {
    throw new CopyJetError('CONTENT_CONTEXT_MISSING', 'The install has no content context (template package, ID maps).');
  }
  return ctx.content;
}

/** Form values of the system columns: people as picker JSON, dates in the web's format (spike 08 C/D). */
export function systemFormValues(item: TemplateItem, which: 'all' | 'modified', locale: IWebLocale, times: LocalTimeConverter, ctx: IInstallContext): IFormValue[] {
  const out: IFormValue[] = [];
  const s = item.system;
  if (!s) return out;
  const person = (name: string, token: string | undefined): void => {
    const logins = token ? principalLogins([token], ctx.tokens) : [];
    if (logins.length) out.push({ FieldName: name, FieldValue: peopleValue(logins) });
  };
  if (which === 'all') person('Author', s.author);
  person('Editor', s.editor);
  if (which === 'all' && s.created) out.push({ FieldName: 'Created', FieldValue: formatSpDateTime(times.local(s.created), locale) });
  if (s.modified) out.push({ FieldName: 'Modified', FieldValue: formatSpDateTime(times.local(s.modified), locale) });
  return out;
}

/**
 * Items of a list (first round): every value except lookups, which the 'itemLookups' step fills once the
 * ID maps of all lists exist. Written only into a list without items – without a run state CopyJet cannot
 * tell which items an earlier run already wrote, and it never duplicates or deletes them.
 */
export class ItemProvider implements IProvider<IListItemsDef> {
  public readonly kind = 'items' as const;
  private readonly _batched: boolean;

  constructor(options: IItemProviderOptions = {}) {
    this._batched = options.batched !== false;
  }

  public async diff(sp: SPFI, def: IListItemsDef, ctx: IInstallContext): Promise<IDiffResult> {
    throwIfAborted(ctx.signal);
    const ref: IArtifactRef = { kind: this.kind, key: itemsKey(def.listKey) };
    const list = sp.web.getList(itemsListUrl(def, ctx));
    let found: Array<{ Id: number }>;
    try {
      // Asked directly every time: ItemCount counts folders and is not a safe "empty" signal (spike 08 F).
      found = await list.items.filter('FSObjType eq 0').select('Id').top(1)<Array<{ Id: number }>>();
    } catch (e) {
      if (isHttpStatus(e, 404)) return { ref, status: 'unsupported', changes: ['listMissing'] };
      // Refused (large list without index): only a list that is empty by count is empty.
      const info = await list.select('ItemCount')<{ ItemCount: number }>();
      return info.ItemCount > 0 ? { ref, status: 'different', changes: ['targetHasItems'] } : { ref, status: 'new' };
    }
    return found.length > 0 ? { ref, status: 'different', changes: ['targetHasItems'] } : { ref, status: 'new' };
  }

  public async apply(sp: SPFI, def: IListItemsDef, _mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;
    if (diff.status === 'new') {
      return this._write(sp, def, ref, ctx);
    }
    if (diff.status === 'different') {
      ctx.log.warn('The target list already has items; nothing was added, so no item is duplicated.', { artifact: ref, code: 'ITEMS_TARGET_NOT_EMPTY' });
    } else {
      ctx.log.warn(`Items skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'ITEMS_UNSUPPORTED', detail: diff.changes });
    }
    return { ref, outcome: 'skipped' };
  }

  private async _write(sp: SPFI, def: IListItemsDef, ref: IArtifactRef, ctx: IInstallContext): Promise<IApplyResult> {
    const content = contentContext(ctx);
    const file = await content.reader.getJson<IItemsFile>(def.source, ctx.signal);
    ctx.log.info(`The target list has no items; writing ${file.items.length}.`, { artifact: ref, code: 'ITEMS_TARGET_EMPTY' });
    const listUrl = itemsListUrl(def, ctx);
    const fields = await readTargetFields(sp, listUrl, def.listKey, ctx.tokens);
    this._warnMissingFields(file.items, fields, ref, ctx);

    const locale = await readWebLocale(sp);
    const times = webLocalTime(sp);
    await times.prepare(dateValuesOf(file.items, fields, 'all'), ctx.signal);
    await content.principals.map(principalKeysOf(file.items, fields), ctx.tokens, ctx.log, ctx.signal);
    await this._ensureFolders(sp, listUrl, file.items, ref, ctx);
    const contentTypes = await readListContentTypes(sp, listUrl);

    const missingCts: { [id: string]: boolean } = {};
    const ops = file.items.map((item) => {
      const formValues: IFormValue[] = [];
      Object.keys(item.values).forEach((name) => {
        const f = fields[name];
        if (!f || isLookupKind(f.kind)) return;
        const v = toTargetValue(f, item.values[name], { locale, tokens: ctx.tokens, localTime: (iso) => times.local(iso) });
        if (v !== undefined) formValues.push({ FieldName: name, FieldValue: v });
      });
      if (item.contentType) {
        const ct = listContentTypeFor(item.contentType, contentTypes);
        if (ct) formValues.push({ FieldName: 'ContentTypeId', FieldValue: ct });
        else missingCts[item.contentType] = true;
      }
      formValues.push(...systemFormValues(item, 'all', locale, times, ctx));
      const folder = item.folder ? `${listUrl}/${item.folder}` : listUrl;
      return (s: SPFI) => s.web.getList(listUrl).addValidateUpdateItemUsingPath(formValues, folder, true);
    });
    Object.keys(missingCts).forEach((id) =>
      ctx.log.warn(`Content type ${id} is not on the target list; its items get the default content type.`, { artifact: ref, code: 'ITEM_CONTENT_TYPE_MISSING', detail: id })
    );

    const { results } = await runItemWrites(sp, ops, this._batched, ctx.signal);
    const idMap = (content.idMaps[def.listKey] = content.idMaps[def.listKey] || {});
    const failures: Array<{ sourceId: number; errors: string[] }> = [];
    results.forEach((r, i) => {
      const sourceId = file.items[i].sourceId;
      if (r.id) idMap[sourceId] = r.id;
      else failures.push({ sourceId, errors: r.errors || ['No item ID returned.'] });
    });
    logItemFailures(ctx, ref, failures);
    const written = file.items.length - failures.length;
    if (file.items.length > 0 && written === 0) {
      throw new CopyJetError('ITEMS_FAILED', `None of the ${file.items.length} items could be written.`, failures.slice(0, 5));
    }
    ctx.log.info(`Items written: ${written} of ${file.items.length}.`, { artifact: ref });
    return { ref, outcome: 'created' };
  }

  /** Values of columns the target list lacks would fail the whole item (spike 08 C): dropped with one warning each. */
  private _warnMissingFields(items: TemplateItem[], fields: { [name: string]: ITargetField }, ref: IArtifactRef, ctx: IInstallContext): void {
    const missing: { [name: string]: boolean } = {};
    items.forEach((item) => Object.keys(item.values).forEach((name) => !fields[name] && (missing[name] = true)));
    Object.keys(missing).forEach((name) =>
      ctx.log.warn(`Column ${name} is not a writable column of the target list; its values are skipped.`, { artifact: ref, code: 'ITEM_FIELD_MISSING', detail: name })
    );
  }

  /** Item folders missing on the target, parents first: an item in a missing folder fails with HTTP 500 (spike 08 E). */
  private async _ensureFolders(sp: SPFI, listUrl: string, items: TemplateItem[], ref: IArtifactRef, ctx: IInstallContext): Promise<void> {
    const wanted = items.map((i) => i.folder).filter((f, i, all): f is string => !!f && all.indexOf(f) === i);
    if (!wanted.length) return;
    // Items are copied into custom lists only (library files come with the file copy).
    const create = missingFolders(wanted, await existingFolderPaths(sp, listUrl, false, ctx.signal));
    for (const path of create) {
      throwIfAborted(ctx.signal);
      await createFolder(sp, listUrl, path, false);
    }
    if (create.length) ctx.log.info(`Folders created for items: ${create.length}.`, { artifact: ref });
  }
}
