import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/attachments';
import { CopyJetError, throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import { isHttpStatus } from '../http/status';
import {
  attachmentFileName,
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
  termKeysOf,
  toTargetValue,
  warnMissingFields,
  webLocalTime,
  type IFormValue,
  type IListItemsDef,
  type TemplateItem
} from '../items';
import { formatSpDateTime, type IWebLocale } from '../items/locale';
import type { LocalTimeConverter } from '../items/localTime';
import { createFolder, existingFolderPaths, listContentTypeFor, missingFolders, readListContentTypes, toServerRelativeUrl } from '../lists';
import type { ConflictMode, IApplyResult, IArtifactRef, IContentContext, IDiffResult, IInstallContext, IItemsFile, IProvider } from '../model';

/** Items written between two saves of the run state. */
const CHECKPOINT_CHUNK = 500;

/** Saves the run state mid-step when the install keeps one; a failed save does not stop the items. */
export async function saveCheckpoint(ctx: IInstallContext): Promise<void> {
  if (!ctx.checkpoint) return;
  try {
    await ctx.checkpoint();
  } catch (e) {
    ctx.log.warn(`The run state could not be saved: ${e instanceof Error ? e.message : String(e)}`, { code: 'STATE_SAVE_FAILED' });
  }
}

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
 * ID maps of all lists exist. Written into a list without items, or – resuming – into a list holding only
 * items the earlier run recorded; CopyJet never duplicates or deletes items.
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
    if (diff.status === 'different' && (await this._resumable(sp, def, ctx))) {
      return this._write(sp, def, ref, ctx, true);
    }
    if (diff.status === 'different') {
      ctx.log.warn('The target list already has items; nothing was added, so no item is duplicated.', { artifact: ref, code: 'ITEMS_TARGET_NOT_EMPTY' });
    } else {
      ctx.log.warn(`Items skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'ITEMS_UNSUPPORTED', detail: diff.changes });
    }
    return { ref, outcome: 'skipped' };
  }

  /**
   * Resume (spike 14): the list has items, but each is either in the ID map the earlier run saved, or one of the
   * chunk it was writing when it stopped (`pendingItems`). SharePoint writes a batch in order, so the unrecorded
   * items, by ascending ID, are a prefix of that chunk; each pair is checked by title. Then the rest is written
   * without duplicates. Any other item (added by someone, or not matching) leaves the list as it is.
   */
  private async _resumable(sp: SPFI, def: IListItemsDef, ctx: IInstallContext): Promise<boolean> {
    const content = ctx.content;
    if (!content) return false;
    const known = content.idMaps[def.listKey] || {};
    const pending = (content.pendingItems && content.pendingItems[def.listKey]) || [];
    if (!Object.keys(known).length && !pending.length) return false;
    const ref: IArtifactRef = { kind: this.kind, key: itemsKey(def.listKey) };
    const mine: { [id: number]: boolean } = {};
    Object.keys(known).forEach((k) => (mine[known[Number(k)]] = true));
    const unknown: Array<{ Id: number; Title?: string }> = [];
    // Id/FSObjType without a filter: paging by ID stays under the list view threshold.
    for await (const page of sp.web.getList(itemsListUrl(def, ctx)).items.select('Id', 'FSObjType', 'Title').top(5000)) {
      throwIfAborted(ctx.signal);
      (page as Array<{ Id: number; FSObjType: number; Title?: string }>).forEach((i) => Number(i.FSObjType) === 0 && !mine[i.Id] && unknown.push(i));
    }
    const refuse = (): boolean => {
      ctx.log.warn('The target list has items the earlier run did not record; nothing was added, so no item is duplicated.', { artifact: ref, code: 'ITEMS_RESUME_UNKNOWN_ITEMS' });
      return false;
    };
    if (!unknown.length) return true;
    if (unknown.length > pending.length) return refuse();
    unknown.sort((x, y) => x.Id - y.Id);
    const file = await content.reader.getJson<IItemsFile>(def.source, ctx.signal);
    const bySource: { [id: number]: TemplateItem } = {};
    file.items.forEach((i) => (bySource[i.sourceId] = i));
    const titleMatches = (target: { Title?: string }, sourceId: number): boolean => {
      const item = bySource[sourceId];
      const title = item && item.values.Title;
      return typeof title !== 'string' || title === (target.Title || '');
    };
    if (!unknown.every((t, i) => titleMatches(t, pending[i]))) return refuse();
    const idMap = (content.idMaps[def.listKey] = known);
    unknown.forEach((t, i) => (idMap[pending[i]] = t.Id));
    ctx.log.info(`${unknown.length} items written just before the interruption were recognized.`, { artifact: ref, code: 'ITEMS_RESUME_RECOVERED', detail: unknown.length });
    return true;
  }

  private async _write(sp: SPFI, def: IListItemsDef, ref: IArtifactRef, ctx: IInstallContext, resuming: boolean = false): Promise<IApplyResult> {
    const content = contentContext(ctx);
    const file = await content.reader.getJson<IItemsFile>(def.source, ctx.signal);
    const idMap = (content.idMaps[def.listKey] = content.idMaps[def.listKey] || {});
    const todo = resuming ? file.items.filter((i) => !idMap[i.sourceId]) : file.items;
    if (resuming) {
      ctx.log.info(`Resuming: ${file.items.length - todo.length} items were written by the earlier run; writing ${todo.length}.`, { artifact: ref, code: 'ITEMS_RESUMED' });
    } else {
      ctx.log.info(`The target list has no items; writing ${file.items.length}.`, { artifact: ref, code: 'ITEMS_TARGET_EMPTY' });
    }
    const listUrl = itemsListUrl(def, ctx);
    const fields = await readTargetFields(sp, listUrl, def.listKey, ctx.tokens);
    warnMissingFields(todo, fields, ref, ctx);

    const locale = await readWebLocale(sp);
    const times = webLocalTime(sp);
    await times.prepare(dateValuesOf(file.items, fields, 'all'), ctx.signal);
    await content.principals.map(principalKeysOf(file.items, fields), ctx.tokens, ctx.log, ctx.signal);
    const terms = content.terms;
    if (terms) await terms.map(termKeysOf(todo, fields), ctx.log, ctx.signal);
    await this._ensureFolders(sp, listUrl, todo, ref, ctx);
    const contentTypes = await readListContentTypes(sp, listUrl);

    const missingCts: { [id: string]: boolean } = {};
    const ops = todo.map((item) => {
      const formValues: IFormValue[] = [];
      Object.keys(item.values).forEach((name) => {
        const f = fields[name];
        if (!f || isLookupKind(f.kind)) return;
        const v = toTargetValue(f, item.values[name], { locale, tokens: ctx.tokens, localTime: (iso) => times.local(iso), termValues: (keys) => (terms ? terms.resolved(keys) : []) });
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

    // In chunks, with the run state saved after each: an interrupted install knows which items it wrote.
    const failures: Array<{ sourceId: number; errors: string[] }> = [];
    const pending = (content.pendingItems = content.pendingItems || {});
    for (let start = 0; start < ops.length; start += CHECKPOINT_CHUNK) {
      // The chunk is saved before it is written: a resume can recognize what an interruption left unrecorded.
      pending[def.listKey] = todo.slice(start, start + CHECKPOINT_CHUNK).map((i) => i.sourceId);
      await saveCheckpoint(ctx);
      const { results } = await runItemWrites(sp, ops.slice(start, start + CHECKPOINT_CHUNK), this._batched, ctx.signal);
      results.forEach((r, i) => {
        const sourceId = todo[start + i].sourceId;
        if (r.id) idMap[sourceId] = r.id;
        else failures.push({ sourceId, errors: r.errors || ['No item ID returned.'] });
      });
      delete pending[def.listKey];
      await saveCheckpoint(ctx);
    }
    logItemFailures(ctx, ref, failures);
    const written = todo.length - failures.length;
    if (todo.length > 0 && written === 0) {
      throw new CopyJetError('ITEMS_FAILED', `None of the ${todo.length} items could be written.`, failures.slice(0, 5));
    }
    ctx.log.info(`Items written: ${written} of ${todo.length}.`, { artifact: ref });
    await this._attachments(sp, listUrl, file.items, idMap, ref, ctx, (item) => systemFormValues(item, 'modified', locale, times, ctx));
    return { ref, outcome: 'created' };
  }

  /**
   * Attachments of the written items, one item after the other within 4 in flight (spike 09 B). Each file makes
   * a new version and sets Editor/Modified to the installer, so those are sent again afterwards; the extra
   * versions stay in the history (known limit). A name already on the item (HTTP 400) is skipped.
   */
  private async _attachments(
    sp: SPFI,
    listUrl: string,
    items: TemplateItem[],
    idMap: { [sourceId: number]: number },
    ref: IArtifactRef,
    ctx: IInstallContext,
    systemValues: (item: TemplateItem) => IFormValue[]
  ): Promise<void> {
    const content = contentContext(ctx);
    const todo = items.filter((item) => item.attachments && item.attachments.length && idMap[item.sourceId]);
    if (!todo.length) return;
    let added = 0;
    const failures: Array<{ sourceId: number; errors: string[] }> = [];
    const results = await limitConcurrency(
      todo.map((item) => async () => {
        const target = sp.web.getList(listUrl).items.getById(idMap[item.sourceId]);
        const errors: string[] = [];
        for (const path of item.attachments!) {
          throwIfAborted(ctx.signal);
          try {
            await target.attachmentFiles.add(attachmentFileName(path), await content.reader.getBlob(path, ctx.signal));
            added++;
          } catch (e) {
            if (isHttpStatus(e, 400) && /-2130575257/.test(e instanceof Error ? e.message : '')) continue; // already there
            errors.push(`${attachmentFileName(path)}: ${e instanceof Error ? e.message : String(e)}`);
          }
        }
        const restore = systemValues(item);
        if (restore.length) await target.validateUpdateListItem(restore, true);
        if (errors.length) failures.push({ sourceId: item.sourceId, errors });
      }),
      4,
      ctx.signal
    );
    results.forEach((r, i) => {
      if (!r.ok) failures.push({ sourceId: todo[i].sourceId, errors: [r.error instanceof Error ? r.error.message : String(r.error)] });
    });
    failures.forEach((f) =>
      ctx.log.warn(`Attachments of item ${f.sourceId} could not all be added: ${f.errors.join('; ')}`, { artifact: ref, code: 'ATTACHMENT_FAILED', detail: f })
    );
    ctx.log.info(`Attachments added: ${added}.`, { artifact: ref });
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
