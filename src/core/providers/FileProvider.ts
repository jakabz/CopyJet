import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { CopyJetError, isAbortError, throwIfAborted } from '../errors';
import { CHUNK_SIZE, filesKey, type IListFilesDef } from '../files';
import { limitConcurrency } from '../http/concurrency';
import { isHttpStatus } from '../http/status';
import {
  dateValuesOf,
  forEachItemPage,
  isLookupKind,
  lookupValue,
  principalKeysOf,
  readTargetFields,
  readWebLocale,
  toTargetValue,
  warnMissingFields,
  webLocalTime,
  type IFormValue,
  type TemplateItem
} from '../items';
import { createFolder, existingFolderPaths, listContentTypeFor, missingFolders, readListContentTypes } from '../lists';
import type { ConflictMode, IApplyResult, IArtifactRef, IDiffResult, IFilesMetaFile, IInstallContext, IProvider } from '../model';
import { contentContext, itemsListUrl, systemFormValues } from './ItemProvider';

type FileEntry = IFilesMetaFile['files'][number];

/** SharePoint's "a file with this name already exists" (spike 10 B2, HTTP 400). */
const ALREADY_EXISTS = /-2130575257/;

const isAlreadyThere = (e: unknown): boolean => isHttpStatus(e, 400) && ALREADY_EXISTS.test(e instanceof Error ? e.message : String(e));

/** A file entry as the item helpers see it (values, system values). */
const asItem = (f: { sourceId?: number; values?: FileEntry['values']; system?: FileEntry['system'] }): TemplateItem => ({
  sourceId: f.sourceId || 0,
  values: (f.values || {}) as TemplateItem['values'],
  system: f.system
});

/**
 * Files of a document library (spike 10 B). Never overwrites: a path that already exists is kept and counted,
 * so a rerun only adds what is missing. Up to 10 MB a file goes in one request, above that in chunks. The
 * metadata (values, content type, author, editor, dates) follows with ValidateUpdateListItem, which on a
 * library makes no new version. Earlier versions, when carried, are uploaded oldest first, each followed by
 * its editor and date; their author stays the installer (known limit).
 */
export class FileProvider implements IProvider<IListFilesDef> {
  public readonly kind = 'files' as const;

  public async diff(sp: SPFI, def: IListFilesDef, ctx: IInstallContext): Promise<IDiffResult> {
    throwIfAborted(ctx.signal);
    const ref: IArtifactRef = { kind: this.kind, key: filesKey(def.listKey) };
    let found: Array<{ Id: number }>;
    try {
      found = await sp.web.getList(itemsListUrl(def, ctx)).items.filter('FSObjType eq 0').select('Id').top(1)<Array<{ Id: number }>>();
    } catch (e) {
      if (isHttpStatus(e, 404)) return { ref, status: 'unsupported', changes: ['listMissing'] };
      found = [{ Id: 0 }]; // refused on a large library: treat as having files; missing ones are still added
    }
    return found.length ? { ref, status: 'different', changes: ['targetHasFiles'] } : { ref, status: 'new' };
  }

  public async apply(sp: SPFI, def: IListFilesDef, _mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    if (diff.status === 'unsupported') {
      ctx.log.warn(`Files skipped: ${(diff.changes || []).join(', ')}.`, { artifact: diff.ref, code: 'FILES_UNSUPPORTED', detail: diff.changes });
      return { ref: diff.ref, outcome: 'skipped' };
    }
    return this._write(sp, def, diff.ref, diff.status === 'new', ctx);
  }

  private async _write(sp: SPFI, def: IListFilesDef, ref: IArtifactRef, empty: boolean, ctx: IInstallContext): Promise<IApplyResult> {
    const content = contentContext(ctx);
    const meta = await content.reader.getJson<IFilesMetaFile>(def.source, ctx.signal);
    const listUrl = itemsListUrl(def, ctx);
    const existing = empty ? {} : await this._existingFiles(sp, listUrl, ctx);
    const todo = meta.files.filter((f) => !existing[f.path.toLowerCase()]);
    const kept = meta.files.length - todo.length;
    if (kept) ctx.log.info(`Files already present and kept: ${kept}.`, { artifact: ref, code: 'FILES_KEPT' });
    if (!todo.length) return { ref, outcome: 'skipped' };

    const fields = await readTargetFields(sp, listUrl, def.listKey, ctx.tokens);
    const items = todo.map(asItem);
    warnMissingFields(items, fields, ref, ctx);
    const versionItems = ([] as TemplateItem[]).concat(...todo.map((f) => (f.versions || []).map((v) => asItem({ system: v.system }))));
    const locale = await readWebLocale(sp);
    const times = webLocalTime(sp);
    await times.prepare(dateValuesOf(items, fields, 'all').concat(dateValuesOf(versionItems, {}, 'modified')), ctx.signal);
    await content.principals.map(principalKeysOf(items.concat(versionItems), fields), ctx.tokens, ctx.log, ctx.signal);
    await this._ensureFolders(sp, listUrl, todo, ref, ctx);
    const contentTypes = await readListContentTypes(sp, listUrl);
    const idMap = (content.idMaps[def.listKey] = content.idMaps[def.listKey] || {});

    const unresolved: { [name: string]: boolean } = {};
    const missingCts: { [id: string]: boolean } = {};
    const formValuesOf = (f: FileEntry): IFormValue[] => {
      const values: IFormValue[] = [];
      Object.keys(f.values || {}).forEach((name) => {
        const field = fields[name];
        const value = (f.values || {})[name];
        if (!field) return;
        let v: string | undefined;
        if (isLookupKind(field.kind)) {
          // Lookup targets' items were written before this step (planner); their ID maps resolve the values.
          const map = field.lookupListKey ? content.idMaps[field.lookupListKey] : undefined;
          if (!map) unresolved[name] = true;
          else v = lookupValue(field, ((value as { lookup?: number[] } | null)?.lookup || []).map((id) => map[id]).filter((id) => !!id));
        } else {
          v = toTargetValue(field, value, { locale, tokens: ctx.tokens, localTime: (iso) => times.local(iso) });
        }
        if (v !== undefined) values.push({ FieldName: name, FieldValue: v });
      });
      if (f.contentType) {
        const ct = listContentTypeFor(f.contentType, contentTypes);
        if (ct) values.push({ FieldName: 'ContentTypeId', FieldValue: ct });
        else missingCts[f.contentType] = true;
      }
      return values.concat(systemFormValues(asItem(f), 'all', locale, times, ctx));
    };

    let added = 0;
    let alreadyThere = 0;
    const failures: Array<{ path: string; error: string }> = [];
    const results = await limitConcurrency(
      todo.map((f) => async () => {
        const id = await this._upload(sp, listUrl, f, content.reader, (v) => systemFormValues(asItem({ system: v.system }), 'modified', locale, times, ctx), ctx);
        if (id === undefined) {
          alreadyThere++;
          return;
        }
        await sp.web.getList(listUrl).items.getById(id).validateUpdateListItem(formValuesOf(f), true);
        if (f.sourceId) idMap[f.sourceId] = id;
        added++;
      }),
      3,
      ctx.signal
    );
    results.forEach((r, i) => {
      if (!r.ok) {
        if (isAbortError(r.error)) throw r.error;
        failures.push({ path: todo[i].path, error: r.error instanceof Error ? r.error.message : String(r.error) });
      }
    });
    Object.keys(unresolved).forEach((name) =>
      ctx.log.warn(`Column ${name}: the items of its lookup list were not written in this run; its values are skipped.`, { artifact: ref, code: 'ITEM_LOOKUP_UNRESOLVED', detail: name })
    );
    Object.keys(missingCts).forEach((id) =>
      ctx.log.warn(`Content type ${id} is not on the target library; its files get the default content type.`, { artifact: ref, code: 'ITEM_CONTENT_TYPE_MISSING', detail: id })
    );
    failures.slice(0, 20).forEach((f) => ctx.log.warn(`${f.path} could not be copied: ${f.error}`, { artifact: ref, code: 'FILE_FAILED', detail: f }));
    if (failures.length > 20) ctx.log.warn(`${failures.length - 20} more files could not be copied.`, { artifact: ref, code: 'FILE_FAILED_MORE' });
    if (alreadyThere) ctx.log.info(`Files already present and kept: ${alreadyThere}.`, { artifact: ref, code: 'FILES_KEPT' });
    if (added === 0 && failures.length > 0) {
      throw new CopyJetError('FILES_FAILED', `None of the ${todo.length} files could be copied.`, failures.slice(0, 5));
    }
    ctx.log.info(`Files copied: ${added} of ${todo.length}.`, { artifact: ref });
    return { ref, outcome: empty ? 'created' : 'updated' };
  }

  /**
   * Uploads the file's versions oldest first and the current content last; returns the item ID, or undefined
   * when the path already exists (nothing is overwritten). After each earlier version its editor and date are
   * sent again, so the version history keeps the source dates.
   */
  private async _upload(
    sp: SPFI,
    listUrl: string,
    f: FileEntry,
    reader: NonNullable<IInstallContext['content']>['reader'],
    versionValues: (v: NonNullable<FileEntry['versions']>[number]) => IFormValue[],
    ctx: IInstallContext
  ): Promise<number | undefined> {
    const i = f.path.lastIndexOf('/');
    const folderUrl = i < 0 ? listUrl : `${listUrl}/${f.path.slice(0, i)}`;
    const name = f.path.slice(i + 1);
    const folder = sp.web.getFolderByServerRelativePath(folderUrl);
    const file = sp.web.getFileByServerRelativePath(`${folderUrl}/${name}`);
    const steps: Array<{ blob: string; version?: NonNullable<FileEntry['versions']>[number] }> = (f.versions || []).map((v) => ({ blob: v.blob, version: v }));
    steps.push({ blob: f.blob! });
    let id: number | undefined;
    for (let n = 0; n < steps.length; n++) {
      throwIfAborted(ctx.signal);
      const blob = await reader.getBlob(steps[n].blob, ctx.signal);
      try {
        if (blob.size <= CHUNK_SIZE) {
          await folder.files.addUsingPath(name, blob, { Overwrite: n > 0 });
        } else {
          if (n === 0) await folder.files.addUsingPath(name, '', { Overwrite: false });
          await file.setContentChunked(blob, { chunkSize: CHUNK_SIZE });
        }
      } catch (e) {
        if (n === 0 && isAlreadyThere(e)) return undefined;
        throw e;
      }
      if (id === undefined) id = (await file.listItemAllFields.select('Id')<{ Id: number }>()).Id;
      const version = steps[n].version;
      if (version) {
        const values = versionValues(version);
        if (values.length) await sp.web.getList(listUrl).items.getById(id).validateUpdateListItem(values, true);
      }
    }
    return id;
  }

  /** Paths (lower case, relative to the library) of the files already there; paged by ID, so large libraries work. */
  private async _existingFiles(sp: SPFI, listUrl: string, ctx: IInstallContext): Promise<{ [path: string]: boolean }> {
    const out: { [path: string]: boolean } = {};
    const root = `${listUrl.toLowerCase()}/`;
    await forEachItemPage(
      sp,
      listUrl,
      ['ID', 'FSObjType', 'FileRef'],
      [],
      (page) =>
        page
          .filter((r) => r.FSObjType !== 1 && typeof r.FileRef === 'string' && (r.FileRef as string).toLowerCase().indexOf(root) === 0)
          .forEach((r) => (out[(r.FileRef as string).toLowerCase().slice(root.length)] = true)),
      ctx.signal
    );
    return out;
  }

  /** Folders of the files, parents first (the ListProvider creates the template's folders; this covers the rest). */
  private async _ensureFolders(sp: SPFI, listUrl: string, files: FileEntry[], ref: IArtifactRef, ctx: IInstallContext): Promise<void> {
    const wanted = files.map((f) => f.path.slice(0, Math.max(0, f.path.lastIndexOf('/')))).filter((p, i, all) => !!p && all.indexOf(p) === i);
    if (!wanted.length) return;
    const create = missingFolders(wanted, await existingFolderPaths(sp, listUrl, true, ctx.signal));
    for (const path of create) {
      throwIfAborted(ctx.signal);
      await createFolder(sp, listUrl, path, true);
    }
    if (create.length) ctx.log.info(`Folders created for files: ${create.length}.`, { artifact: ref });
  }
}

