import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/files';
import '@pnp/sp/site-users/web';
import { throwIfAborted } from '../errors';
import { DEFAULT_MAX_FILE_BYTES, fileEntryPath, fileVersionEntryPath, filesFolderPath, filesKey, filesMetaPath, type IListFilesDef } from '../files';
import { limitConcurrency } from '../http/concurrency';
import { PrincipalCollector, SITE_USER_SELECT, forEachItemPage, itemSelect, readSourceColumns, toTemplateItem, type ISiteUserLike, type RawItem } from '../items';
import { loadSourceSite } from '../lists';
import { TermCollector, TermStoreClient, prepareTerms } from '../taxonomy';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, IFilesMetaFile, ITemplateWriter } from '../model';

type FileEntry = IFilesMetaFile['files'][number];

/** SP.CheckOutType.None: a checked-out file (0 online, 1 offline) is skipped (rendszerterv §7). */
const NOT_CHECKED_OUT = 2;

interface IVersionInfoLike {
  ID: number;
  VersionLabel: string;
  Created: string;
  Size?: number | string;
  CheckInComment?: string;
  CreatedBy?: { Id?: number };
}

/**
 * Files of document libraries (spike 10 A): current content, metadata like list items and, per library on
 * request, earlier versions. Written to files/<listkey>/_meta.json plus one entry per file (version).
 * Selected through 'files:<listkey>' refs; the library must already be in the template.
 */
export class FileExtractor implements IExtractor<IListFilesDef> {
  public readonly kind = 'files' as const;

  /** Content is a per-library option in the Setup, not a tree entry of its own. */
  public async discover(): Promise<IDiscoveredArtifact[]> {
    return [];
  }

  public dependencies(def: IListFilesDef): IArtifactRef[] {
    return [{ kind: 'list', key: `list:${def.listKey}` }];
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) return;
    const [site, users] = await Promise.all([loadSourceSite(sp, opts.signal), sp.web.siteUsers.select(...SITE_USER_SELECT)<ISiteUserLike[]>()]);
    const principals = new PrincipalCollector(out.manifest.principals, users);
    const terms = new TermCollector((out.manifest.terms = out.manifest.terms || []), new TermStoreClient(sp));
    const maxBytes = opts.maxFileBytes || DEFAULT_MAX_FILE_BYTES;

    for (const l of site.lists.filter((x) => wanted.indexOf(filesKey(x.key)) >= 0)) {
      throwIfAborted(opts.signal);
      const ref: IArtifactRef = { kind: this.kind, key: filesKey(l.key) };
      const listDef = out.manifest.lists.filter((x) => x.key === l.key)[0];
      if (!listDef) {
        opts.log.warn(`Library ${l.info.Title} is not in the template; its files are skipped.`, { artifact: ref, code: 'FILES_LIST_MISSING' });
        continue;
      }
      const listUrl = l.info.RootFolder.ServerRelativeUrl;
      const columns = await readSourceColumns(sp, listUrl, ref, opts.log);
      await prepareTerms(terms, columns.fields, (message, detail) => opts.log.warn(message, { artifact: ref, code: 'TERM_SET_NOT_FOUND', detail }), opts.signal);
      const withVersions = (opts.versionsFor || []).indexOf(l.key) >= 0;
      const ctx = { listUrl, columns, tokens: opts.tokens, principals, terms: (id: string) => terms.key(id), preserveAuthors: opts.preserveAuthors };
      const found: Array<{ raw: RawItem; entry: FileEntry }> = [];
      await forEachItemPage(
        sp,
        listUrl,
        itemSelect(columns.fields, ['FileRef', 'FileLeafRef', 'File/Length', 'File/CheckOutType']),
        ['File'],
        (page) =>
          page
            .filter((raw) => raw.FSObjType !== 1)
            .forEach((raw) => {
              const file = (raw.File || {}) as { Length?: string | number; CheckOutType?: number };
              const item = toTemplateItem(raw, ctx);
              const path = item.folder ? `${item.folder}/${raw.FileLeafRef}` : String(raw.FileLeafRef);
              const size = Number(file.Length) || 0;
              if (file.CheckOutType !== undefined && file.CheckOutType !== NOT_CHECKED_OUT) {
                opts.log.warn(`${path} is checked out; it is skipped.`, { artifact: ref, code: 'FILE_CHECKED_OUT', detail: path });
                return;
              }
              if (size > maxBytes) {
                opts.log.warn(`${path} (${Math.round(size / 1048576)} MB) is larger than the limit; it is skipped.`, { artifact: ref, code: 'FILE_TOO_LARGE', detail: { path, size } });
                return;
              }
              const entry: FileEntry = { path, blob: fileEntryPath(l.key, path), sizeBytes: size, sourceId: item.sourceId, values: item.values };
              if (item.contentType) entry.contentType = item.contentType;
              if (item.system) entry.system = item.system;
              if (entry.blob !== `files/${l.key}/${path}`) {
                opts.log.warn(`${path} is stored in the package as ${entry.blob} (".." is not allowed in package paths).`, { artifact: ref, code: 'FILE_RENAMED', detail: path });
              }
              found.push({ raw, entry });
            }),
        opts.signal
      );

      const results = await limitConcurrency(
        found.map((f) => () => this._download(sp, l.key, f.raw, f.entry.path, f.entry.blob!, withVersions, principals, opts, out)),
        4,
        opts.signal
      );
      const files: FileEntry[] = [];
      let versionBytes = 0;
      results.forEach((r, i) => {
        if (r.ok) {
          if (r.value.versions.length) found[i].entry.versions = r.value.versions;
          versionBytes += r.value.bytes;
          files.push(found[i].entry);
        } else {
          opts.log.warn(`${found[i].entry.path} could not be read: ${r.error instanceof Error ? r.error.message : String(r.error)}`, {
            artifact: ref,
            code: 'FILE_READ_FAILED',
            detail: found[i].entry.path
          });
        }
      });
      // Package size estimate: current files plus the earlier versions read along.
      const sizeBytes = files.reduce((n, f) => n + f.sizeBytes, 0) + versionBytes;
      const meta: IFilesMetaFile = { listKey: l.key, files };
      out.addJson(filesMetaPath(l.key), meta);
      listDef.content = { mode: 'files', sourceMode: 'embedded', includeVersions: withVersions, fileCount: files.length, sizeBytes, source: filesFolderPath(l.key) };
      out.manifest.meta.includesContent = true;
      const versions = files.reduce((n, f) => n + (f.versions || []).length, 0);
      opts.log.info(`Files extracted: ${files.length} (${Math.round(sizeBytes / 1024)} KB)${withVersions ? `, earlier versions: ${versions}` : ''}.`, { artifact: ref });
    }

    wanted
      .filter((key) => !site.lists.some((x) => filesKey(x.key) === key))
      .forEach((key) => opts.log.warn('Library not found on the source site.', { artifact: { kind: this.kind, key }, code: 'FILES_LIST_NOT_FOUND' }));
  }

  /** The file (and its earlier versions, oldest first) into the package; returns the versions and their bytes. */
  private async _download(
    sp: SPFI,
    listKey: string,
    raw: RawItem,
    path: string,
    blobPath: string,
    withVersions: boolean,
    principals: PrincipalCollector,
    opts: IExtractOptions,
    out: ITemplateWriter
  ): Promise<{ versions: NonNullable<FileEntry['versions']>; bytes: number }> {
    const file = sp.web.getFileByServerRelativePath(String(raw.FileRef));
    out.addBlob(blobPath, await file.getBlob());
    if (!withVersions) return { versions: [], bytes: 0 };
    const versions = await file.versions.select('ID', 'VersionLabel', 'Created', 'CheckInComment', 'Size', 'CreatedBy/Id').expand('CreatedBy')<IVersionInfoLike[]>();
    let bytes = 0;
    const list: NonNullable<FileEntry['versions']> = [];
    for (const v of versions.slice().sort((a, b) => a.ID - b.ID)) {
      throwIfAborted(opts.signal);
      const blob = fileVersionEntryPath(listKey, v.VersionLabel, path);
      out.addBlob(blob, await file.versions.getById(v.ID).getBlob());
      const version: NonNullable<FileEntry['versions']>[number] = { label: v.VersionLabel, blob };
      const editor = v.CreatedBy && typeof v.CreatedBy.Id === 'number' ? principals.token(v.CreatedBy.Id) : undefined;
      if (opts.preserveAuthors !== false) version.system = editor ? { editor, modified: v.Created } : { modified: v.Created };
      if (v.CheckInComment) version.comment = v.CheckInComment;
      list.push(version);
      bytes += Number(v.Size) || 0;
    }
    return { versions: list, bytes };
  }
}
