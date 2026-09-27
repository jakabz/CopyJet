import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/files';
import '@pnp/sp/site-users/web';
import { throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import {
  ITEM_PAGE_SIZE,
  PrincipalCollector,
  SITE_USER_SELECT,
  attachmentEntryPath,
  forEachItemPage,
  itemSelect,
  itemsEntryPath,
  itemsKey,
  readSourceColumns,
  toTemplateItem,
  type ISiteUserLike
} from '../items';
import { loadSourceSite } from '../lists';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, IItemsFile, ITemplateWriter } from '../model';
import type { IListItemsDef } from '../items';

export { ITEM_PAGE_SIZE };

/**
 * List items (not folders – the ListExtractor carries those) into items/<listkey>.json. Selected per list
 * through 'items:<listkey>' refs; the list itself must already be in the template.
 */
export class ItemExtractor implements IExtractor<IListItemsDef> {
  public readonly kind = 'items' as const;

  /** Content is a per-list option in the Setup, not a tree entry of its own. */
  public async discover(): Promise<IDiscoveredArtifact[]> {
    return [];
  }

  public dependencies(def: IListItemsDef): IArtifactRef[] {
    return [{ kind: 'list', key: `list:${def.listKey}` }];
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) return;
    const [site, users] = await Promise.all([loadSourceSite(sp, opts.signal), sp.web.siteUsers.select(...SITE_USER_SELECT)<ISiteUserLike[]>()]);
    const principals = new PrincipalCollector(out.manifest.principals, users);

    for (const l of site.lists.filter((x) => wanted.indexOf(itemsKey(x.key)) >= 0)) {
      throwIfAborted(opts.signal);
      const ref: IArtifactRef = { kind: this.kind, key: itemsKey(l.key) };
      const listDef = out.manifest.lists.filter((x) => x.key === l.key)[0];
      if (!listDef) {
        opts.log.warn(`List ${l.info.Title} is not in the template; its items are skipped.`, { artifact: ref, code: 'ITEMS_LIST_MISSING' });
        continue;
      }
      const listUrl = l.info.RootFolder.ServerRelativeUrl;
      const columns = await readSourceColumns(sp, listUrl, ref, opts.log);
      const file: IItemsFile = { listKey: l.key, items: [] };
      const ctx = { listUrl, columns, tokens: opts.tokens, principals, preserveAuthors: opts.preserveAuthors };
      const withFiles: number[] = [];
      await forEachItemPage(
        sp,
        listUrl,
        itemSelect(columns.fields, ['Attachments']),
        [],
        (page) =>
          page
            .filter((raw) => raw.FSObjType !== 1)
            .forEach((raw) => {
              if (raw.Attachments) withFiles.push(raw.ID);
              file.items.push(toTemplateItem(raw, ctx));
            }),
        opts.signal
      );

      const withAttachments = file.items.filter((item) => withFiles.indexOf(item.sourceId) >= 0);
      const attachments = withAttachments.length ? await this._attachments(sp, listUrl, l.key, withAttachments, ref, opts, out) : 0;

      const source = itemsEntryPath(l.key);
      out.addJson(source, file);
      listDef.content = { mode: 'items', source, itemCount: file.items.length, includeAttachments: !!l.info.EnableAttachments };
      out.manifest.meta.includesContent = true;
      opts.log.info(`Items extracted: ${file.items.length}${attachments ? `, attachments: ${attachments}` : ''}.`, { artifact: ref });
    }

    wanted
      .filter((key) => !site.lists.some((x) => itemsKey(x.key) === key))
      .forEach((key) => opts.log.warn('List not found on the source site.', { artifact: { kind: this.kind, key }, code: 'ITEMS_LIST_NOT_FOUND' }));
  }

  /**
   * Downloads the attachments of the given items into the package (spike 09 A) and lists them on the items.
   * Four items at a time; returns the number of files.
   */
  private async _attachments(
    sp: SPFI,
    listUrl: string,
    listKey: string,
    items: IItemsFile['items'],
    ref: IArtifactRef,
    opts: IExtractOptions,
    out: ITemplateWriter
  ): Promise<number> {
    let count = 0;
    const results = await limitConcurrency(
      items.map((item) => async () => {
        const files = await sp.web.getList(listUrl).items.getById(item.sourceId).attachmentFiles.select('FileName', 'ServerRelativeUrl')<Array<{ FileName: string; ServerRelativeUrl: string }>>();
        for (const f of files) {
          throwIfAborted(opts.signal);
          const path = attachmentEntryPath(listKey, item.sourceId, f.FileName);
          if (path.slice(path.lastIndexOf('/') + 1) !== f.FileName) {
            opts.log.warn(`Attachment "${f.FileName}" of item ${item.sourceId} is stored as "${path.slice(path.lastIndexOf('/') + 1)}" (".." is not allowed in package paths).`, {
              artifact: ref,
              code: 'ATTACHMENT_RENAMED',
              detail: f.FileName
            });
          }
          out.addBlob(path, await sp.web.getFileByServerRelativePath(f.ServerRelativeUrl).getBlob());
          (item.attachments = item.attachments || []).push(path);
          count++;
        }
      }),
      4,
      opts.signal
    );
    results.forEach((r, i) => {
      if (!r.ok) {
        opts.log.warn(`Attachments of item ${items[i].sourceId} could not be read: ${r.error instanceof Error ? r.error.message : String(r.error)}`, {
          artifact: ref,
          code: 'ATTACHMENT_READ_FAILED',
          detail: items[i].sourceId
        });
      }
    });
    return count;
  }
}
