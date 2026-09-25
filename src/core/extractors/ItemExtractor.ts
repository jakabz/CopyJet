import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import { throwIfAborted } from '../errors';
import {
  NOT_YET_COPIED_TYPES,
  PrincipalCollector,
  SITE_USER_SELECT,
  itemsEntryPath,
  itemsKey,
  restPropertyOf,
  toItemField,
  toTemplateValue,
  type IItemField,
  type ISiteUserLike
} from '../items';
import { loadSourceSite, siteContentTypeIdOf } from '../lists';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, IItemsFile, ITemplateWriter } from '../model';
import type { IListItemsDef } from '../items';

/** Items are read in ID order, in pages filtered on ID (indexed), so large lists stay under the view threshold. */
export const ITEM_PAGE_SIZE = 2000;

type RawItem = { [prop: string]: unknown } & { ID: number; FileDirRef?: string; FSObjType?: number; ContentTypeId?: string };

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
      const list = sp.web.getList(listUrl);
      const infos = await list.fields.select('InternalName', 'TypeAsString', 'Hidden', 'ReadOnlyField')<
        Array<{ InternalName: string; TypeAsString: string; Hidden: boolean; ReadOnlyField: boolean }>
      >();
      const fields = infos.map(toItemField).filter((f): f is IItemField => !!f);
      infos
        .filter((f) => !f.Hidden && NOT_YET_COPIED_TYPES.indexOf(f.TypeAsString) >= 0)
        .forEach((f) =>
          opts.log.warn(`Column ${f.InternalName} (${f.TypeAsString}): its values are not copied yet.`, { artifact: ref, code: 'ITEM_FIELD_NOT_COPIED', detail: f.TypeAsString })
        );

      const select = ['ID', 'FileDirRef', 'FSObjType', 'ContentTypeId', 'AuthorId', 'EditorId', 'Created', 'Modified'].concat(fields.map(restPropertyOf));
      const file: IItemsFile = { listKey: l.key, items: [] };
      const ctx = { tokens: opts.tokens, principal: (id: number) => principals.token(id) };
      for (let last = 0; ; ) {
        throwIfAborted(opts.signal);
        const page = await list.items
          .select(...select)
          .filter(`ID gt ${last}`)
          .orderBy('ID', true)
          .top(ITEM_PAGE_SIZE)<RawItem[]>();
        page
          .filter((raw) => raw.FSObjType !== 1)
          .forEach((raw) => {
            const item: IItemsFile['items'][number] = { sourceId: raw.ID, values: {} };
            const dir = raw.FileDirRef || '';
            if (dir.toLowerCase().indexOf(`${listUrl.toLowerCase()}/`) === 0) item.folder = dir.slice(listUrl.length + 1);
            if (raw.ContentTypeId) item.contentType = siteContentTypeIdOf(raw.ContentTypeId);
            fields.forEach((f) => {
              const v = toTemplateValue(f, raw[restPropertyOf(f)], ctx);
              if (v !== undefined) item.values[f.internalName] = v;
            });
            const system: NonNullable<typeof item.system> = {};
            const author = typeof raw.AuthorId === 'number' ? principals.token(raw.AuthorId) : undefined;
            const editor = typeof raw.EditorId === 'number' ? principals.token(raw.EditorId) : undefined;
            if (author) system.author = author;
            if (editor) system.editor = editor;
            if (typeof raw.Created === 'string') system.created = raw.Created;
            if (typeof raw.Modified === 'string') system.modified = raw.Modified;
            if (Object.keys(system).length) item.system = system;
            file.items.push(item);
          });
        if (page.length < ITEM_PAGE_SIZE) break;
        last = page[page.length - 1].ID;
      }

      const source = itemsEntryPath(l.key);
      out.addJson(source, file);
      listDef.content = { mode: 'items', source, itemCount: file.items.length, includeAttachments: false };
      out.manifest.meta.includesContent = true;
      opts.log.info(`Items extracted: ${file.items.length}.`, { artifact: ref });
    }

    wanted
      .filter((key) => !site.lists.some((x) => itemsKey(x.key) === key))
      .forEach((key) => opts.log.warn('List not found on the source site.', { artifact: { kind: this.kind, key }, code: 'ITEMS_LIST_NOT_FOUND' }));
  }
}
