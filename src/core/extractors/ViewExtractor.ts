import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/views';
import { throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import {
  COPIED_VIEWS_FILTER,
  VIEW_SELECT,
  isSupportedTemplate,
  loadSourceSite,
  toViewDef,
  viewKey,
  type IListViewDef,
  type IViewInfoLike
} from '../lists';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, ITemplateWriter } from '../model';

/** Public views of lists; written into the template's list entry, which the ListExtractor must have created. */
export class ViewExtractor implements IExtractor<IListViewDef> {
  public readonly kind = 'view' as const;

  public async discover(sp: SPFI, signal?: AbortSignal): Promise<IDiscoveredArtifact[]> {
    const site = await loadSourceSite(sp, signal);
    const lists = site.lists.filter((l) => isSupportedTemplate(l.info.BaseTemplate));
    const results = await limitConcurrency(lists.map((l) => () => this._views(sp, l.info.RootFolder.ServerRelativeUrl)), 4, signal);
    const out: IDiscoveredArtifact[] = [];
    results.forEach((r, i) => {
      if (!r.ok) return;
      r.value.forEach((v) => {
        const found: IDiscoveredArtifact = { ref: { kind: this.kind, key: viewKey(lists[i].key, v.Title) }, title: v.Title, parentKey: `list:${lists[i].key}` };
        if (v.ViewType && v.ViewType !== 'HTML') found.unsupported = 'VIEW_TYPE_UNSUPPORTED';
        out.push(found);
      });
    });
    return out;
  }

  /** The list, and every column the view shows (the Setup keeps only list columns it discovered). */
  public dependencies(def: IListViewDef): IArtifactRef[] {
    const refs: IArtifactRef[] = [{ kind: 'list', key: `list:${def.listKey}` }];
    return refs.concat(def.view.fields.map((f) => ({ kind: 'listField' as const, key: `listField:${def.listKey}/${f}` })));
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) {
      return;
    }
    const site = await loadSourceSite(sp, opts.signal);
    const found: string[] = [];

    for (const l of site.lists.filter((x) => wanted.some((k) => k.indexOf(`view:${x.key}/`) === 0))) {
      throwIfAborted(opts.signal);
      const listDef = out.manifest.lists.filter((x) => x.key === l.key)[0];
      if (!listDef) {
        opts.log.warn(`List ${l.info.Title} is not in the template; its views are skipped.`, { artifact: { kind: 'list', key: `list:${l.key}` }, code: 'VIEW_LIST_MISSING' });
        continue;
      }
      const listUrl = l.info.RootFolder.ServerRelativeUrl;
      for (const v of (await this._views(sp, listUrl)).filter((x) => wanted.indexOf(viewKey(l.key, x.Title)) >= 0)) {
        const ref: IArtifactRef = { kind: this.kind, key: viewKey(l.key, v.Title) };
        found.push(ref.key);
        const fields = await sp.web.getList(listUrl).views.getById(v.Id).fields<{ Items: string[] }>();
        const view = toViewDef(v, fields.Items || [], opts.tokens);
        if (view.viewType && view.viewType !== 'HTML') {
          opts.log.warn(`View ${v.Title} is a ${view.viewType} view; CopyJet installs only HTML views, so it will be skipped on install.`, {
            artifact: ref,
            code: 'VIEW_TYPE_UNSUPPORTED'
          });
        }
        const views = (listDef.views = listDef.views || []);
        const at = views.findIndex((x) => x.title === view.title);
        if (at >= 0) {
          views[at] = view;
        } else {
          views.push(view);
        }
      }
    }

    wanted
      .filter((key) => found.indexOf(key) < 0)
      .forEach((key) => opts.log.warn('View not found on the source site.', { artifact: { kind: this.kind, key }, code: 'VIEW_NOT_FOUND' }));
  }

  private _views(sp: SPFI, listServerRelativeUrl: string): Promise<IViewInfoLike[]> {
    return sp.web.getList(listServerRelativeUrl).views.filter(COPIED_VIEWS_FILTER).select(...VIEW_SELECT)<IViewInfoLike[]>();
  }
}
