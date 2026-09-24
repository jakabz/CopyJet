import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/views';
import { CopyJetError, throwIfAborted } from '../errors';
import { isHttpStatus } from '../http/status';
import {
  COPIED_VIEWS_FILTER,
  VIEW_SELECT,
  compareViews,
  isCreatableViewType,
  matchView,
  scopeNumber,
  toServerRelativeUrl,
  viewKey,
  type IListViewDef,
  type IViewInfoLike
} from '../lists';
import type { ConflictMode, IApplyResult, IDiffResult, IInstallContext, IProvider, IView } from '../model';
import { resolve } from '../tokenizer';

interface IViewDiff extends IDiffResult {
  target?: IViewInfoLike;
}

export class ViewProvider implements IProvider<IListViewDef> {
  public readonly kind = 'view' as const;

  public async diff(sp: SPFI, def: IListViewDef, ctx: IInstallContext): Promise<IViewDiff> {
    throwIfAborted(ctx.signal);
    const ref = { kind: this.kind, key: viewKey(def.listKey, def.view.title) };
    const listUrl = this._listUrl(def, ctx);
    let views: IViewInfoLike[];
    try {
      views = await sp.web.getList(listUrl).views.filter(COPIED_VIEWS_FILTER).select(...VIEW_SELECT)<IViewInfoLike[]>();
    } catch (e) {
      if (isHttpStatus(e, 404)) return { ref, status: 'unsupported', changes: ['listMissing'] };
      throw e;
    }
    const view = this._resolved(def.view, ctx);
    const target = matchView(view, views);
    if (!target) {
      return isCreatableViewType(view.viewType) ? { ref, status: 'new' } : { ref, status: 'unsupported', changes: ['viewType'] };
    }
    const fields = await sp.web.getList(listUrl).views.getById(target.Id).fields<{ Items: string[] }>();
    const changes = compareViews(view, target, fields.Items || []);
    return { ref, status: changes.length ? 'different' : 'same', changes: changes.length ? changes : undefined, target };
  }

  public async apply(sp: SPFI, def: IListViewDef, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;

    switch (diff.status) {
      case 'new':
        return this._create(sp, def, ctx);
      case 'same':
        ctx.log.info('View already present; skipped.', { artifact: ref });
        return { ref, outcome: 'skipped' };
      case 'different':
        if (mode === 'update') {
          await this._configure(sp, def, diff.target!.Id, diff.changes || [], ctx);
          ctx.log.info(`View updated: ${(diff.changes || []).join(', ')}.`, { artifact: ref });
          return { ref, outcome: 'updated' };
        }
        if (mode === 'rename') {
          ctx.log.warn('Views cannot be installed renamed; the existing view is kept.', { artifact: ref, code: 'VIEW_RENAME_UNSUPPORTED' });
        }
        ctx.log.info(`View differs (${(diff.changes || []).join(', ')}); kept as it is.`, { artifact: ref });
        return { ref, outcome: 'skipped' };
      default:
        ctx.log.warn(`View skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'VIEW_UNSUPPORTED', detail: diff.changes });
        return { ref, outcome: 'skipped' };
    }
  }

  /** Created by title (SharePoint derives the same URL as on the source, spike 06), then configured. */
  private async _create(sp: SPFI, def: IListViewDef, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = { kind: this.kind, key: viewKey(def.listKey, def.view.title) };
    const view = this._resolved(def.view, ctx);
    const settings: { [k: string]: unknown } = { Paged: !!view.paged };
    if (view.query) settings.ViewQuery = view.query;
    if (view.rowLimit) settings.RowLimit = view.rowLimit;
    const created = await sp.web.getList(this._listUrl(def, ctx)).views.add(view.title, false, settings);
    if (!created.Id) {
      throw new CopyJetError('VIEW_CREATE_FAILED', `SharePoint returned no Id for view ${view.title}.`);
    }
    const changes = ['fields'];
    if (scopeNumber(view.scope) !== 0) changes.push('scope');
    if (view.customFormatter) changes.push('customFormatter');
    if (view.default) changes.push('default');
    await this._configure(sp, def, created.Id, changes, ctx);
    ctx.log.info('View created.', { artifact: ref });
    return { ref, outcome: 'created' };
  }

  /** Applies the given changes to a view: fields replaced in template order, the rest in one MERGE. */
  private async _configure(sp: SPFI, def: IListViewDef, viewId: string, changes: string[], ctx: IInstallContext): Promise<void> {
    const view = this._resolved(def.view, ctx);
    const target = sp.web.getList(this._listUrl(def, ctx)).views.getById(viewId);
    if (changes.indexOf('fields') >= 0) {
      await target.fields.removeAll();
      for (const f of view.fields) {
        throwIfAborted(ctx.signal);
        await target.fields.add(f);
      }
    }
    const props: { [k: string]: unknown } = {};
    if (changes.indexOf('title') >= 0) props.Title = view.title;
    if (changes.indexOf('query') >= 0) props.ViewQuery = view.query || '';
    if (changes.indexOf('rowLimit') >= 0) props.RowLimit = view.rowLimit;
    if (changes.indexOf('paged') >= 0) props.Paged = !!view.paged;
    if (changes.indexOf('scope') >= 0) props.Scope = scopeNumber(view.scope);
    if (changes.indexOf('customFormatter') >= 0) props.CustomFormatter = view.customFormatter || '';
    if (changes.indexOf('default') >= 0) props.DefaultView = true;
    if (Object.keys(props).length) {
      await target.update(props);
    }
  }

  private _resolved(view: IView, ctx: IInstallContext): IView {
    return {
      ...view,
      title: resolve(view.title, ctx.tokens),
      query: view.query === undefined ? undefined : resolve(view.query, ctx.tokens),
      customFormatter: view.customFormatter === undefined ? undefined : resolve(view.customFormatter, ctx.tokens)
    };
  }

  private _listUrl(def: IListViewDef, ctx: IInstallContext): string {
    const web = ctx.tokens.get('siterelative');
    if (web === undefined) {
      throw new CopyJetError('TOKEN_UNRESOLVED', 'The install context has no {siterelative} value.');
    }
    return toServerRelativeUrl(ctx.tokens.get('listurl', def.listKey) || def.listUrl, web);
  }
}
