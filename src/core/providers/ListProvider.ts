import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { CopyJetError, throwIfAborted } from '../errors';
import { isHttpStatus } from '../http/status';
import {
  CUSTOM_LIST,
  DOCUMENT_LIBRARY,
  LIST_SELECT,
  compareLists,
  isSupportedTemplate,
  listUpdateProps,
  toServerRelativeUrl,
  urlLeaf,
  type IListInfoLike
} from '../lists';
import type { ConflictMode, IApplyResult, IDiffResult, IInstallContext, IList, IProvider } from '../model';

interface IListDiff extends IDiffResult {
  target?: IListInfoLike;
}

const odataString = (s: string): string => s.replace(/'/g, "''");

/** Where SharePoint puts a list created by title (spike 03): custom lists under Lists/, libraries at the web root. */
function creatableUrl(template: number, leaf: string): string | undefined {
  if (template === CUSTOM_LIST) return `Lists/${leaf}`;
  if (template === DOCUMENT_LIBRARY) return leaf;
  return undefined;
}

export class ListProvider implements IProvider<IList> {
  public readonly kind = 'list' as const;

  public async diff(sp: SPFI, def: IList, ctx: IInstallContext): Promise<IListDiff> {
    throwIfAborted(ctx.signal);
    const ref = { kind: this.kind, key: `list:${def.key}` };
    if (!isSupportedTemplate(def.template)) {
      return { ref, status: 'unsupported', changes: ['template'] };
    }
    const target = await this._get(sp, def.url, ctx);
    throwIfAborted(ctx.signal);

    if (!target) {
      // List titles are unique per web: another list with this title blocks the rename after creation.
      if (await this._titleTaken(sp, def.title)) {
        return { ref, status: 'unsupported', changes: ['titleConflict'] };
      }
      // Only URLs SharePoint derives from a title can be created; refusing others here avoids an orphan list.
      if (creatableUrl(def.template, urlLeaf(def.url)) !== def.url) {
        return { ref, status: 'unsupported', changes: ['url'] };
      }
      return { ref, status: 'new' };
    }
    const changes = compareLists(def, target);
    if (changes.indexOf('template') >= 0) {
      return { ref, status: 'unsupported', changes: ['template'], target };
    }
    return { ref, status: changes.length ? 'different' : 'same', changes: changes.length ? changes : undefined, target };
  }

  public async apply(sp: SPFI, def: IList, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;

    switch (diff.status) {
      case 'new':
        return this._create(sp, def, def.url, def.title, ctx);
      case 'same':
        ctx.log.info('List already present; skipped.', { artifact: ref });
        return this._done(ref, 'skipped', def, diff.target!, ctx);
      case 'different':
        if (mode === 'update') {
          await sp.web.getList(this._serverUrl(def.url, ctx)).update(listUpdateProps(def, diff.changes));
          ctx.log.info(`List updated: ${(diff.changes || []).join(', ')}.`, { artifact: ref });
          return this._done(ref, 'updated', def, diff.target!, ctx);
        }
        if (mode === 'rename') {
          return this._createCopy(sp, def, ctx);
        }
        ctx.log.info(`List differs (${(diff.changes || []).join(', ')}); kept as it is.`, { artifact: ref });
        return this._done(ref, 'skipped', def, diff.target!, ctx);
      default:
        ctx.log.warn(`List skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'LIST_UNSUPPORTED', detail: diff.changes });
        return { ref, outcome: 'skipped' };
    }
  }

  /** 'rename' conflict mode: a separate copy at <url>_copy titled "<title> (copy)"; the key's tokens point to it. */
  private async _createCopy(sp: SPFI, def: IList, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = { kind: this.kind, key: `list:${def.key}` };
    const url = creatableUrl(def.template, `${urlLeaf(def.url)}_copy`)!;
    const title = `${def.title} (copy)`;
    const existing = await this._get(sp, url, ctx);
    if (existing) {
      ctx.log.info('Renamed copy already present; skipped.', { artifact: ref });
      return this._done(ref, 'skipped', { ...def, url }, existing, ctx);
    }
    if (await this._titleTaken(sp, title)) {
      ctx.log.warn(`Cannot create the renamed copy: title "${title}" is taken.`, { artifact: ref, code: 'LIST_UNSUPPORTED', detail: ['titleConflict'] });
      return { ref, outcome: 'skipped' };
    }
    return this._create(sp, def, url, title, ctx);
  }

  /**
   * Created with the URL name as title (so SharePoint derives the wanted URL), then renamed and configured in
   * one MERGE (spike 03). The resulting URL is verified.
   */
  private async _create(sp: SPFI, def: IList, url: string, title: string, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = { kind: this.kind, key: `list:${def.key}` };
    await sp.web.lists.add(urlLeaf(url), '', def.template, false);
    const created = await this._get(sp, url, ctx);
    if (!created) {
      throw new CopyJetError('LIST_URL_MISMATCH', `List ${title} was not created at ${url}.`, { url });
    }
    await sp.web.getList(this._serverUrl(url, ctx)).update(listUpdateProps({ ...def, title }));
    ctx.log.info(url === def.url ? 'List created.' : `List created as a renamed copy at ${url}.`, { artifact: ref });
    return this._done(ref, 'created', { ...def, url }, created, ctx);
  }

  private _serverUrl(siteRelativeUrl: string, ctx: IInstallContext): string {
    const web = ctx.tokens.get('siterelative');
    if (web === undefined) {
      throw new CopyJetError('TOKEN_UNRESOLVED', 'The install context has no {siterelative} value.');
    }
    return toServerRelativeUrl(siteRelativeUrl, web);
  }

  /** The list at a site-relative URL, or undefined (SharePoint answers 404, spike 03). */
  private async _get(sp: SPFI, siteRelativeUrl: string, ctx: IInstallContext): Promise<IListInfoLike | undefined> {
    try {
      return await sp.web.getList(this._serverUrl(siteRelativeUrl, ctx)).select(...LIST_SELECT).expand('RootFolder')<IListInfoLike>();
    } catch (e) {
      if (isHttpStatus(e, 404)) {
        return undefined;
      }
      throw e;
    }
  }

  private async _titleTaken(sp: SPFI, title: string): Promise<boolean> {
    const found = await sp.web.lists.filter(`Title eq '${odataString(title)}'`).select('Id')<Array<{ Id: string }>>();
    return found.length > 0;
  }

  private _done(ref: IApplyResult['ref'], outcome: IApplyResult['outcome'], def: IList, target: IListInfoLike, ctx: IInstallContext): IApplyResult {
    const id = target.Id.replace(/^\{|\}$/g, '').toLowerCase();
    ctx.tokens.set('listkey', def.key, id);
    ctx.tokens.set('listurl', def.key, def.url);
    return { ref, outcome, tokens: { listkey: { [def.key]: id }, listurl: { [def.key]: def.url } } };
  }
}
