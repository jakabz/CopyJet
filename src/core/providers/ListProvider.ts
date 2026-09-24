import { spPost, type SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/folders';
import { ContentTypes } from '@pnp/sp/content-types';
import { normalizeContentTypeId } from '../contentTypes';
import { CopyJetError, throwIfAborted } from '../errors';
import { setContentTypeOrder, type FetchLike } from '../http/raw';
import { isHttpStatus } from '../http/status';
import {
  CUSTOM_LIST,
  DOCUMENT_LIBRARY,
  LIST_SELECT,
  compareLists,
  isSupportedTemplate,
  listContentTypeFor,
  listFolderPaths,
  listUpdateProps,
  missingFolders,
  readListContentTypes,
  siteContentTypeIdOf,
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

/** Additive structure changes 'update' mode (and creation) apply; see _syncStructure. */
const STRUCTURE_CHANGES = ['contentTypes', 'contentTypeOrder', 'folders'];

export class ListProvider implements IProvider<IList> {
  public readonly kind = 'list' as const;
  private readonly _fetch?: FetchLike;

  /** `fetchImpl` is used for the raw REST call that sets the content type order (injected in tests). */
  constructor(fetchImpl?: FetchLike) {
    this._fetch = fetchImpl;
  }

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
    changes.push(...(await this._structureChanges(sp, def, target.RootFolder.ServerRelativeUrl)));
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
          const settings = (diff.changes || []).filter((c) => STRUCTURE_CHANGES.indexOf(c) < 0);
          if (settings.length) {
            await sp.web.getList(this._serverUrl(def.url, ctx)).update(listUpdateProps(def, settings));
          }
          await this._syncStructure(sp, def, diff.target!.RootFolder.ServerRelativeUrl, ctx);
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
    await this._syncStructure(sp, def, created.RootFolder.ServerRelativeUrl, ctx);
    ctx.log.info(url === def.url ? 'List created.' : `List created as a renamed copy at ${url}.`, { artifact: ref });
    return this._done(ref, 'created', { ...def, url }, created, ctx);
  }

  /** Content types missing from the list, a different order (first = default), missing folders. */
  private async _structureChanges(sp: SPFI, def: IList, listUrl: string): Promise<string[]> {
    const changes: string[] = [];
    if (def.contentTypes && def.contentTypes.length) {
      const cts = await readListContentTypes(sp, listUrl);
      if (def.contentTypes.some((id) => !listContentTypeFor(id, cts.all))) {
        changes.push('contentTypes');
      } else if (!this._orderMatches(def.contentTypes, cts.ordered)) {
        changes.push('contentTypeOrder');
      }
    }
    if (def.folders && def.folders.length) {
      const existing = await listFolderPaths(sp, listUrl);
      if (missingFolders(def.folders.map((f) => f.path), existing).length) {
        changes.push('folders');
      }
    }
    return changes;
  }

  /** The template's content types lead the target's visible order, in the same sequence. */
  private _orderMatches(siteIds: string[], orderedListIds: string[]): boolean {
    return siteIds.every((id, i) => orderedListIds[i] !== undefined && siteContentTypeIdOf(orderedListIds[i]) === normalizeContentTypeId(id));
  }

  /**
   * Adds missing content types, puts the template's ones first (first = default), creates missing folders
   * parent-first. Never removes a content type or folder (spike 04).
   */
  private async _syncStructure(sp: SPFI, def: IList, listUrl: string, ctx: IInstallContext): Promise<void> {
    const ref = { kind: this.kind, key: `list:${def.key}` };
    if (def.contentTypes && def.contentTypes.length) {
      const list = sp.web.getList(listUrl);
      const before = await readListContentTypes(sp, listUrl);
      for (const siteId of def.contentTypes) {
        throwIfAborted(ctx.signal);
        if (!listContentTypeFor(siteId, before.all)) {
          try {
            // URL-parameter form, as verified in spike 04 (PnPjs' helper posts the ID in a body instead).
            await spPost(ContentTypes(list.contentTypes, `addAvailableContentType('${normalizeContentTypeId(siteId)}')`));
          } catch (e) {
            throw new CopyJetError('LIST_CT_FAILED', `Could not add content type ${siteId} to list ${def.url}; is it installed on the site?`, e);
          }
        }
      }
      const cts = await readListContentTypes(sp, listUrl);
      if (!this._orderMatches(def.contentTypes, cts.ordered)) {
        const lead = def.contentTypes.map((id) => listContentTypeFor(id, cts.all)!).filter((id) => !!id);
        const rest = cts.ordered.filter((id) => lead.indexOf(id) < 0);
        await setContentTypeOrder(sp, listUrl, lead.concat(rest), ctx.signal, this._fetch);
      }
    }
    if (def.folders && def.folders.length) {
      const create = missingFolders(def.folders.map((f) => f.path), await listFolderPaths(sp, listUrl, ctx.signal));
      for (const path of create) {
        throwIfAborted(ctx.signal);
        await sp.web.folders.addUsingPath(`${listUrl}/${path}`);
      }
      if (create.length) {
        ctx.log.info(`Folders created: ${create.length}.`, { artifact: ref });
      }
    }
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
