import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { CopyJetError, throwIfAborted } from '../errors';
import { isHttpStatus } from '../http/status';
import { createFolder, existingFolderPaths, missingFolders } from '../lists';
import type { ConflictMode, IApplyResult, IArtifactRef, IDiffResult, IInstallContext, IPageFile, IProvider } from '../model';
import {
  checkoutSitePage,
  clientSideWebParts,
  createSitePage,
  ensureSiteAssets,
  listSitePages,
  pageKey,
  publishSitePage,
  readSitePage,
  saveSitePageDraft,
  siteAssetsPathOf,
  type IClientSideWebPart,
  type IPageDef,
  type ISitePageInfo
} from '../pages';
import { resolveKnown } from '../tokenizer';
import { contentContext } from './ItemProvider';

interface IPageDiff extends IDiffResult {
  target?: ISitePageInfo;
}

/** SharePoint's "a file with this name already exists" (spike 10 B2). */
const ALREADY_EXISTS = /-2130575257/;

/** Tokens still in the canvas after resolving: references the target cannot serve. */
const LEFTOVER = /\{(site|siterelative|siteid|webid|listkey|listurl|viewid|fieldid):?[^{}"]*\}/g;

/** Per run caches (the page list and the target's web parts), keyed by the install context. */
const pageLists = new WeakMap<object, Promise<ISitePageInfo[]>>();
const webPartLists = new WeakMap<object, Promise<IClientSideWebPart[]>>();

/**
 * Modern pages (spike 12 B): create → check out → save the resolved canvas as draft → rename to the
 * template's file name → publish, so the page ends at 1.0. The rename must follow the first draft save: that
 * save names a never-published page after its title, and a name taken by the page itself gave "(1)" (2026-10-05). SiteAssets images are uploaded first (never
 * overwritten). An existing page is replaced only in update mode; the target's home page is left alone
 * (setting it is a navigation decision).
 */
export class PageProvider implements IProvider<IPageDef> {
  public readonly kind = 'page' as const;

  public async diff(sp: SPFI, def: IPageDef, ctx: IInstallContext): Promise<IPageDiff> {
    throwIfAborted(ctx.signal);
    const ref: IArtifactRef = { kind: this.kind, key: pageKey(def.name) };
    if (!pageLists.has(ctx)) pageLists.set(ctx, listSitePages(sp));
    const target = (await pageLists.get(ctx)!).filter((p) => p.FileName.toLowerCase() === def.name.toLowerCase())[0];
    return target ? { ref, status: 'different', changes: ['pageExists'], target } : { ref, status: 'new' };
  }

  public async apply(sp: SPFI, def: IPageDef, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;
    if (diff.status === 'different') {
      if (mode === 'update') {
        await this._write(sp, def, diff.target!.Id, ref, ctx);
        ctx.log.info('Page replaced with the template\'s content.', { artifact: ref });
        return { ref, outcome: 'updated' };
      }
      if (mode === 'rename') {
        ctx.log.warn('Pages cannot be installed renamed; the existing page is kept.', { artifact: ref, code: 'PAGE_RENAME_UNSUPPORTED' });
      } else {
        ctx.log.info('Page already present; kept as it is.', { artifact: ref });
      }
      return { ref, outcome: 'skipped' };
    }
    const created = await createSitePage(sp, def.layout, def.promotedState || 0);
    if (!created || !created.Id) {
      throw new CopyJetError('PAGE_CREATE_FAILED', `SharePoint returned no page for ${def.name}.`);
    }
    await this._write(sp, def, created.Id, ref, ctx, def.name);
    const final = (await readSitePage(sp, created.Id)).FileName;
    if (final && final.toLowerCase() !== def.name.toLowerCase()) {
      ctx.log.warn(`The page was saved as ${final} instead of ${def.name}; a later run will not recognise it.`, { artifact: ref, code: 'PAGE_NAME_DIFFERS', detail: final });
    }
    ctx.log.info('Page created and published.', { artifact: ref });
    return { ref, outcome: 'created' };
  }

  /** `rename` names a new page after its first draft save, before it is published. */
  private async _write(sp: SPFI, def: IPageDef, id: number, ref: IArtifactRef, ctx: IInstallContext, rename?: string): Promise<void> {
    const content = contentContext(ctx);
    const file = await content.reader.getJson<IPageFile>(def.source, ctx.signal);
    await this._uploadAssets(sp, def.assets || [], ref, ctx);
    await this._checkWebParts(sp, def, ref, ctx);

    const canvas = resolveKnown(JSON.stringify(file.canvasContent), ctx.tokens);
    const layout =
      file.layoutWebpartsContent === undefined || file.layoutWebpartsContent === null
        ? undefined
        : resolveKnown(typeof file.layoutWebpartsContent === 'string' ? file.layoutWebpartsContent : JSON.stringify(file.layoutWebpartsContent), ctx.tokens);
    const leftover: string[] = (canvas + (layout || '')).match(LEFTOVER) || [];
    if (leftover.length) {
      ctx.log.warn(`Page references the target cannot serve: ${leftover.filter((t, i) => leftover.indexOf(t) === i).join(', ')}.`, {
        artifact: ref,
        code: 'PAGE_TOKEN_UNRESOLVED',
        detail: leftover
      });
    }
    const values: { [k: string]: unknown } = { Title: resolveKnown(def.title, ctx.tokens), CanvasContent1: canvas };
    if (layout !== undefined) values.LayoutWebpartsContent = layout;
    if (file.description) values.Description = file.description;
    if (file.topicHeader) values.TopicHeader = file.topicHeader;
    if (file.bannerImage) {
      const path = siteAssetsPathOf(file.bannerImage).split('/').map(encodeURIComponent).join('/');
      values.BannerImageUrl = `${ctx.targetSiteUrl.replace(/\/$/, '')}/SiteAssets/${path}`;
    }
    await checkoutSitePage(sp, id);
    await saveSitePageDraft(sp, id, values);
    if (rename) await this._rename(sp, id, rename, ctx);
    await publishSitePage(sp, id);
  }

  /** The new page gets the template's file name (FileLeafRef without .aspx). */
  private async _rename(sp: SPFI, id: number, name: string, ctx: IInstallContext): Promise<void> {
    const web = ctx.tokens.get('siterelative') || '';
    const result = await sp.web
      .getList(`${web.replace(/\/$/, '')}/SitePages`)
      .items.getById(id)
      .validateUpdateListItem([{ FieldName: 'FileLeafRef', FieldValue: name.replace(/\.aspx$/i, '') }]);
    const failed = (result || []).filter((v) => v.HasException);
    if (failed.length) {
      throw new CopyJetError('PAGE_RENAME_FAILED', `The new page could not be named ${name}: ${failed.map((v) => v.ErrorMessage).join('; ')}`, failed);
    }
  }

  /** The page's images into the target SiteAssets at the same paths; existing files are kept. */
  private async _uploadAssets(sp: SPFI, assets: string[], ref: IArtifactRef, ctx: IInstallContext): Promise<void> {
    if (!assets.length) return;
    const content = contentContext(ctx);
    await ensureSiteAssets(sp);
    const root = `${(ctx.tokens.get('siterelative') || '').replace(/\/$/, '')}/SiteAssets`;
    const paths = assets.map(siteAssetsPathOf);
    const folders = paths.map((p) => p.slice(0, Math.max(0, p.lastIndexOf('/')))).filter((f, i, all) => !!f && all.indexOf(f) === i);
    for (const folder of missingFolders(folders, await existingFolderPaths(sp, root, true, ctx.signal))) {
      await createFolder(sp, root, folder, true);
    }
    let added = 0;
    for (let i = 0; i < assets.length; i++) {
      throwIfAborted(ctx.signal);
      const path = paths[i];
      const at = path.lastIndexOf('/');
      try {
        await sp.web
          .getFolderByServerRelativePath(at < 0 ? root : `${root}/${path.slice(0, at)}`)
          .files.addUsingPath(path.slice(at + 1), await content.reader.getBlob(assets[i], ctx.signal), { Overwrite: false });
        added++;
      } catch (e) {
        if (!(isHttpStatus(e, 400) && ALREADY_EXISTS.test(e instanceof Error ? e.message : String(e)))) throw e;
      }
    }
    if (added) ctx.log.info(`Page images added to SiteAssets: ${added}.`, { artifact: ref });
  }

  /** Web parts the target site does not offer: the page still installs, those parts show an error. */
  private async _checkWebParts(sp: SPFI, def: IPageDef, ref: IArtifactRef, ctx: IInstallContext): Promise<void> {
    const required = def.requiredWebParts || [];
    if (!required.length) return;
    if (!webPartLists.has(ctx)) webPartLists.set(ctx, clientSideWebParts(sp));
    const available = (await webPartLists.get(ctx)!).map((p) => (p.Id || '').toLowerCase());
    required
      .filter((w) => available.indexOf(w.id.toLowerCase()) < 0)
      .forEach((w) =>
        ctx.log.warn(`Web part "${w.title}" is not available on the target site${w.isCustom ? ' (a custom SPFx web part: deploy its app first)' : ''}.`, {
          artifact: ref,
          code: 'PAGE_WEBPART_MISSING',
          detail: w
        })
      );
  }
}
