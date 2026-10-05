import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/sites';
import '@pnp/sp/lists';
import '@pnp/sp/views';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { throwIfAborted } from '../errors';
import { loadSourceSite } from '../lists';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, IPage, IPageFile, ITemplateWriter } from '../model';
import {
  assetEntryPath,
  clientSideWebParts,
  isInternalWebPart,
  listSitePages,
  pageEntryPath,
  pageKey,
  readSitePage,
  siteAssetsIn,
  toLayout,
  webPartsOf,
  type IPageDef
} from '../pages';
import { TokenContext, tokenize } from '../tokenizer';

/**
 * Modern pages (spike 12): the canvas and header as the sitepages API gives them, tokenized for the web, the
 * template's lists and views; SiteAssets images referenced by the page go into the package. Pages come after
 * lists and views in the registry, so the template's lists and views are known here.
 */
export class PageExtractor implements IExtractor<IPageDef> {
  public readonly kind = 'page' as const;

  public async discover(sp: SPFI): Promise<IDiscoveredArtifact[]> {
    return (await listSitePages(sp)).map((p) => ({ ref: { kind: this.kind, key: pageKey(p.FileName) }, title: p.Title || p.FileName, group: p.PageLayoutType }));
  }

  public dependencies(): IArtifactRef[] {
    return [];
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) return;
    const [web, root, pages, site, parts, siteInfo] = await Promise.all([
      sp.web.select('Id', 'Url', 'ServerRelativeUrl', 'Title')<{ Id: string; Url: string; ServerRelativeUrl: string; Title: string }>(),
      sp.web.rootFolder.select('WelcomePage')<{ WelcomePage?: string }>(),
      listSitePages(sp),
      loadSourceSite(sp, opts.signal),
      clientSideWebParts(sp),
      sp.site.select('Id')<{ Id?: string }>().catch(() => ({ Id: undefined as string | undefined }))
    ]);
    const tokens = await this._tokens(sp, web, site, out);
    if (siteInfo.Id) tokens.set('siteid', undefined, siteInfo.Id);
    const inTemplate = (key: string): boolean => out.manifest.lists.some((l) => l.key === key);
    const outside = site.lists.filter((l) => !inTemplate(l.key));

    for (const page of pages.filter((p) => wanted.indexOf(pageKey(p.FileName)) >= 0)) {
      throwIfAborted(opts.signal);
      const ref: IArtifactRef = { kind: this.kind, key: pageKey(page.FileName) };
      const content = await readSitePage(sp, page.Id);
      const canvasText = content.CanvasContent1 || '[]';
      let canvas: unknown[];
      try {
        canvas = JSON.parse(tokenize(canvasText, tokens)) as unknown[];
      } catch (e) {
        opts.log.warn(`The canvas of ${page.FileName} is not valid JSON; the page is skipped.`, { artifact: ref, code: 'PAGE_CANVAS_INVALID', detail: e instanceof Error ? e.message : e });
        continue;
      }
      outside
        .filter((l) => canvasText.toLowerCase().indexOf(l.info.Id.replace(/[{}]/g, '').toLowerCase()) >= 0)
        .forEach((l) =>
          opts.log.warn(`${page.FileName} shows list ${l.info.Title}, which is not in the template; that web part will point to the source list.`, {
            artifact: ref,
            code: 'PAGE_LIST_OUTSIDE_TEMPLATE',
            detail: l.key
          })
        );

      // SiteAssets files of the canvas, the header and the banner.
      const assets = siteAssetsIn([canvasText, content.LayoutWebpartsContent, content.BannerImageUrl], web.ServerRelativeUrl);
      const packed: string[] = [];
      for (const path of assets) {
        try {
          out.addBlob(assetEntryPath(path), await sp.web.getFileByServerRelativePath(`${web.ServerRelativeUrl.replace(/\/$/, '')}/SiteAssets/${path}`).getBlob());
          packed.push(assetEntryPath(path));
        } catch {
          opts.log.warn(`Image SiteAssets/${path} of ${page.FileName} could not be read.`, { artifact: ref, code: 'PAGE_ASSET_MISSING', detail: path });
        }
      }
      const file: IPageFile = { name: page.FileName, canvasContent: canvas };
      if (content.LayoutWebpartsContent) file.layoutWebpartsContent = tokenize(content.LayoutWebpartsContent, tokens);
      if (content.Description) file.description = content.Description;
      if (content.TopicHeader) file.topicHeader = content.TopicHeader;
      if (content.BannerImageUrl) {
        const banner = siteAssetsIn([content.BannerImageUrl], web.ServerRelativeUrl)[0];
        if (banner && packed.indexOf(assetEntryPath(banner)) >= 0) {
          file.bannerImage = assetEntryPath(banner);
        } else if (!/\/_layouts\//i.test(content.BannerImageUrl)) {
          opts.log.warn(`The header image of ${page.FileName} is outside SiteAssets; the page is installed without it.`, { artifact: ref, code: 'PAGE_BANNER_EXTERNAL' });
        }
      }
      out.addJson(pageEntryPath(page.FileName), file);

      const def: IPage = { name: page.FileName, title: content.Title || page.Title, layout: toLayout(content.PageLayoutType || page.PageLayoutType), source: pageEntryPath(page.FileName) };
      const promoted = content.PromotedState !== undefined ? content.PromotedState : page.PromotedState;
      if (promoted === 0 || promoted === 1 || promoted === 2) def.promotedState = promoted;
      if ((root.WelcomePage || '').toLowerCase() === `sitepages/${page.FileName.toLowerCase()}`) def.isHomePage = true;
      if (packed.length) def.assets = packed;
      const used = webPartsOf(canvas);
      if (used.length) {
        def.requiredWebParts = used.map((w) => {
          const known = parts.filter((p) => (p.Id || '').toLowerCase() === w.id)[0];
          return { id: w.id, title: w.title, isCustom: known ? !isInternalWebPart(known) : true };
        });
      }
      const at = out.manifest.pages.findIndex((p) => p.name.toLowerCase() === def.name.toLowerCase());
      if (at >= 0) out.manifest.pages[at] = def;
      else out.manifest.pages.push(def);
      out.manifest.meta.includesContent = true;
      opts.log.info(`Page extracted: ${used.length} web parts, ${packed.length} images.`, { artifact: ref });
    }

    wanted
      .filter((key) => !pages.some((p) => pageKey(p.FileName) === key))
      .forEach((key) => opts.log.warn('Page not found on the source site.', { artifact: { kind: this.kind, key }, code: 'PAGE_NOT_FOUND' }));
  }

  /**
   * Source values the canvas is tokenized with: the site, the web ID, and only the template's lists and their
   * views (a list outside the template keeps its source ID; the extractor warns about it).
   */
  private async _tokens(
    sp: SPFI,
    web: { Id: string; Url: string; ServerRelativeUrl: string; Title: string },
    site: Awaited<ReturnType<typeof loadSourceSite>>,
    out: ITemplateWriter
  ): Promise<TokenContext> {
    const tokens = TokenContext.forSite({ absoluteUrl: web.Url, serverRelativeUrl: web.ServerRelativeUrl, title: web.Title });
    tokens.set('webid', undefined, web.Id);
    for (const list of out.manifest.lists) {
      const source = site.lists.filter((l) => l.key === list.key)[0];
      if (!source) continue;
      tokens.set('listkey', list.key, source.info.Id);
      tokens.set('listurl', list.key, source.url);
      const titles = (list.views || []).map((v) => v.title);
      if (!titles.length) continue;
      const views = await sp.web.getList(source.info.RootFolder.ServerRelativeUrl).views.select('Id', 'Title')<Array<{ Id: string; Title: string }>>();
      views
        .filter((v) => titles.indexOf(v.Title) >= 0 && /^[^{}:]+$/.test(`${list.key}/${v.Title}`))
        .forEach((v) => tokens.set('viewid', `${list.key}/${v.Title}`, v.Id));
    }
    return tokens;
  }
}
