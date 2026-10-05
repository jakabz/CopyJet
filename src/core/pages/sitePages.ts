import { SPQueryable, spPost, type SPFI } from '@pnp/sp';
import { body } from '@pnp/queryable';
import '@pnp/sp/webs';

/**
 * The modern page API (_api/sitepages, spike 12) on the web's behaviors. PnPjs' clientside-pages module
 * reorders and rewrites the canvas; CopyJet needs it carried as it is, so the REST calls are made directly.
 */

export interface ISitePageInfo {
  Id: number;
  FileName: string;
  Title: string;
  PageLayoutType?: string;
  PromotedState?: number;
}

export interface ISitePageContent extends ISitePageInfo {
  CanvasContent1?: string;
  LayoutWebpartsContent?: string;
  BannerImageUrl?: string;
  Description?: string;
  TopicHeader?: string;
}

export interface IClientSideWebPart {
  Id: string;
  Name?: string;
  Manifest?: string;
}

const webUrl = (sp: SPFI): string => sp.web.toUrl().replace(/\/_api\/web\/?$/i, '');
const q = (sp: SPFI, path: string): ReturnType<typeof SPQueryable> => SPQueryable([sp.web, `${webUrl(sp)}/_api/${path}`]);
const asArray = <T>(r: unknown): T[] => (Array.isArray(r) ? (r as T[]) : ((r as { value?: T[] }) || {}).value || []);

export async function listSitePages(sp: SPFI): Promise<ISitePageInfo[]> {
  return asArray<ISitePageInfo>(await q(sp, 'sitepages/pages?$select=Id,FileName,Title,PageLayoutType,PromotedState&$top=5000')());
}

export function readSitePage(sp: SPFI, id: number): Promise<ISitePageContent> {
  return q(sp, `sitepages/pages(${id})?$select=Id,FileName,Title,PageLayoutType,PromotedState,CanvasContent1,LayoutWebpartsContent,BannerImageUrl,Description,TopicHeader`)<ISitePageContent>();
}

/** A new empty page (named Page.aspx or similar until renamed). */
export async function createSitePage(sp: SPFI, layout: string, promotedState: number): Promise<ISitePageInfo> {
  return spPost<ISitePageInfo>(q(sp, 'sitepages/pages'), body({ PageLayoutType: layout, PromotedState: promotedState }));
}

export async function checkoutSitePage(sp: SPFI, id: number): Promise<void> {
  await spPost(q(sp, `sitepages/pages(${id})/checkoutpage`));
}

export async function saveSitePageDraft(sp: SPFI, id: number, content: { [k: string]: unknown }): Promise<void> {
  await spPost(q(sp, `sitepages/pages(${id})/savepageasdraft`), body(content));
}

export async function publishSitePage(sp: SPFI, id: number): Promise<void> {
  await spPost(q(sp, `sitepages/pages(${id})/publish`));
}

/**
 * Makes sure the web has its SiteAssets library. The result is not used: PnPjs' ensureSiteAssetsLibrary builds
 * the list from OData metadata a minimal response may lack; the library always lives at <web>/SiteAssets.
 */
export async function ensureSiteAssets(sp: SPFI): Promise<void> {
  await spPost(q(sp, 'web/lists/ensuresiteassetslibrary'));
}

export async function clientSideWebParts(sp: SPFI): Promise<IClientSideWebPart[]> {
  return asArray<IClientSideWebPart>(await q(sp, 'web/GetClientSideWebParts')());
}

/** Built-in web parts carry isInternal in their manifest. */
export function isInternalWebPart(part: IClientSideWebPart): boolean {
  try {
    return !!(part.Manifest && (JSON.parse(part.Manifest) as { isInternal?: boolean }).isInternal);
  } catch {
    return false;
  }
}
