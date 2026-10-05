import type { IPage } from '../model';

/** A page step's input: the manifest entry (the canvas is read from the package at install). */
export type IPageDef = IPage;

export const pageKey = (name: string): string => `page:${name}`;

/** pages/<name without .aspx>.json */
export const pageEntryPath = (name: string): string => `pages/${name.replace(/\.aspx$/i, '')}.json`;

/** Package path of a SiteAssets file: assets/SiteAssets/<path inside SiteAssets>. */
export const assetEntryPath = (siteAssetsPath: string): string => `assets/SiteAssets/${siteAssetsPath.replace(/\.{2,}/g, '.')}`;

/** The path inside SiteAssets of an asset entry (the reverse of assetEntryPath). */
export const siteAssetsPathOf = (entry: string): string => entry.replace(/^assets\/SiteAssets\//, '');

/** SharePoint's page layouts the schema knows; others are carried as Article. */
const LAYOUTS: IPage['layout'][] = ['Article', 'Home', 'SingleWebPartAppPage', 'RepostPage', 'Spaces'];
export const toLayout = (value: string | undefined): IPage['layout'] => (LAYOUTS.indexOf(value as IPage['layout']) >= 0 ? (value as IPage['layout']) : 'Article');

const escapeRegExp = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Paths inside the web's SiteAssets library that the texts refer to (images of the canvas, the header, the
 * banner), decoded and unique: "/sites/a/SiteAssets/SitePages/P/k%C3%A9p.jpg" → "SitePages/P/kép.jpg".
 */
export function siteAssetsIn(texts: Array<string | undefined>, webServerRelativeUrl: string): string[] {
  const prefix = `${webServerRelativeUrl.replace(/\/$/, '')}/SiteAssets/`;
  // eslint-disable-next-line @rushstack/security/no-unsafe-regexp -- built from an escaped web URL
  const re = new RegExp(`${escapeRegExp(prefix).replace(/ /g, '(?: |%20)')}([^"'\\s\\\\?#<>]+)`, 'gi');
  const out: string[] = [];
  texts.forEach((text) => {
    if (!text) return;
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text))) {
      let path = m[1];
      try {
        path = decodeURIComponent(path);
      } catch {
        // not encoded
      }
      if (out.indexOf(path) < 0) out.push(path);
    }
  });
  return out;
}

export interface ICanvasWebPart {
  id: string;
  title: string;
}

/** Web parts on a canvas (controlType 3), in order, each once. */
export function webPartsOf(canvas: unknown[]): ICanvasWebPart[] {
  const out: ICanvasWebPart[] = [];
  canvas.forEach((c) => {
    const control = c as { controlType?: number; webPartId?: string; webPartData?: { id?: string; title?: string } };
    const id = (control.webPartId || (control.webPartData && control.webPartData.id) || '').toLowerCase();
    if (control.controlType === 3 && id && !out.some((w) => w.id === id)) out.push({ id, title: (control.webPartData && control.webPartData.title) || id });
  });
  return out;
}
