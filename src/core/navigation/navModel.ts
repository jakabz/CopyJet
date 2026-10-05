import type { INavNode } from '../model';
import type { TokenContext } from '../tokenizer';
import { resolveKnown } from '../tokenizer';

/** The menus CopyJet copies (hub navigation is out of scope). */
export type NavMenu = 'quickLaunch' | 'topNavigation';

/** A navigation step: one menu, or setting the home page. */
export type INavStepDef =
  | { part: NavMenu; nodes: INavNode[]; mode?: 'append' | 'replace'; sourceTenant: string }
  | { part: 'homePage'; page: string; inTemplate: boolean };

export const navKeys = {
  quickLaunch: 'navigation:quickLaunch',
  topNavigation: 'navigation:topNavigation',
  homePage: 'navigation:homePage'
};

/** MenuState starting node keys of the menus (spike 13 A). */
export const MENU_NODE_KEY: { [m in NavMenu]: string } = { quickLaunch: '1025', topNavigation: '1002' };
/** REST collection names of the menus. */
export const MENU_COLLECTION: { [m in NavMenu]: string } = { quickLaunch: 'QuickLaunch', topNavigation: 'TopNavigationBar' };

/** What the UI stores for a heading without a link (spike 13 A); carried unchanged. */
export const LINKLESS = 'http://linkless.header/';

/** Levels CopyJet copies (the modern UI edits three). */
export const MAX_DEPTH = 3;

/**
 * Built-in nodes (Home 1031, Recent 1033, Site contents 1034 …) have keys 1000–1999 and exist on every
 * site; Recent's children are filled by SharePoint. They are never copied.
 */
export function isSystemNode(key: string | number): boolean {
  const n = typeof key === 'number' ? key : parseInt(key, 10);
  return !isNaN(n) && n >= 1000 && n < 2000;
}

const lower = (s: string): string => s.toLowerCase();
const startsAtBoundary = (path: string, prefix: string): boolean => {
  const p = lower(path);
  const x = lower(prefix);
  return p === x || (p.indexOf(x) === 0 && (p.charAt(x.length) === '/' || x.charAt(x.length - 1) === '/'));
};

export interface INavUrlSource {
  /** Absolute URL of the source web. */
  webUrl: string;
  /** Server-relative URL of the source web ("/sites/Forras"). */
  webServerRelativeUrl: string;
  /** The template's lists: key and site-relative URL ("Lists/Teszt lista"). */
  lists: Array<{ key: string; url: string }>;
}

/**
 * A node URL for the template: links into the source web become {siterelative}/…, links into a list of the
 * template {siterelative}/{listurl:K}/… (a renamed copy is followed); external links stay as they are.
 */
export function tokenizeNavUrl(url: string, src: INavUrlSource): string {
  if (!url || url === LINKLESS) return url;
  const web = src.webServerRelativeUrl.replace(/\/+$/, '');
  const abs = src.webUrl.replace(/\/+$/, '');
  let path = url;
  if (/^https?:\/\//i.test(url)) {
    if (!startsAtBoundary(url, abs)) return url;
    path = `${web}${url.slice(abs.length)}`;
  }
  if (path.charAt(0) !== '/' || !startsAtBoundary(path, web || '/')) return url;
  const rest = path.slice(web.length); // '' or '/…'
  const lists = src.lists.slice().sort((a, b) => b.url.length - a.url.length);
  for (const l of lists) {
    const prefix = `/${l.url.replace(/^\/+|\/+$/g, '')}`;
    if (startsAtBoundary(rest, prefix)) return `{siterelative}/{listurl:${l.key}}${rest.slice(prefix.length)}`;
  }
  return `{siterelative}${rest}`;
}

/** A template URL on the target; "//x" from a root web's empty {siterelative} is collapsed. */
export function resolveNavUrl(url: string, tokens: TokenContext): string {
  return resolveKnown(url, tokens).replace(/^\/\/(?!\/)/, '/');
}

/** Comparable form of a node URL on the target web: server-relative, decoded, lowercase, no trailing slash. */
export function comparableUrl(url: string, targetWebUrl: string): string {
  if (!url || url === LINKLESS) return url;
  let u = url;
  const origin = (/^https?:\/\/[^/]+/i.exec(targetWebUrl) || [''])[0];
  if (origin && startsAtBoundary(u, origin)) u = u.slice(origin.length) || '/';
  try {
    u = decodeURI(u);
  } catch {
    // keep as it is
  }
  return lower(u).replace(/\/+$/, '') || '/';
}

/** Number of links in a node tree. */
export function countNodes(nodes: INavNode[] | undefined): number {
  return (nodes || []).reduce((n, x) => n + 1 + countNodes(x.children), 0);
}

/** Host of a site URL ("contoso.sharepoint.com"), as meta.sourceTenant stores it. */
export const hostOf = (url: string): string => lower(url.replace(/^https?:\/\//i, '').split('/')[0]);
