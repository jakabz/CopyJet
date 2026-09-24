import type { IList } from '../model';

/** List properties CopyJet reads from the REST API. */
export interface IListInfoLike {
  Id: string;
  Title: string;
  Description?: string;
  BaseTemplate: number;
  Hidden?: boolean;
  IsCatalog?: boolean;
  IsSystemList?: boolean;
  ItemCount?: number;
  OnQuickLaunch?: boolean;
  EnableVersioning?: boolean;
  EnableMinorVersions?: boolean;
  MajorVersionLimit?: number;
  EnableModeration?: boolean;
  EnableAttachments?: boolean;
  EnableFolderCreation?: boolean;
  ForceCheckout?: boolean;
  ContentTypesEnabled?: boolean;
  RootFolder: { ServerRelativeUrl: string };
}

export const LIST_SELECT = [
  'Id', 'Title', 'Description', 'BaseTemplate', 'Hidden', 'IsCatalog', 'IsSystemList', 'ItemCount', 'OnQuickLaunch',
  'EnableVersioning', 'EnableMinorVersions', 'MajorVersionLimit', 'EnableModeration', 'EnableAttachments',
  'EnableFolderCreation', 'ForceCheckout', 'ContentTypesEnabled', 'RootFolder/ServerRelativeUrl'
];

export const CUSTOM_LIST = 100;
export const DOCUMENT_LIBRARY = 101;
/** Templates the list pair supports in v1; Site Pages (119) is a system list handled by the page pair. */
export const SUPPORTED_TEMPLATES = [CUSTOM_LIST, DOCUMENT_LIBRARY];

/**
 * Lists a user created or can meaningfully copy: not hidden, not a system list or catalog (spike 03: this
 * leaves exactly the user's lists; Site Pages, Site Assets, Style Library, Form Templates are system lists).
 */
export function isUserList(info: IListInfoLike): boolean {
  return !info.Hidden && !info.IsSystemList && !info.IsCatalog;
}

export function isSupportedTemplate(template: number): boolean {
  return SUPPORTED_TEMPLATES.indexOf(template) >= 0;
}

/** '/sites/a/Lists/X' relative to web '/sites/a' → 'Lists/X'. */
export function toSiteRelativeUrl(serverRelativeUrl: string, webServerRelativeUrl: string): string {
  const web = webServerRelativeUrl.replace(/\/+$/, '');
  const url = serverRelativeUrl.replace(/\/+$/, '');
  return (url.toLowerCase().indexOf(web.toLowerCase() + '/') === 0 ? url.slice(web.length + 1) : url.replace(/^\/+/, '')) || url;
}

/** 'Lists/X' on web '/sites/a' → '/sites/a/Lists/X' (root web '/' → '/Lists/X'). */
export function toServerRelativeUrl(siteRelativeUrl: string, webServerRelativeUrl: string): string {
  return `${webServerRelativeUrl.replace(/\/+$/, '')}/${siteRelativeUrl.replace(/^\/+/, '')}`;
}

/** Last URL segment: the list's name in its URL ('Lists/Teszt lista' → 'Teszt lista'). */
export function urlLeaf(siteRelativeUrl: string): string {
  const parts = siteRelativeUrl.replace(/\/+$/, '').split('/');
  return parts[parts.length - 1];
}

/**
 * Template key from the URL leaf, matching the schema's key pattern: 'Shared Documents' → 'Shared_Documents',
 * 'Ügyfelek' → 'Ugyfelek' (accents dropped, not the letters).
 */
export function listKeyFromUrl(siteRelativeUrl: string): string {
  const key = urlLeaf(siteRelativeUrl)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_.-]/g, '')
    .slice(0, 120);
  return key || 'list';
}

/** Keys for all given list URLs, unique within the site (a clash gets _2, _3 …). Order-independent. */
export function assignListKeys(siteRelativeUrls: string[]): { [url: string]: string } {
  const out: { [url: string]: string } = {};
  const used: { [key: string]: boolean } = {};
  siteRelativeUrls
    .slice()
    .sort()
    .forEach((url) => {
      const base = listKeyFromUrl(url);
      let key = base;
      for (let n = 2; used[key.toLowerCase()]; n++) key = `${base}_${n}`;
      used[key.toLowerCase()] = true;
      out[url] = key;
    });
  return out;
}

export function toListDef(info: IListInfoLike, key: string, webServerRelativeUrl: string): IList {
  const isLibrary = info.BaseTemplate === DOCUMENT_LIBRARY;
  const def: IList = {
    key,
    url: toSiteRelativeUrl(info.RootFolder.ServerRelativeUrl, webServerRelativeUrl),
    title: info.Title,
    template: info.BaseTemplate as IList['template'],
    content: { mode: 'none' }
  };
  if (info.Description) def.description = info.Description;
  def.onQuickLaunch = !!info.OnQuickLaunch;
  def.enableVersioning = !!info.EnableVersioning;
  if (info.EnableVersioning && info.MajorVersionLimit) def.majorVersionLimit = info.MajorVersionLimit;
  def.enableModeration = !!info.EnableModeration;
  def.enableFolderCreation = !!info.EnableFolderCreation;
  def.contentTypesEnabled = !!info.ContentTypesEnabled;
  if (isLibrary) {
    def.enableMinorVersions = !!info.EnableMinorVersions;
    def.forceCheckout = !!info.ForceCheckout;
  } else {
    def.enableAttachments = !!info.EnableAttachments;
  }
  return def;
}

/** Template property → REST property, for the settings CopyJet copies. */
const SETTINGS: Array<[keyof IList, keyof IListInfoLike]> = [
  ['title', 'Title'],
  ['description', 'Description'],
  ['onQuickLaunch', 'OnQuickLaunch'],
  ['enableVersioning', 'EnableVersioning'],
  ['enableMinorVersions', 'EnableMinorVersions'],
  ['majorVersionLimit', 'MajorVersionLimit'],
  ['enableModeration', 'EnableModeration'],
  ['enableAttachments', 'EnableAttachments'],
  ['enableFolderCreation', 'EnableFolderCreation'],
  ['forceCheckout', 'ForceCheckout'],
  ['contentTypesEnabled', 'ContentTypesEnabled']
];

const blank = (v: unknown): unknown => (v === undefined || v === null || v === '' ? '' : v);

/** Template value vs REST value; a missing REST flag counts as false. */
const same = (templateValue: unknown, targetValue: unknown): boolean =>
  typeof templateValue === 'boolean' ? templateValue === !!targetValue : blank(templateValue) === blank(targetValue);

/** Template properties whose value differs on the target. Properties absent from the template are not compared. */
export function compareLists(def: IList, target: IListInfoLike): string[] {
  const changes = SETTINGS.filter(([p, r]) => def[p] !== undefined && !same(def[p], target[r])).map(([p]) => p as string);
  if (def.template !== target.BaseTemplate) changes.unshift('template');
  return changes;
}

/**
 * REST properties to MERGE for the given template properties. MajorVersionLimit only goes with versioning
 * on (SharePoint rejects it otherwise).
 */
export function listUpdateProps(def: IList, only?: string[]): { [prop: string]: unknown } {
  const props: { [prop: string]: unknown } = {};
  SETTINGS.forEach(([p, r]) => {
    if (def[p] === undefined || (only && only.indexOf(p as string) < 0)) return;
    if (p === 'majorVersionLimit' && !def.enableVersioning) return;
    props[r] = def[p];
  });
  return props;
}
