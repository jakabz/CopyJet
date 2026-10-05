import { SPQueryable, spPost, spPostMerge, type SPFI } from '@pnp/sp';
import { body } from '@pnp/queryable';
import '@pnp/sp/webs';
import '@pnp/sp/folders';
import { isHttpStatus } from '../http/status';
import { MENU_COLLECTION, MENU_NODE_KEY, type NavMenu } from './navModel';

/**
 * Navigation through REST (spike 13): the menus are read with MenuState (the only API returning audiences and
 * hidden nodes) and written node by node through web/navigation – SaveMenuState replaces the whole menu and
 * would delete the target's own nodes.
 */

export interface IMenuStateNode {
  Key: string;
  Title: string;
  SimpleUrl?: string;
  IsHidden?: boolean;
  IsDeleted?: boolean;
  AudienceIds?: string[];
  Nodes?: IMenuStateNode[];
}

/** A node of the target menu as REST gives it. */
export interface ITargetNavNode {
  Id: number;
  Title: string;
  Url: string;
  Children?: ITargetNavNode[];
}

const webUrl = (sp: SPFI): string => sp.web.toUrl().replace(/\/_api\/web\/?$/i, '');
const q = (sp: SPFI, path: string): ReturnType<typeof SPQueryable> => SPQueryable([sp.web, `${webUrl(sp)}/_api/${path}`]);
const asArray = <T>(r: unknown): T[] => (Array.isArray(r) ? (r as T[]) : ((r as { value?: T[] }) || {}).value || []);

export async function readMenuState(sp: SPFI, menu: NavMenu): Promise<IMenuStateNode[]> {
  const r = await q(sp, `navigation/MenuState?menuNodeKey='${MENU_NODE_KEY[menu]}'&mapProviderName='SPNavigationProvider'&depth=10`)<{ Nodes?: IMenuStateNode[] }>();
  return (r && r.Nodes) || [];
}

const NODE_FIELDS = ['Id', 'Title', 'Url'];

/** The target menu three levels deep. */
export async function readTargetMenu(sp: SPFI, menu: NavMenu): Promise<ITargetNavNode[]> {
  const select = NODE_FIELDS.concat(NODE_FIELDS.map((f) => `Children/${f}`), NODE_FIELDS.map((f) => `Children/Children/${f}`)).join(',');
  return asArray<ITargetNavNode>(await q(sp, `web/navigation/${MENU_COLLECTION[menu]}?$expand=Children,Children/Children&$select=${select}`)());
}

/** Adds a node at the top of a menu or under a node; the new node's ID. */
export async function addNavNode(sp: SPFI, menu: NavMenu, parentId: number | undefined, node: { Title: string; Url: string; IsExternal: boolean }): Promise<number> {
  const path = parentId === undefined ? `web/navigation/${MENU_COLLECTION[menu]}` : `web/navigation/GetNodeById(${parentId})/Children`;
  const created = await spPost<{ Id: number }>(q(sp, path), body(node));
  return created.Id;
}

export async function setNavNodeAudiences(sp: SPFI, id: number, audiences: string[]): Promise<void> {
  await spPostMerge(q(sp, `web/navigation/GetNodeById(${id})`), body({ AudienceIds: audiences }));
}

/** SharePoint refuses an internal link whose target does not exist (spike 13 B/C), also with IsExternal. */
export function isMissingLinkTarget(e: unknown): boolean {
  return isHttpStatus(e, 500) && /-2130247147/.test(e instanceof Error ? e.message : String(e));
}

/** The web's home page, site-relative ("SitePages/Home.aspx"). */
export async function readWelcomePage(sp: SPFI): Promise<string> {
  return (await sp.web.rootFolder.select('WelcomePage')<{ WelcomePage?: string }>()).WelcomePage || '';
}

/** Not validated by SharePoint (spike 13 B): the caller checks that the page exists. */
export async function setWelcomePage(sp: SPFI, page: string): Promise<void> {
  await spPostMerge(q(sp, 'web/rootfolder'), body({ WelcomePage: page }));
}
