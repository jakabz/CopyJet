import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import { throwIfAborted } from '../errors';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, INavNode, ITemplateWriter } from '../model';
import {
  countNodes,
  hostOf,
  isSystemNode,
  LINKLESS,
  MAX_DEPTH,
  navKeys,
  readMenuState,
  readWelcomePage,
  tokenizeNavUrl,
  type IMenuStateNode,
  type INavStepDef,
  type INavUrlSource,
  type NavMenu
} from '../navigation';

const MENUS: NavMenu[] = ['quickLaunch', 'topNavigation'];
const HOME_PREFIX = /^sitepages\//i;

/**
 * The left (QuickLaunch) and top menus and the home page (spike 13). Menus come from MenuState, without the
 * built-in nodes; links into the web and the template's lists are tokenized. It runs after the lists and pages
 * extractors, so the template's lists and pages are known here.
 */
export class NavigationExtractor implements IExtractor<INavStepDef> {
  public readonly kind = 'navigation' as const;

  public async discover(sp: SPFI): Promise<IDiscoveredArtifact[]> {
    const out: IDiscoveredArtifact[] = [];
    for (const menu of MENUS) {
      const count = userNodeCount(await readMenuState(sp, menu));
      if (count > 0) out.push({ ref: { kind: this.kind, key: navKeys[menu] }, title: menu, itemCount: count });
    }
    const home = await readWelcomePage(sp);
    if (HOME_PREFIX.test(home)) out.push({ ref: { kind: this.kind, key: navKeys.homePage }, title: home.replace(HOME_PREFIX, '') });
    return out;
  }

  public dependencies(): IArtifactRef[] {
    return [];
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) return;
    const web = await sp.web.select('Url', 'ServerRelativeUrl')<{ Url: string; ServerRelativeUrl: string }>();
    const src: INavUrlSource = { webUrl: web.Url, webServerRelativeUrl: web.ServerRelativeUrl, lists: out.manifest.lists.map((l) => ({ key: l.key, url: l.url })) };
    const nav = (out.manifest.navigation = out.manifest.navigation || { mode: 'append' });

    for (const menu of MENUS.filter((m) => wanted.indexOf(navKeys[m]) >= 0)) {
      throwIfAborted(opts.signal);
      const ref: IArtifactRef = { kind: this.kind, key: navKeys[menu] };
      const nodes = this._convert(await readMenuState(sp, menu), 1, src, hostOf(web.Url), ref, opts);
      nav[menu] = nodes;
      opts.log.info(`Menu extracted: ${countNodes(nodes)} links.`, { artifact: ref });
    }

    if (wanted.indexOf(navKeys.homePage) >= 0) {
      const ref: IArtifactRef = { kind: this.kind, key: navKeys.homePage };
      const home = await readWelcomePage(sp);
      if (!HOME_PREFIX.test(home)) {
        opts.log.warn(`The home page ${home || '(none)'} is not a page of Site Pages; it is not carried.`, { artifact: ref, code: 'NAV_HOME_PAGE_UNSUPPORTED', detail: home });
      } else {
        const name = home.replace(HOME_PREFIX, '');
        nav.homePage = name;
        if (!out.manifest.pages.some((p) => p.name.toLowerCase() === name.toLowerCase())) {
          opts.log.warn(`The home page ${name} is not in the template; the target gets it as home page only if it has that page.`, {
            artifact: ref,
            code: 'NAV_HOME_PAGE_NOT_IN_TEMPLATE',
            detail: name
          });
        }
        opts.log.info(`Home page: ${name}.`, { artifact: ref });
      }
    }
  }

  private _convert(nodes: IMenuStateNode[], depth: number, src: INavUrlSource, sourceHost: string, ref: IArtifactRef, opts: IExtractOptions): INavNode[] {
    const out: INavNode[] = [];
    nodes.forEach((n) => {
      if (n.IsDeleted || isSystemNode(n.Key)) return;
      if (n.IsHidden) {
        opts.log.info(`Hidden menu node ${n.Title} is not carried.`, { artifact: ref, code: 'NAV_HIDDEN_SKIPPED', detail: n.Title });
        return;
      }
      const raw = n.SimpleUrl || '';
      const url = raw ? tokenizeNavUrl(raw, src) : LINKLESS;
      const node: INavNode = { title: n.Title, url };
      if (url === LINKLESS || (/^https?:\/\//i.test(url) && hostOf(url) !== sourceHost)) node.isExternal = true;
      if (n.AudienceIds && n.AudienceIds.length) node.audiences = n.AudienceIds.map((a) => a.replace(/[{}]/g, '').toLowerCase());
      const children = n.Nodes || [];
      if (children.length) {
        if (depth >= MAX_DEPTH) {
          opts.log.warn(`Menu levels below ${n.Title} are deeper than ${MAX_DEPTH} and are not carried.`, { artifact: ref, code: 'NAV_TOO_DEEP', detail: n.Title });
        } else {
          const kids = this._convert(children, depth + 1, src, sourceHost, ref, opts);
          if (kids.length) node.children = kids;
        }
      }
      out.push(node);
    });
    return out;
  }
}

function userNodeCount(nodes: IMenuStateNode[]): number {
  return nodes.filter((n) => !n.IsDeleted && !n.IsHidden && !isSystemNode(n.Key)).reduce((s, n) => s + 1 + userNodeCount(n.Nodes || []), 0);
}

