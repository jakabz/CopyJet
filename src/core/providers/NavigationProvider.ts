import type { SPFI } from '@pnp/sp';
import { throwIfAborted } from '../errors';
import type { ConflictMode, IApplyResult, IArtifactRef, IDiffResult, IInstallContext, INavNode, IProvider } from '../model';
import {
  addNavNode,
  comparableUrl,
  countNodes,
  hostOf,
  isMissingLinkTarget,
  LINKLESS,
  navKeys,
  readTargetMenu,
  readWelcomePage,
  resolveNavUrl,
  setNavNodeAudiences,
  setWelcomePage,
  type INavStepDef,
  type ITargetNavNode,
  type NavMenu
} from '../navigation';
import { listSitePages } from '../pages';

const UNRESOLVED = /\{[a-z]+(?::[^{}]*)?\}/;

interface IMenuCounts {
  added: number;
  present: number;
  broken: string[];
}

/**
 * Menus and the home page (spike 13). Menus are only appended to – a node with the same title and URL on the
 * same level counts as present, so a rerun adds nothing and nothing of the target is deleted. An internal link
 * whose target is missing is skipped (a heading with children is kept as a heading without link). The home
 * page is set when the page exists on the target; SharePoint itself does not check that.
 */
export class NavigationProvider implements IProvider<INavStepDef> {
  public readonly kind = 'navigation' as const;

  public async diff(sp: SPFI, def: INavStepDef, ctx: IInstallContext): Promise<IDiffResult> {
    throwIfAborted(ctx.signal);
    const ref: IArtifactRef = { kind: this.kind, key: navKeys[def.part] };
    if (def.part === 'homePage') {
      const current = (await readWelcomePage(sp)).toLowerCase();
      if (current === `sitepages/${def.page.toLowerCase()}`) return { ref, status: 'same' };
      if (!def.inTemplate && !(await this._pageExists(sp, def.page))) return { ref, status: 'unsupported', changes: ['homePageMissing'] };
      return { ref, status: 'new', changes: ['homePageDiffers'] };
    }
    const missing = this._missing(def.nodes, await readTargetMenu(sp, def.part), ctx);
    return missing > 0 ? { ref, status: 'new', changes: [`newLinks:${missing}`] } : { ref, status: 'same' };
  }

  public async apply(sp: SPFI, def: INavStepDef, _mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    throwIfAborted(ctx.signal);
    const ref: IArtifactRef = { kind: this.kind, key: navKeys[def.part] };
    if (def.part === 'homePage') return this._home(sp, def.page, ref, ctx);

    if (def.mode === 'replace') {
      ctx.log.warn('The template asks to replace the menu; CopyJet never deletes on the target, so the links are appended.', { artifact: ref, code: 'NAV_REPLACE_AS_APPEND' });
    }
    const sameTenant = hostOf(ctx.targetSiteUrl) === def.sourceTenant.toLowerCase();
    const counts: IMenuCounts = { added: 0, present: 0, broken: [] };
    let audiencesDropped = 0;
    const walk = async (nodes: INavNode[], parentId: number | undefined, existing: ITargetNavNode[]): Promise<void> => {
      for (const node of nodes) {
        throwIfAborted(ctx.signal);
        const url = node.url ? resolveNavUrl(node.url, ctx.tokens) : LINKLESS;
        const match = this._find(existing, node.title, url, ctx);
        if (match) {
          counts.present++;
          if (node.children && node.children.length) await walk(node.children, match.Id, match.Children || []);
          continue;
        }
        const hasChildren = !!(node.children && node.children.length);
        let target = url;
        let external = !!node.isExternal || url === LINKLESS;
        if (UNRESOLVED.test(url)) {
          counts.broken.push(node.title);
          if (!hasChildren) continue;
          target = LINKLESS;
          external = true;
        }
        let id: number;
        try {
          id = await addNavNode(sp, def.part as NavMenu, parentId, { Title: node.title, Url: target, IsExternal: external });
        } catch (e) {
          if (!isMissingLinkTarget(e)) throw e;
          counts.broken.push(node.title);
          if (!hasChildren) continue;
          id = await addNavNode(sp, def.part as NavMenu, parentId, { Title: node.title, Url: LINKLESS, IsExternal: true });
        }
        counts.added++;
        if (node.audiences && node.audiences.length) {
          if (sameTenant) await setNavNodeAudiences(sp, id, node.audiences);
          else audiencesDropped++;
        }
        if (hasChildren) await walk(node.children!, id, []);
      }
    };
    await walk(def.nodes, undefined, await readTargetMenu(sp, def.part));

    if (counts.broken.length) {
      ctx.log.warn(`Links whose target is not on this site were left out: ${counts.broken.join(', ')}.`, { artifact: ref, code: 'NAV_LINK_BROKEN', detail: counts.broken });
    }
    if (audiencesDropped) {
      ctx.log.warn(`Audiences of ${audiencesDropped} links were left out: their groups belong to the source tenant.`, { artifact: ref, code: 'NAV_AUDIENCE_OTHER_TENANT' });
    }
    ctx.log.info(`Menu links added: ${counts.added}, already present: ${counts.present}.`, { artifact: ref });
    return { ref, outcome: counts.added > 0 ? 'created' : 'skipped' };
  }

  private async _home(sp: SPFI, page: string, ref: IArtifactRef, ctx: IInstallContext): Promise<IApplyResult> {
    const wanted = `SitePages/${page}`;
    if ((await readWelcomePage(sp)).toLowerCase() === wanted.toLowerCase()) {
      ctx.log.info('The home page is already this page.', { artifact: ref });
      return { ref, outcome: 'skipped' };
    }
    if (!(await this._pageExists(sp, page))) {
      ctx.log.warn(`Page ${page} is not on this site; the home page is left as it is.`, { artifact: ref, code: 'NAV_HOME_PAGE_MISSING', detail: page });
      return { ref, outcome: 'skipped' };
    }
    await setWelcomePage(sp, wanted);
    ctx.log.info(`Home page set to ${page}.`, { artifact: ref });
    return { ref, outcome: 'updated' };
  }

  private async _pageExists(sp: SPFI, page: string): Promise<boolean> {
    return (await listSitePages(sp)).some((p) => p.FileName.toLowerCase() === page.toLowerCase());
  }

  /**
   * A node on the same level with the same title and URL. An unlinked heading matches by title on either side:
   * a heading whose link was broken on an earlier run was added without link and must not be added again.
   */
  private _find(existing: ITargetNavNode[], title: string, url: string, ctx: IInstallContext): ITargetNavNode | undefined {
    const want = comparableUrl(url, ctx.targetSiteUrl);
    return existing.filter(
      (e) =>
        e.Title.trim().toLowerCase() === title.trim().toLowerCase() &&
        (url === LINKLESS || e.Url === LINKLESS || comparableUrl(e.Url, ctx.targetSiteUrl) === want)
    )[0];
  }

  /** Links of the template not yet in the target menu (a missing parent counts with its subtree). */
  private _missing(nodes: INavNode[], existing: ITargetNavNode[], ctx: IInstallContext): number {
    return nodes.reduce((n, node) => {
      const match = this._find(existing, node.title, node.url ? resolveNavUrl(node.url, ctx.tokens) : LINKLESS, ctx);
      if (!match) return n + 1 + countNodes(node.children);
      return n + this._missing(node.children || [], match.Children || [], ctx);
    }, 0);
  }
}

