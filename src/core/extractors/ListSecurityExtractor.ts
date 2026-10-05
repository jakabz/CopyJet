import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import { PrincipalCollector, SITE_USER_SELECT, type ISiteUserLike } from '../items';
import { isSupportedTemplate, loadSourceSite } from '../lists';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, ITemplateWriter } from '../model';
import { listSecurityDependencies } from '../planner/dependencies';
import { readSourceGroups } from './GroupExtractor';
import { listSecurityKey, readListRoleAssignments, toListSecurity, type IListSecurityDef } from '../security';

/**
 * Unique permissions of lists (spike 15): only lists that do not inherit are carried; their assignments go into
 * the template's list entry (list.security), which the ListExtractor must have created.
 */
export class ListSecurityExtractor implements IExtractor<IListSecurityDef> {
  public readonly kind = 'listSecurity' as const;

  public async discover(sp: SPFI, signal?: AbortSignal): Promise<IDiscoveredArtifact[]> {
    const site = await loadSourceSite(sp, signal);
    const unique = await this._uniqueListIds(sp);
    const lists = site.lists.filter((l) => isSupportedTemplate(l.info.BaseTemplate) && unique[l.info.Id.toLowerCase()]);
    const counts = await limitConcurrency(lists.map((l) => () => readListRoleAssignments(sp, l.info.RootFolder.ServerRelativeUrl)), 4, signal);
    return lists.map((l, i) => {
      const r = counts[i];
      const found: IDiscoveredArtifact = { ref: { kind: this.kind, key: listSecurityKey(l.key) }, title: 'uniquePermissions', parentKey: `list:${l.key}` };
      if (r.ok) found.itemCount = r.value.filter((a) => a.RoleDefinitionBindings.some((b) => b.RoleTypeKind !== 1)).length;
      return found;
    });
  }

  public dependencies(def: IListSecurityDef): IArtifactRef[] {
    return listSecurityDependencies(def);
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) return;
    const [site, source, users] = await Promise.all([
      loadSourceSite(sp, opts.signal),
      readSourceGroups(sp, opts.signal),
      sp.web.siteUsers.select(...SITE_USER_SELECT)<ISiteUserLike[]>()
    ]);
    // Groups by the GroupExtractor's keys, also those not selected: the Setup then offers them as dependencies.
    const groupKeyById: { [id: number]: string } = {};
    source.groups.forEach((g) => (groupKeyById[g.info.Id] = g.key));
    const people = new PrincipalCollector(out.manifest.principals, users);
    let count = 0;

    for (const l of site.lists.filter((x) => wanted.indexOf(listSecurityKey(x.key)) >= 0)) {
      throwIfAborted(opts.signal);
      const ref: IArtifactRef = { kind: this.kind, key: listSecurityKey(l.key) };
      const listDef = out.manifest.lists.filter((x) => x.key === l.key)[0];
      if (!listDef) {
        opts.log.warn(`List ${l.info.Title} is not in the template; its permissions are skipped.`, { artifact: ref, code: 'LIST_SECURITY_LIST_MISSING' });
        continue;
      }
      const converted = toListSecurity(await readListRoleAssignments(sp, l.info.RootFolder.ServerRelativeUrl), { associated: source.security.associated, groupKeyById, people });
      listDef.security = converted.security;
      count++;
      converted.skipped
        .filter((s) => s.reason !== 'limitedAccessOnly')
        .forEach((s) =>
          opts.log.warn(
            s.reason === 'groupOutsideTemplate'
              ? `Group ${s.member} has permissions on ${l.info.Title} but none on the site, so CopyJet does not copy it; its list permissions are not carried.`
              : `${s.member} has permissions on ${l.info.Title} but is not an account CopyJet carries; left out.`,
            { artifact: ref, code: s.reason === 'groupOutsideTemplate' ? 'LIST_SECURITY_GROUP_OUTSIDE_TEMPLATE' : 'LIST_SECURITY_PRINCIPAL_UNSUPPORTED', detail: s.member }
          )
        );
      opts.log.info(`Unique permissions: ${(converted.security.roleAssignments || []).length} assignments.`, { artifact: ref });
    }
    opts.log.info(`Extracted: ${count} × listSecurity.`, { step: this.kind });
  }

  private async _uniqueListIds(sp: SPFI): Promise<{ [id: string]: boolean }> {
    const lists = await sp.web.lists.select('Id', 'HasUniqueRoleAssignments').filter('Hidden eq false')<Array<{ Id: string; HasUniqueRoleAssignments: boolean }>>();
    const out: { [id: string]: boolean } = {};
    lists.forEach((l) => l.HasUniqueRoleAssignments && (out[l.Id.toLowerCase()] = true));
    return out;
  }
}
