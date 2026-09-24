import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import { throwIfAborted } from '../errors';
import {
  PRINCIPAL_SHAREPOINT_GROUP,
  groupKeyName,
  groupRefKey,
  readWebSecurity,
  roleKind,
  templateGroupTitle,
  toGroupDef,
  type IGroupInfoLike,
  type IWebSecurity
} from '../groups';
import { uniqueKeys } from '../keys';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, IGroup, ITemplateWriter } from '../model';

interface ISourceGroup {
  info: IGroupInfoLike;
  key: string;
}

/**
 * The site's own SharePoint groups: groups with a web role assignment, except the associated
 * Owners/Members/Visitors groups, which exist on every target and are referenced by token (spike 07).
 */
export class GroupExtractor implements IExtractor<IGroup> {
  public readonly kind = 'group' as const;

  public async discover(sp: SPFI, signal?: AbortSignal): Promise<IDiscoveredArtifact[]> {
    const { groups } = await this._source(sp, signal);
    return groups.map((g) => ({ ref: { kind: this.kind, key: groupRefKey(g.key) }, title: g.info.Title }));
  }

  /** A group owning this one. */
  public dependencies(def: IGroup): IArtifactRef[] {
    const m = def.owner ? /^\{groupkey:([^{}]+)\}$/.exec(def.owner) : null;
    return m && m[1] !== def.key ? [{ kind: 'group', key: groupRefKey(m[1]) }] : [];
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) {
      return;
    }
    const { groups, security, siteTitle } = await this._source(sp, opts.signal);
    const keyById: { [id: number]: string } = {};
    groups.forEach((g) => (keyById[g.info.Id] = g.key));

    groups
      .filter((g) => wanted.indexOf(groupRefKey(g.key)) >= 0)
      .forEach((g) => {
        const ref: IArtifactRef = { kind: this.kind, key: groupRefKey(g.key) };
        const roles = security.rolesByPrincipal[g.info.Id] || [];
        roles
          .filter((r) => roleKind(r) === undefined)
          .forEach((r) =>
            opts.log.warn(`Custom permission level "${r}" is matched by name on the target; it is not created there.`, { artifact: ref, code: 'GROUP_CUSTOM_ROLE' })
          );
        const def = toGroupDef(g.info, g.key, siteTitle, this._ownerToken(g.info, security, keyById, opts, ref), roles);
        const at = out.manifest.groups.findIndex((x) => x.key === def.key);
        if (at >= 0) {
          out.manifest.groups[at] = def;
        } else {
          out.manifest.groups.push(def);
        }
      });

    wanted
      .filter((key) => !groups.some((g) => groupRefKey(g.key) === key))
      .forEach((key) => opts.log.warn('Group not found on the source site.', { artifact: { kind: this.kind, key }, code: 'GROUP_NOT_FOUND' }));
  }

  /** Owner as a token: an associated group, another copied group, or – for a user – the Owners group until users are mapped (phase 2). */
  private _ownerToken(info: IGroupInfoLike, security: IWebSecurity, keyById: { [id: number]: string }, opts: IExtractOptions, ref: IArtifactRef): string {
    const owner = info.Owner;
    const a = security.associated;
    if (owner && owner.PrincipalType === PRINCIPAL_SHAREPOINT_GROUP) {
      if (owner.Id === a.owner) return '{associatedownergroup}';
      if (owner.Id === a.member) return '{associatedmembergroup}';
      if (owner.Id === a.visitor) return '{associatedvisitorgroup}';
      if (keyById[owner.Id]) return `{groupkey:${keyById[owner.Id]}}`;
    }
    opts.log.warn('The group owner is a user or a group that is not copied; the Owners group will own it on the target.', { artifact: ref, code: 'GROUP_OWNER_USER' });
    return '{associatedownergroup}';
  }

  private async _source(sp: SPFI, signal?: AbortSignal): Promise<{ groups: ISourceGroup[]; security: IWebSecurity; siteTitle: string }> {
    const [security, web] = await Promise.all([readWebSecurity(sp, signal), sp.web.select('Title')<{ Title: string }>()]);
    const a = security.associated;
    const own = security.groups.filter((g) => security.rolesByPrincipal[g.Id] !== undefined && [a.owner, a.member, a.visitor].indexOf(g.Id) < 0);
    const keys = uniqueKeys(own.map((g) => [String(g.Id), groupKeyName(templateGroupTitle(g.Title, web.Title))] as [string, string]), 'group');
    return { groups: own.map((info) => ({ info, key: keys[String(info.Id)] })), security, siteTitle: web.Title };
  }
}
