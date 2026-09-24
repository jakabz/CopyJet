import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-groups';
import { RoleAssignment } from '@pnp/sp/security';
import { CopyJetError, throwIfAborted } from '../errors';
import {
  GROUP_SELECT,
  compareGroups,
  groupRefKey,
  groupUpdateProps,
  roleKind,
  roleName,
  type IGroupInfoLike,
  type IRoleBinding
} from '../groups';
import { csomSetGroupOwner, type FetchLike } from '../http/csom';
import { isHttpStatus } from '../http/status';
import type { ConflictMode, IApplyResult, IDiffResult, IGroup, IInstallContext, IProvider } from '../model';
import { resolve } from '../tokenizer';

interface IGroupDiff extends IDiffResult {
  target?: IGroupInfoLike;
}

const odataString = (s: string): string => s.replace(/'/g, "''");

/** Changes 'update' mode applies; permission levels are only added, never removed. */
const UPDATABLE = ['description', 'onlyAllowMembersViewMembership', 'allowMembersEditMembership', 'allowRequestToJoinLeave', 'roles', 'owner'];

export class GroupProvider implements IProvider<IGroup> {
  public readonly kind = 'group' as const;
  private readonly _fetch?: FetchLike;

  /** `fetchImpl` is used for the CSOM call that sets the owner (injected in tests). */
  constructor(fetchImpl?: FetchLike) {
    this._fetch = fetchImpl;
  }

  public async diff(sp: SPFI, def: IGroup, ctx: IInstallContext): Promise<IGroupDiff> {
    throwIfAborted(ctx.signal);
    const ref = { kind: this.kind, key: groupRefKey(def.key) };
    const title = resolve(def.title, ctx.tokens);
    const target = await this._find(sp, title);
    if (!target) {
      return { ref, status: 'new' };
    }
    const changes = compareGroups(def, target, await this._roles(sp, target.Id), this._ownerId(def, ctx));
    return { ref, status: changes.length ? 'different' : 'same', changes: changes.length ? changes : undefined, target };
  }

  public async apply(sp: SPFI, def: IGroup, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;

    switch (diff.status) {
      case 'new':
        return this._create(sp, def, resolve(def.title, ctx.tokens), ctx);
      case 'same':
        ctx.log.info('Group already present; skipped.', { artifact: ref });
        return this._done(ref, 'skipped', def, diff.target!.Id, ctx);
      case 'different':
        if (mode === 'update') {
          return this._update(sp, def, diff, ctx);
        }
        if (mode === 'rename') {
          const title = `${resolve(def.title, ctx.tokens)} (copy)`;
          const copy = await this._find(sp, title);
          return copy ? this._done(ref, 'skipped', def, copy.Id, ctx) : this._create(sp, def, title, ctx);
        }
        ctx.log.info(`Group differs (${(diff.changes || []).join(', ')}); kept as it is.`, { artifact: ref });
        return this._done(ref, 'skipped', def, diff.target!.Id, ctx);
      default:
        ctx.log.warn(`Group skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'GROUP_UNSUPPORTED', detail: diff.changes });
        return { ref, outcome: 'skipped' };
    }
  }

  private async _create(sp: SPFI, def: IGroup, title: string, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = { kind: this.kind, key: groupRefKey(def.key) };
    const created = await sp.web.siteGroups.add({ Title: title, ...groupUpdateProps(def, UPDATABLE) });
    if (!created.Id) {
      throw new CopyJetError('GROUP_CREATE_FAILED', `SharePoint returned no Id for group ${title}.`);
    }
    // The token is registered first so the group can own itself ({groupkey:self}).
    this._done(ref, 'created', def, created.Id, ctx);
    await this._setOwner(sp, def, created.Id, ctx);
    await this._addRoles(sp, def, created.Id, [], ctx);
    ctx.log.info('Group created.', { artifact: ref });
    return this._done(ref, 'created', def, created.Id, ctx);
  }

  private async _update(sp: SPFI, def: IGroup, diff: IGroupDiff, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = diff.ref;
    const target = diff.target!;
    const changes = diff.changes || [];
    const props = groupUpdateProps(def, changes);
    if (Object.keys(props).length) {
      await sp.web.siteGroups.getById(target.Id).update(props);
    }
    if (changes.indexOf('owner') >= 0) {
      await this._setOwner(sp, def, target.Id, ctx);
    }
    if (changes.indexOf('roles') >= 0) {
      await this._addRoles(sp, def, target.Id, await this._roles(sp, target.Id), ctx);
    }
    ctx.log.info(`Group updated: ${changes.join(', ')}.`, { artifact: ref });
    return this._done(ref, 'updated', def, target.Id, ctx);
  }

  /** Owner through CSOM (REST's SetUserAsOwner ignores groups, spike 07), then verified. */
  private async _setOwner(sp: SPFI, def: IGroup, groupId: number, ctx: IInstallContext): Promise<void> {
    const ref = { kind: this.kind, key: groupRefKey(def.key) };
    const ownerId = this._ownerId(def, ctx);
    if (ownerId === undefined) return;
    await csomSetGroupOwner(sp, groupId, ownerId, ctx.signal, this._fetch);
    const after = await sp.web.siteGroups.getById(groupId).select('Owner/Id').expand('Owner')<{ Owner?: { Id: number } }>();
    if (!after.Owner || after.Owner.Id !== ownerId) {
      ctx.log.warn('The group owner could not be set; SharePoint kept another owner.', { artifact: ref, code: 'GROUP_OWNER_NOT_SET', detail: { expected: ownerId } });
    }
  }

  /** Adds the template's permission levels the group does not have yet on the web. */
  private async _addRoles(sp: SPFI, def: IGroup, groupId: number, existing: string[], ctx: IInstallContext): Promise<void> {
    const ref = { kind: this.kind, key: groupRefKey(def.key) };
    const have = existing.map((r) => r.toLowerCase());
    for (const name of (def.roles || []).filter((r) => have.indexOf(r.toLowerCase()) < 0)) {
      throwIfAborted(ctx.signal);
      const roleDefId = await this._roleDefinitionId(sp, name);
      if (roleDefId === undefined) {
        ctx.log.warn(`Permission level "${name}" does not exist on the target site; not assigned.`, { artifact: ref, code: 'ROLE_NOT_FOUND', detail: name });
        continue;
      }
      await sp.web.roleAssignments.add(groupId, roleDefId);
    }
  }

  /** Built-in levels by RoleTypeKind (language-independent, spike 07), custom ones by name. */
  private async _roleDefinitionId(sp: SPFI, name: string): Promise<number | undefined> {
    const kind = roleKind(name);
    try {
      const d =
        kind !== undefined
          ? await sp.web.roleDefinitions.getByType(kind as 2 | 3 | 4 | 5 | 6).select('Id')<{ Id: number }>()
          : await sp.web.roleDefinitions.getByName(name).select('Id')<{ Id: number }>();
      return d && d.Id ? d.Id : undefined;
    } catch (e) {
      if (isHttpStatus(e, 404)) return undefined;
      throw e;
    }
  }

  /** Template role names of the group's web role assignment (none → []). */
  private async _roles(sp: SPFI, groupId: number): Promise<string[]> {
    try {
      const ra = await RoleAssignment(sp.web.roleAssignments, `getByPrincipalId(${groupId})`)
        .select('RoleDefinitionBindings/Name', 'RoleDefinitionBindings/RoleTypeKind')
        .expand('RoleDefinitionBindings')<{ RoleDefinitionBindings?: IRoleBinding[] }>();
      return (ra.RoleDefinitionBindings || []).map(roleName).filter((n): n is string => !!n);
    } catch (e) {
      if (isHttpStatus(e, 404)) return [];
      throw e;
    }
  }

  private _ownerId(def: IGroup, ctx: IInstallContext): number | undefined {
    if (!def.owner) return undefined;
    const id = Number(resolve(def.owner, ctx.tokens));
    return isNaN(id) ? undefined : id;
  }

  private async _find(sp: SPFI, title: string): Promise<IGroupInfoLike | undefined> {
    const found = await sp.web.siteGroups.filter(`Title eq '${odataString(title)}'`).select(...GROUP_SELECT).expand('Owner')<IGroupInfoLike[]>();
    return found[0];
  }

  private _done(ref: IApplyResult['ref'], outcome: IApplyResult['outcome'], def: IGroup, id: number, ctx: IInstallContext): IApplyResult {
    ctx.tokens.set('groupkey', def.key, String(id));
    return { ref, outcome, tokens: { groupkey: { [def.key]: String(id) } } };
  }
}
