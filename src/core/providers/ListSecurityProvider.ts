import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/security';
import '@pnp/sp/site-users/web';
import { CopyJetError, throwIfAborted } from '../errors';
import { roleDefinitionId, roleName } from '../groups';
import { isHttpStatus } from '../http/status';
import { toServerRelativeUrl } from '../lists';
import type { ConflictMode, IApplyResult, IArtifactRef, IDiffResult, IInstallContext, IProvider, IRoleAssignment } from '../model';
import { listSecurityKey, readListRoleAssignments, targetPrincipal, type IListSecurityDef, type IRoleAssignmentLike } from '../security';

const FULL_CONTROL = 'Full Control';

interface IWanted {
  assignment: IRoleAssignment;
  /** Target group ID, or the login of a person / directory group. */
  groupId?: number;
  login?: string;
}

/**
 * Unique list permissions (spike 15). A list that inherits gets its inheritance broken without copying (the source
 * named every principal), then the template's assignments are added; adding one already there is a no-op on
 * SharePoint. Nothing else is removed – only the installer's own Full Control that breaking the inheritance put
 * there, when the template does not name the installer and the installer is a site collection admin (otherwise
 * removing it could lock the installer out; it is kept with a note). In 'skip' mode an existing list is left alone.
 */
export class ListSecurityProvider implements IProvider<IListSecurityDef> {
  public readonly kind = 'listSecurity' as const;

  public async diff(sp: SPFI, def: IListSecurityDef, ctx: IInstallContext): Promise<IDiffResult> {
    throwIfAborted(ctx.signal);
    const ref: IArtifactRef = { kind: this.kind, key: listSecurityKey(def.listKey) };
    const url = this._listUrl(def, ctx);
    let unique: boolean;
    try {
      unique = (await sp.web.getList(url).select('HasUniqueRoleAssignments')<{ HasUniqueRoleAssignments: boolean }>()).HasUniqueRoleAssignments;
    } catch (e) {
      if (isHttpStatus(e, 404)) return { ref, status: 'unsupported', changes: ['listMissing'] };
      throw e;
    }
    const missing = this._missing(await this._wanted(def, ctx), unique ? await readListRoleAssignments(sp, url) : []);
    if (!unique) return { ref, status: 'different', changes: ['inheritsPermissions', `assignments:${missing}`] };
    return missing > 0 ? { ref, status: 'different', changes: [`assignments:${missing}`] } : { ref, status: 'same' };
  }

  public async apply(sp: SPFI, def: IListSecurityDef, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    const ref = diff.ref;
    if (diff.status === 'same') {
      ctx.log.info('List permissions already match.', { artifact: ref });
      return { ref, outcome: 'skipped' };
    }
    if (diff.status === 'unsupported') {
      ctx.log.warn('The list is not on the target; its permissions are skipped.', { artifact: ref, code: 'LIST_SECURITY_LIST_MISSING' });
      return { ref, outcome: 'skipped' };
    }
    if (mode === 'skip') {
      ctx.log.info('The list existed before this install; its permissions are left as they are (conflict mode: skip).', { artifact: ref, code: 'LIST_SECURITY_SKIPPED' });
      return { ref, outcome: 'skipped' };
    }
    const list = sp.web.getList(this._listUrl(def, ctx));
    const broke = (diff.changes || []).indexOf('inheritsPermissions') >= 0;
    if (broke) {
      // clearSubscopes=false: items with permissions of their own on the target keep them.
      await list.breakRoleInheritance(false, false);
    }

    const roleIds: { [name: string]: Promise<number | undefined> } = {};
    const roleId = (name: string): Promise<number | undefined> => (roleIds[name] = roleIds[name] || roleDefinitionId(sp, name));
    let added = 0;
    const principalIds: number[] = [];
    for (const w of await this._wanted(def, ctx)) {
      throwIfAborted(ctx.signal);
      if (w.groupId === undefined && w.login === undefined) {
        ctx.log.warn(`${w.assignment.principal} is not on the target (not in the template or not mapped); its permissions are skipped.`, {
          artifact: ref,
          code: 'LIST_SECURITY_PRINCIPAL_MISSING',
          detail: w.assignment.principal
        });
        continue;
      }
      let id = w.groupId;
      if (id === undefined) {
        try {
          id = (await sp.web.ensureUser(w.login!)).Id;
        } catch {
          ctx.log.warn(`${w.login} is not a user of the target tenant; its permissions are skipped.`, { artifact: ref, code: 'LIST_SECURITY_PRINCIPAL_MISSING', detail: w.login });
          continue;
        }
      }
      principalIds.push(id);
      for (const name of w.assignment.roles) {
        const rid = await roleId(name);
        if (rid === undefined) {
          ctx.log.warn(`Permission level "${name}" does not exist on the target site; not assigned.`, { artifact: ref, code: 'ROLE_NOT_FOUND', detail: name });
          continue;
        }
        await list.roleAssignments.add(id, rid);
        added++;
      }
    }

    if (broke) await this._installer(sp, def, principalIds, ref, ctx);
    ctx.log.info(`List permissions: ${broke ? 'inheritance broken, ' : ''}${added} role assignments added.`, { artifact: ref });
    return { ref, outcome: 'updated' };
  }

  /** Breaking the inheritance left the installer with Full Control; see the class comment. */
  private async _installer(sp: SPFI, def: IListSecurityDef, principalIds: number[], ref: IArtifactRef, ctx: IInstallContext): Promise<void> {
    const me = await sp.web.currentUser.select('Id', 'IsSiteAdmin')<{ Id: number; IsSiteAdmin: boolean }>();
    if (principalIds.indexOf(me.Id) >= 0) return;
    if (!me.IsSiteAdmin) {
      ctx.log.warn('You keep Full Control on the list: removing it could lock you out (you are not a site collection admin).', {
        artifact: ref,
        code: 'LIST_SECURITY_INSTALLER_KEPT'
      });
      return;
    }
    const full = await roleDefinitionId(sp, FULL_CONTROL);
    if (full === undefined) throw new CopyJetError('ROLE_NOT_FOUND', 'The target site has no Full Control level.');
    await sp.web.getList(this._listUrl(def, ctx)).roleAssignments.remove(me.Id, full);
    ctx.log.info('Your own Full Control, added by breaking the inheritance, was removed (the source list did not name you).', {
      artifact: ref,
      code: 'LIST_SECURITY_INSTALLER_REMOVED'
    });
  }

  /**
   * The assignments with their target principals. People are mapped here (the mapping step's choices); a person
   * mapped only to the fallback user is left out – the fallback must not get the person's permissions.
   */
  private async _wanted(def: IListSecurityDef, ctx: IInstallContext): Promise<IWanted[]> {
    const assignments = def.security.roleAssignments || [];
    const fallback: { [principal: string]: boolean } = {};
    const keys = assignments.map((a) => /^\{principal:([^{}]+)\}$/.exec(a.principal)).filter((m): m is RegExpExecArray => !!m).map((m) => m[1]);
    if (keys.length && ctx.content) {
      (await ctx.content.principals.map(keys, ctx.tokens, ctx.log, ctx.signal)).forEach((m) => m.strategy === 'fallback' && (fallback[`{principal:${m.key}}`] = true));
    }
    return assignments.map((assignment) => ({ assignment, ...((!fallback[assignment.principal] && targetPrincipal(assignment.principal, ctx.tokens)) || {}) }));
  }

  /** Template (principal, role) pairs the target list does not have. */
  private _missing(wanted: IWanted[], existing: IRoleAssignmentLike[]): number {
    const have = (w: IWanted, role: string): boolean =>
      existing.some(
        (e) =>
          (w.groupId !== undefined ? e.Member.Id === w.groupId : !!w.login && e.Member.LoginName.toLowerCase() === w.login.toLowerCase()) &&
          e.RoleDefinitionBindings.some((b) => (roleName(b) || '').toLowerCase() === role.toLowerCase())
      );
    return wanted.reduce((n, w) => n + w.assignment.roles.filter((r) => !have(w, r)).length, 0);
  }

  private _listUrl(def: IListSecurityDef, ctx: IInstallContext): string {
    const web = ctx.tokens.get('siterelative');
    if (web === undefined) throw new CopyJetError('TOKEN_UNRESOLVED', 'The install context has no {siterelative} value.');
    return toServerRelativeUrl(ctx.tokens.get('listurl', def.listKey) || def.listUrl, web);
  }
}
