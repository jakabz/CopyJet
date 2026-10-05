import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-groups';
import '@pnp/sp/security';
import { throwIfAborted } from '../errors';
import { isHttpStatus } from '../http/status';
import { GROUP_SELECT, PRINCIPAL_SHAREPOINT_GROUP, roleKind, roleName, type IGroupInfoLike, type IRoleBinding } from './groupModel';

export interface IAssociatedGroups {
  owner?: number;
  member?: number;
  visitor?: number;
}

export interface IWebSecurity {
  associated: IAssociatedGroups;
  groups: IGroupInfoLike[];
  /** Template role names of each principal's web-level role assignment. */
  rolesByPrincipal: { [principalId: number]: string[] };
}

async function associatedId(read: () => Promise<{ Id?: number }>): Promise<number | undefined> {
  try {
    const g = await read();
    return g && g.Id ? g.Id : undefined;
  } catch {
    // A web without that associated group.
    return undefined;
  }
}

/** The web's associated groups: their IDs back the {associated…group} tokens on either side. */
export async function readAssociatedGroups(sp: SPFI): Promise<IAssociatedGroups> {
  const [owner, member, visitor] = await Promise.all([
    associatedId(() => sp.web.associatedOwnerGroup.select('Id')<{ Id: number }>()),
    associatedId(() => sp.web.associatedMemberGroup.select('Id')<{ Id: number }>()),
    associatedId(() => sp.web.associatedVisitorGroup.select('Id')<{ Id: number }>())
  ]);
  return { owner, member, visitor };
}

/** Groups, associated groups and web role assignments of a web (spike 07 A). */
export async function readWebSecurity(sp: SPFI, signal?: AbortSignal): Promise<IWebSecurity> {
  throwIfAborted(signal);
  const [associated, groups, assignments] = await Promise.all([
    readAssociatedGroups(sp),
    sp.web.siteGroups.select(...GROUP_SELECT).expand('Owner')<IGroupInfoLike[]>(),
    sp.web.roleAssignments
      .select('PrincipalId', 'Member/PrincipalType', 'RoleDefinitionBindings/Name', 'RoleDefinitionBindings/RoleTypeKind')
      .expand('Member', 'RoleDefinitionBindings')<Array<{ PrincipalId: number; Member: { PrincipalType: number }; RoleDefinitionBindings: IRoleBinding[] }>>()
  ]);
  throwIfAborted(signal);
  const rolesByPrincipal: { [id: number]: string[] } = {};
  assignments
    .filter((a) => a.Member && a.Member.PrincipalType === PRINCIPAL_SHAREPOINT_GROUP)
    .forEach((a) => {
      rolesByPrincipal[a.PrincipalId] = a.RoleDefinitionBindings.map(roleName).filter((n): n is string => !!n);
    });
  return { associated, groups, rolesByPrincipal };
}

/**
 * ID of a template permission level on a web: built-in levels by RoleTypeKind (language-independent, spike 07;
 * Restricted View is kind 8, spike 15), custom ones by name from the web's levels – getByName answers a missing
 * level with HTTP 500, not 404. Undefined when the web has no such level.
 */
export async function roleDefinitionId(sp: SPFI, name: string): Promise<number | undefined> {
  const kind = roleKind(name);
  if (kind === undefined) {
    const all = await sp.web.roleDefinitions.select('Id', 'Name')<Array<{ Id: number; Name: string }>>();
    const hit = all.filter((d) => d.Name.toLowerCase() === name.toLowerCase())[0];
    return hit ? hit.Id : undefined;
  }
  try {
    const d = await sp.web.roleDefinitions.getByType(kind as 2 | 3 | 4 | 5 | 6).select('Id')<{ Id: number }>();
    return d && d.Id ? d.Id : undefined;
  } catch (e) {
    if (isHttpStatus(e, 404)) return undefined;
    throw e;
  }
}
