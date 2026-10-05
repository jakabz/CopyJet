import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/security';
import { PRINCIPAL_SHAREPOINT_GROUP, roleName, type IRoleBinding } from '../groups';
import type { PrincipalCollector } from '../items';
import type { ICopyJetTemplate, IRoleAssignment, ISecurity } from '../model';
import { resolveKnown, type TokenContext } from '../tokenizer';

/** An install step: the unique permissions of one template list (spike 15). */
export interface IListSecurityDef {
  listKey: string;
  /** Site-relative list URL from the template; a renamed copy is found through {listurl:K}. */
  listUrl: string;
  security: ISecurity;
}

export const listSecurityKey = (listKey: string): string => `listSecurity:${listKey}`;

/** Steps for the template lists whose permissions were unique on the source. */
export function listSecurityDefs(template: ICopyJetTemplate): IListSecurityDef[] {
  return template.lists
    .filter((l) => l.security && l.security.breakInheritance && (l.security.roleAssignments || []).length > 0)
    .map((l) => ({ listKey: l.key, listUrl: l.url, security: l.security! }));
}

/** A role assignment of a securable as REST returns it. */
export interface IRoleAssignmentLike {
  Member: { Id: number; Title: string; LoginName: string; PrincipalType: number };
  RoleDefinitionBindings: IRoleBinding[];
}

export const ROLE_ASSIGNMENT_SELECT = [
  'Member/Id',
  'Member/Title',
  'Member/LoginName',
  'Member/PrincipalType',
  'RoleDefinitionBindings/Name',
  'RoleDefinitionBindings/RoleTypeKind'
];

export async function readListRoleAssignments(sp: SPFI, listServerRelativeUrl: string): Promise<IRoleAssignmentLike[]> {
  return sp.web
    .getList(listServerRelativeUrl)
    .roleAssignments.select(...ROLE_ASSIGNMENT_SELECT)
    .expand('Member', 'RoleDefinitionBindings')<IRoleAssignmentLike[]>();
}

export interface ISourcePrincipals {
  associated: { owner?: number; member?: number; visitor?: number };
  /** The source web's own groups by ID → their template keys (the GroupExtractor's keys). */
  groupKeyById: { [id: number]: string };
  people: PrincipalCollector;
}

export interface IConvertedSecurity {
  security: ISecurity;
  /** Members left out, with the reason (group without a web role, Limited Access only, unknown account). */
  skipped: Array<{ member: string; reason: 'groupOutsideTemplate' | 'limitedAccessOnly' | 'unsupportedPrincipal' }>;
}

/**
 * The source list's assignments in template form (spike 15 A): associated groups as {associated…group}, template
 * groups as {groupkey:K}, people and directory groups as {principal:key}. Limited Access is SharePoint's own
 * (sharing) and never carried; a member holding only that is left out.
 */
export function toListSecurity(assignments: IRoleAssignmentLike[], src: ISourcePrincipals): IConvertedSecurity {
  const roleAssignments: IRoleAssignment[] = [];
  const skipped: IConvertedSecurity['skipped'] = [];
  assignments.forEach((a) => {
    const roles = a.RoleDefinitionBindings.map(roleName).filter((n, i, all): n is string => !!n && all.indexOf(n) === i);
    if (!roles.length) {
      skipped.push({ member: a.Member.Title, reason: 'limitedAccessOnly' });
      return;
    }
    const principal = principalToken(a.Member, src);
    if (!principal) {
      skipped.push({ member: a.Member.Title, reason: a.Member.PrincipalType === PRINCIPAL_SHAREPOINT_GROUP ? 'groupOutsideTemplate' : 'unsupportedPrincipal' });
      return;
    }
    roleAssignments.push({ principal, roles: roles as [string, ...string[]] });
  });
  roleAssignments.sort((x, y) => (x.principal < y.principal ? -1 : x.principal > y.principal ? 1 : 0));
  return { security: { breakInheritance: true, copyRoleAssignments: false, roleAssignments }, skipped };
}

function principalToken(member: IRoleAssignmentLike['Member'], src: ISourcePrincipals): string | undefined {
  if (member.PrincipalType === PRINCIPAL_SHAREPOINT_GROUP) {
    if (member.Id === src.associated.owner) return '{associatedownergroup}';
    if (member.Id === src.associated.member) return '{associatedmembergroup}';
    if (member.Id === src.associated.visitor) return '{associatedvisitorgroup}';
    const key = src.groupKeyById[member.Id];
    return key ? `{groupkey:${key}}` : undefined;
  }
  return src.people.token(member.Id);
}


/** A template principal on the target: a group ID, or a login to ensure; undefined when unresolved. */
export function targetPrincipal(principal: string, tokens: TokenContext): { groupId?: number; login?: string } | undefined {
  const value = resolveKnown(principal, tokens);
  if (/\{[a-z]+(:[^{}]*)?\}/.test(value)) return undefined;
  if (/^\{(associated|groupkey)/.test(principal)) {
    const id = Number(value);
    return isNaN(id) ? undefined : { groupId: id };
  }
  return { login: value };
}
