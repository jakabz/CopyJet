import { toTemplateKey } from '../keys';
import type { IGroup } from '../model';

export interface IGroupInfoLike {
  Id: number;
  Title: string;
  Description?: string;
  OnlyAllowMembersViewMembership?: boolean;
  AllowMembersEditMembership?: boolean;
  AllowRequestToJoinLeave?: boolean;
  Owner?: { Id: number; PrincipalType: number };
}

export const GROUP_SELECT = ['Id', 'Title', 'Description', 'OnlyAllowMembersViewMembership', 'AllowMembersEditMembership', 'AllowRequestToJoinLeave', 'Owner/Id', 'Owner/PrincipalType'];

/** SharePoint principal type of a SharePoint group. */
export const PRINCIPAL_SHAREPOINT_GROUP = 8;

export interface IRoleBinding {
  Name: string;
  RoleTypeKind: number;
}

/**
 * Built-in permission levels by RoleTypeKind; their names are language-dependent ("Munkatárs" /
 * "Contribute"), the kind is not (spike 07). The template uses these English names.
 */
const BUILT_IN_ROLES: { [kind: number]: string } = { 2: 'Read', 3: 'Contribute', 4: 'Design', 5: 'Full Control', 6: 'Edit' };

/** Limited Access (1) is granted implicitly by SharePoint and never copied. */
const LIMITED_ACCESS = 1;

/** Template name of a permission level: stable English name for built-ins, own name for custom ones. */
export function roleName(binding: IRoleBinding): string | undefined {
  if (binding.RoleTypeKind === LIMITED_ACCESS) return undefined;
  return BUILT_IN_ROLES[binding.RoleTypeKind] || binding.Name;
}

/** RoleTypeKind for a template role name, or undefined for a custom level (looked up by name). */
export function roleKind(name: string): number | undefined {
  const kinds = Object.keys(BUILT_IN_ROLES).map(Number);
  return kinds.filter((k) => BUILT_IN_ROLES[k].toLowerCase() === name.toLowerCase())[0];
}

/** "Forrás Projektmenedzserek" on site "Forrás" → "{sitename} Projektmenedzserek" (the target site's title is used). */
export function templateGroupTitle(title: string, siteTitle: string): string {
  return siteTitle && title.indexOf(`${siteTitle} `) === 0 ? `{sitename}${title.slice(siteTitle.length)}` : title;
}

/** Key from the title without the site-title prefix: "{sitename} Projektmenedzserek" → "Projektmenedzserek". */
export function groupKeyName(templateTitle: string): string {
  return toTemplateKey(templateTitle.replace(/^\{sitename\}\s*/, ''), 'group');
}

export const groupRefKey = (key: string): string => `group:${key}`;

export function toGroupDef(info: IGroupInfoLike, key: string, siteTitle: string, owner: string, roles: string[]): IGroup {
  const def: IGroup = { key, title: templateGroupTitle(info.Title, siteTitle), owner, roles, includeMembers: false, members: [] };
  if (info.Description) def.description = info.Description;
  def.onlyAllowMembersViewMembership = !!info.OnlyAllowMembersViewMembership;
  def.allowMembersEditMembership = !!info.AllowMembersEditMembership;
  def.allowRequestToJoinLeave = !!info.AllowRequestToJoinLeave;
  return def;
}

const FLAGS: Array<[keyof IGroup, keyof IGroupInfoLike]> = [
  ['onlyAllowMembersViewMembership', 'OnlyAllowMembersViewMembership'],
  ['allowMembersEditMembership', 'AllowMembersEditMembership'],
  ['allowRequestToJoinLeave', 'AllowRequestToJoinLeave']
];

/** Group settings to MERGE for the given changes. */
export function groupUpdateProps(def: IGroup, changes: string[]): { [prop: string]: unknown } {
  const props: { [prop: string]: unknown } = {};
  if (changes.indexOf('description') >= 0) props.Description = def.description || '';
  FLAGS.forEach(([p, r]) => {
    if (changes.indexOf(p as string) >= 0) props[r] = !!def[p];
  });
  return props;
}

/**
 * Differences between a (title-resolved) template group and the target group: description, membership
 * flags, missing permission levels (extra ones on the target are fine), owner.
 */
export function compareGroups(def: IGroup, target: IGroupInfoLike, targetRoles: string[], ownerId: number | undefined): string[] {
  const changes: string[] = [];
  if ((def.description || '') !== (target.Description || '')) changes.push('description');
  FLAGS.forEach(([p, r]) => {
    if (def[p] !== undefined && !!def[p] !== !!target[r]) changes.push(p as string);
  });
  const have = targetRoles.map((r) => r.toLowerCase());
  if ((def.roles || []).some((r) => have.indexOf(r.toLowerCase()) < 0)) changes.push('roles');
  if (ownerId !== undefined && (!target.Owner || target.Owner.Id !== ownerId)) changes.push('owner');
  return changes;
}
