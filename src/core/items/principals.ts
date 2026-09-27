import { toTemplateKey } from '../keys';
import type { IPrincipal } from '../model';

/** A site user as REST returns it (web/siteusers). */
export interface ISiteUserLike {
  Id: number;
  LoginName: string;
  Title?: string;
  Email?: string;
  UserPrincipalName?: string;
  PrincipalType?: number;
}

export const SITE_USER_SELECT = ['Id', 'LoginName', 'Title', 'Email', 'UserPrincipalName', 'PrincipalType'];

/** Principal kind from the claims login; undefined for accounts CopyJet does not carry (system, apps). */
export function principalKind(loginName: string): IPrincipal['kind'] | undefined {
  const login = loginName.toLowerCase();
  if (/^i:0#\.f\|membership\|/.test(login)) return /urn%3aspo%3a|app@sharepoint/.test(login) ? undefined : 'user';
  if (login.indexOf('c:0o.c|federateddirectoryclaimprovider|') === 0) return 'm365Group';
  if (login.indexOf('c:0t.c|tenant|') === 0) return 'securityGroup';
  if (login === 'c:0(.s|true' || login.indexOf('c:0-.f|rolemanager|spo-grid-all-users') === 0) return 'everyone';
  return undefined;
}

/**
 * Collects the people item values refer to into the template's principals and hands out their
 * {principal:key} tokens. Keys come from the e-mail's local part (or the display name) and stay unique.
 */
export class PrincipalCollector {
  private readonly _principals: IPrincipal[];
  private readonly _byId: { [id: number]: ISiteUserLike } = {};
  private readonly _tokens: { [id: number]: string | undefined } = {};

  /** `principals` is the template's collection (appended to); `users` the source site's users. */
  constructor(principals: IPrincipal[], users: ISiteUserLike[]) {
    this._principals = principals;
    users.forEach((u) => (this._byId[u.Id] = u));
  }

  /** '{principal:key}' for a site user ID, or undefined (unknown ID, SharePoint group, system account). */
  public token(userId: number): string | undefined {
    if (Object.prototype.hasOwnProperty.call(this._tokens, userId)) return this._tokens[userId];
    const user = this._byId[userId];
    const kind = user ? principalKind(user.LoginName) : undefined;
    let token: string | undefined;
    if (user && kind) {
      const existing = this._principals.filter((p) => (p.loginName || '').toLowerCase() === user.LoginName.toLowerCase())[0];
      token = `{principal:${(existing || this._add(user, kind)).key}}`;
    }
    this._tokens[userId] = token;
    return token;
  }

  private _add(user: ISiteUserLike, kind: IPrincipal['kind']): IPrincipal {
    const base = toTemplateKey(user.Email ? user.Email.split('@')[0] : user.Title || '', 'principal');
    const used = (k: string): boolean => this._principals.some((p) => p.key.toLowerCase() === k.toLowerCase());
    let key = base;
    for (let n = 2; used(key); n++) key = `${base}_${n}`;
    const p: IPrincipal = { key, kind, loginName: user.LoginName };
    if (user.Email) p.email = user.Email;
    if (user.UserPrincipalName) p.upn = user.UserPrincipalName;
    if (user.Title) p.displayName = user.Title;
    this._principals.push(p);
    return p;
  }
}
