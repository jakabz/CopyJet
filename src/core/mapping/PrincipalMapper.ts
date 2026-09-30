import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import { isAbortError } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import type { Logger } from '../logger/Logger';
import type { IPrincipal } from '../model';
import type { TokenContext } from '../tokenizer';
import { replaceDomain, type IDomainRule } from './csv';

export type PrincipalStrategy = 'manual' | 'csv' | 'sameLogin' | 'sameEmail' | 'domain' | 'fallback';

export interface IPrincipalMapping {
  key: string;
  /** Target login name; undefined when the principal was not found on the target. */
  login?: string;
  strategy?: PrincipalStrategy;
}

/** What the MappingStep sets; every value is an e-mail, UPN or claims login ensureUser accepts. */
export interface IPrincipalRules {
  /** Principal key → target chosen by hand in the mapping table. */
  manual?: { [key: string]: string };
  /** Source e-mail / UPN / login (lower case) → target, from the CSV import. */
  csv?: { [source: string]: string };
  /** Source domain → target domain, applied to e-mails and logins. */
  domains?: IDomainRule[];
  /** Used for every principal nothing else found; without it their values stay empty. */
  fallback?: string;
}

/**
 * Template principals → target users (rendszerterv §5). Strategies in order: a target chosen by hand, the CSV
 * table, the same login (same tenant), the same e-mail, domain replacement, then the fallback user. The CSV
 * and the manual choice come first: they are explicit decisions. web.ensureUser accepts e-mails and logins;
 * each value is ensured once per mapper. Found principals become {principal:key} values in the token context.
 */
export class PrincipalMapper {
  private readonly _sp: SPFI;
  private readonly _byKey: { [key: string]: IPrincipal } = {};
  private _done: { [key: string]: IPrincipalMapping } = {};
  private readonly _ensured: { [value: string]: Promise<string | undefined> } = {};
  private readonly _warned: Map<Logger, string[]> = new Map();
  private _rules: IPrincipalRules = {};

  constructor(sp: SPFI, principals: IPrincipal[]) {
    this._sp = sp;
    principals.forEach((p) => (this._byKey[p.key] = p));
  }

  public get rules(): IPrincipalRules {
    return this._rules;
  }

  /** New rules: every principal is mapped again on the next map() (ensured values stay cached). */
  public setRules(rules: IPrincipalRules): void {
    this._rules = rules;
    this._done = {};
  }

  /**
   * Maps the given keys (unknown keys map to nothing), registers the found logins in `tokens` and warns once
   * per log about principals not found or replaced by the fallback user.
   */
  public async map(keys: string[], tokens: TokenContext, log: Logger, signal?: AbortSignal): Promise<IPrincipalMapping[]> {
    const wanted = keys.filter((k, i) => keys.indexOf(k) === i);
    const todo = wanted.filter((k) => !this._done[k]);
    const results = await limitConcurrency(todo.map((key) => () => this._map(key)), 4, signal);
    results.forEach((r, i) => {
      if (!r.ok && isAbortError(r.error)) throw r.error;
      this._done[todo[i]] = r.ok ? r.value : { key: todo[i] };
    });
    const warned = this._warned.get(log) || [];
    this._warned.set(log, warned);
    return wanted.map((key) => {
      const mapping = this._done[key];
      const p = this._byKey[key];
      const name = p ? p.displayName || p.email || p.loginName : key;
      if (mapping.login) tokens.set('principal', key, mapping.login);
      if ((!mapping.login || mapping.strategy === 'fallback') && warned.indexOf(key) < 0) {
        warned.push(key);
        if (mapping.login) {
          log.warn(`${name} was not found on the target site; the fallback user stands in.`, { code: 'PRINCIPAL_FALLBACK', detail: { key, login: mapping.login } });
        } else {
          log.warn(`User or group not found on the target site: ${name}. Its values are left empty.`, {
            code: 'PRINCIPAL_NOT_FOUND',
            detail: p ? { key: p.key, email: p.email } : { key }
          });
        }
      }
      return mapping;
    });
  }

  private async _map(key: string): Promise<IPrincipalMapping> {
    const p = this._byKey[key];
    if (!p) return { key };
    const rules = this._rules;
    const csv = rules.csv || {};
    const csvTarget = [p.email, p.upn, p.loginName].map((v) => (v ? csv[v.toLowerCase()] : undefined)).filter((v) => !!v)[0];
    const domains = rules.domains || [];
    const attempts: Array<[PrincipalStrategy, string | undefined]> = [
      ['manual', rules.manual ? rules.manual[key] : undefined],
      ['csv', csvTarget],
      ['sameLogin', p.loginName],
      ['sameEmail', p.email],
      ['domain', domains.length ? (p.email && replaceDomain(p.email, domains)) || (p.loginName && replaceDomain(p.loginName, domains)) || undefined : undefined],
      ['fallback', rules.fallback]
    ];
    for (const [strategy, value] of attempts) {
      if (!value) continue;
      const login = await this._ensure(value);
      if (login) return { key, login, strategy };
    }
    return { key };
  }

  /** The target login for an e-mail / login, or undefined when ensureUser does not know it (cached). */
  private _ensure(value: string): Promise<string | undefined> {
    const k = value.toLowerCase();
    if (!this._ensured[k]) {
      this._ensured[k] = this._sp.web.ensureUser(value).then(
        (user) => (user && user.LoginName ? user.LoginName : undefined),
        (e: unknown) => {
          if (isAbortError(e)) {
            delete this._ensured[k];
            throw e;
          }
          return undefined;
        }
      );
    }
    return this._ensured[k];
  }
}
