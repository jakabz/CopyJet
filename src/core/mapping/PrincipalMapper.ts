import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import { isAbortError } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import type { Logger } from '../logger/Logger';
import type { IPrincipal } from '../model';
import type { TokenContext } from '../tokenizer';

export type PrincipalStrategy = 'sameLogin' | 'sameEmail';

export interface IPrincipalMapping {
  key: string;
  /** Target login name; undefined when the principal was not found on the target. */
  login?: string;
  strategy?: PrincipalStrategy;
}

/**
 * Template principals → target users (rendszerterv §5, first strategies): the same login (same tenant),
 * then the same e-mail (web.ensureUser accepts both). Found ones become {principal:key} values in the
 * token context; each principal is looked up once per mapper. Domain replacement, CSV and a fallback
 * user come with the MappingStep.
 */
export class PrincipalMapper {
  private readonly _sp: SPFI;
  private readonly _byKey: { [key: string]: IPrincipal } = {};
  private readonly _done: { [key: string]: IPrincipalMapping } = {};

  constructor(sp: SPFI, principals: IPrincipal[]) {
    this._sp = sp;
    principals.forEach((p) => (this._byKey[p.key] = p));
  }

  /** Maps the given keys (unknown keys map to nothing) and registers the logins in `tokens`. */
  public async map(keys: string[], tokens: TokenContext, log: Logger, signal?: AbortSignal): Promise<IPrincipalMapping[]> {
    const todo = keys.filter((k, i) => keys.indexOf(k) === i && !this._done[k]);
    const results = await limitConcurrency(todo.map((key) => () => this._map(key)), 4, signal);
    results.forEach((r, i) => {
      if (!r.ok && isAbortError(r.error)) throw r.error;
      const mapping = r.ok ? r.value : { key: todo[i] };
      this._done[mapping.key] = mapping;
      if (mapping.login) {
        tokens.set('principal', mapping.key, mapping.login);
      } else {
        const p = this._byKey[mapping.key];
        log.warn(`User or group not found on the target site: ${p ? p.displayName || p.email || p.loginName : mapping.key}. Its values are left empty.`, {
          code: 'PRINCIPAL_NOT_FOUND',
          detail: p ? { key: p.key, email: p.email } : { key: mapping.key }
        });
      }
    });
    return keys.map((k) => this._done[k]).filter((m) => !!m);
  }

  private async _map(key: string): Promise<IPrincipalMapping> {
    const p = this._byKey[key];
    if (!p) return { key };
    const attempts: Array<[PrincipalStrategy, string | undefined]> = [
      ['sameLogin', p.loginName],
      ['sameEmail', p.email]
    ];
    for (const [strategy, value] of attempts) {
      if (!value) continue;
      try {
        const user = await this._sp.web.ensureUser(value);
        if (user && user.LoginName) return { key, login: user.LoginName, strategy };
      } catch (e) {
        if (isAbortError(e)) throw e;
        // not resolvable this way; try the next strategy
      }
    }
    return { key };
  }
}
