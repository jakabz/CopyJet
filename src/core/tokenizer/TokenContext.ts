import { CopyJetError } from '../errors';
import { getTokenDefinition } from './registry';

export interface ITokenEntry {
  name: string;
  arg?: string;
  value: string;
}

export interface ISiteInfo {
  absoluteUrl: string;
  serverRelativeUrl: string;
  title: string;
}

const ARG_PATTERN = /^[^{}:]+$/;

/**
 * Token values of one site. On the source side it maps values to tokens (tokenize), on the target side
 * tokens to values (resolve); providers add the identifiers they create.
 */
export class TokenContext {
  private readonly _values: { [id: string]: ITokenEntry } = {};

  public static forSite(site: ISiteInfo): TokenContext {
    const ctx = new TokenContext();
    ctx.set('site', undefined, site.absoluteUrl);
    ctx.set('siterelative', undefined, site.serverRelativeUrl);
    ctx.set('sitename', undefined, site.title);
    return ctx;
  }

  public set(name: string, arg: string | undefined, value: string): this {
    const def = getTokenDefinition(name);
    if (!def) {
      throw new CopyJetError('TOKEN_UNKNOWN', `Unknown token "${name}".`);
    }
    if (def.hasArgument !== (arg !== undefined) || (arg !== undefined && !ARG_PATTERN.test(arg))) {
      throw new CopyJetError('TOKEN_INVALID', `Token "${name}" ${def.hasArgument ? 'needs a valid argument' : 'takes no argument'}.`);
    }
    this._values[key(name, arg)] = { name, arg, value: normalize(def.kind, value) };
    return this;
  }

  public get(name: string, arg?: string): string | undefined {
    const e = this._values[key(name, arg)];
    return e ? e.value : undefined;
  }

  public entries(): ITokenEntry[] {
    return Object.keys(this._values).map((k) => this._values[k]);
  }
}

function key(name: string, arg: string | undefined): string {
  return arg === undefined ? name : `${name}:${arg}`;
}

function normalize(kind: string, value: string): string {
  if (kind === 'guid') {
    return value.replace(/^\{|\}$/g, '').toLowerCase();
  }
  if (kind === 'url') {
    return value.length > 1 ? value.replace(/\/+$/, '') : value;
  }
  return value;
}
