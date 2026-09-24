import { CopyJetError } from '../errors';
import { getTokenDefinition } from './registry';
import type { TokenContext } from './TokenContext';

const TOKEN_PATTERN = /\{([a-z]+)(?::([^{}]+))?\}/g;

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function tokenText(name: string, arg: string | undefined): string {
  return arg === undefined ? `{${name}}` : `{${name}:${arg}}`;
}

interface ICandidate {
  value: string;
  token: string;
  pattern: string;
}

/**
 * Lookarounds keep matches on whole values: a GUID not inside a longer hex run, a URL not continuing
 * into a longer path segment (/sites/a vs /sites/ab) and not preceded by a host (other.com/sites/a).
 * Built with new RegExp: lookbehind is ES2018 and the build targets ES5 syntax checks.
 */
function candidatePattern(kind: string, value: string): string {
  const v = escapeRegExp(value);
  if (kind === 'guid') {
    // A percent-escape such as %7B ends in a hex digit but does not extend the GUID.
    return `(?:(?<![0-9a-fA-F])|(?<=%[0-9a-fA-F]{2}))${v}(?![0-9a-fA-F])`;
  }
  if (kind === 'url') {
    return `(?<![\\w.-])${v}(?![\\w-])`;
  }
  return v;
}

function buildMatcher(ctx: TokenContext): { regex: RegExp; byValue: { [lower: string]: string } } | undefined {
  const candidates: ICandidate[] = [];
  ctx.entries().forEach((e) => {
    const def = getTokenDefinition(e.name);
    // A root site's server-relative URL is "/", which would match every slash.
    if (!def || !def.autoTokenize || !e.value || e.value === '/') {
      return;
    }
    candidates.push({ value: e.value, token: tokenText(e.name, e.arg), pattern: candidatePattern(def.kind, e.value) });
  });
  if (candidates.length === 0) {
    return undefined;
  }
  // Longest value first: at a given position the regex takes the first alternative that matches.
  candidates.sort((a, b) => b.value.length - a.value.length);
  const byValue: { [lower: string]: string } = {};
  candidates.forEach((c) => {
    const k = c.value.toLowerCase();
    if (!(k in byValue)) {
      byValue[k] = c.token;
    }
  });
  // eslint-disable-next-line @rushstack/security/no-unsafe-regexp -- built from escapeRegExp()-escaped values
  const regex = new RegExp(candidates.map((c) => `(?:${c.pattern})`).join('|'), 'gi');
  return { regex, byValue };
}

/** Replaces known source-site values (URLs, list/field GUIDs ...) with tokens; the longest match wins. */
export function tokenize(value: string, ctx: TokenContext): string {
  const m = buildMatcher(ctx);
  return m ? value.replace(m.regex, (found) => m.byValue[found.toLowerCase()]) : value;
}

/**
 * Replaces tokens with target-site values. Throws CopyJetError TOKEN_UNKNOWN for an unregistered token
 * and TOKEN_UNRESOLVED when the context has no value – a template is never installed half-resolved.
 */
export function resolve(value: string, ctx: TokenContext): string {
  return value.replace(TOKEN_PATTERN, (token: string, name: string, arg: string | undefined) => {
    const def = getTokenDefinition(name);
    if (!def) {
      throw new CopyJetError('TOKEN_UNKNOWN', `Unknown token ${token}.`, { token });
    }
    if (def.hasArgument !== (arg !== undefined)) {
      throw new CopyJetError('TOKEN_INVALID', `Token ${token} ${def.hasArgument ? 'needs an argument' : 'takes no argument'}.`, { token });
    }
    const resolved = def.resolve ? def.resolve(arg, ctx) : ctx.get(name, arg);
    if (resolved === undefined) {
      throw new CopyJetError('TOKEN_UNRESOLVED', `No value for token ${token}.`, { token });
    }
    return resolved;
  });
}

/** Applies `fn` to every string in a JSON-like value (object keys are left alone). Returns a copy. */
function deepMap(value: unknown, fn: (s: string) => string): unknown {
  if (typeof value === 'string') {
    return fn(value);
  }
  if (Array.isArray(value)) {
    return value.map((v) => deepMap(v, fn));
  }
  if (value !== null && typeof value === 'object') {
    const out: { [k: string]: unknown } = {};
    Object.keys(value as object).forEach((k) => {
      out[k] = deepMap((value as { [k: string]: unknown })[k], fn);
    });
    return out;
  }
  return value;
}

export function deepTokenize<T>(value: T, ctx: TokenContext): T {
  const m = buildMatcher(ctx);
  return m ? (deepMap(value, (s) => s.replace(m.regex, (found) => m.byValue[found.toLowerCase()])) as T) : value;
}

export function deepResolve<T>(value: T, ctx: TokenContext): T {
  return deepMap(value, (s) => resolve(s, ctx)) as T;
}
