import type { TokenContext } from './TokenContext';

/**
 * How tokenize() recognizes a value:
 * - guid: bare GUID, case-insensitive; surrounding braces / %7B..%7D are left in place
 * - url:  URL or server-relative path, case-insensitive, only at path-segment boundaries
 * - text: exact text (only when autoTokenize is on)
 */
export type TokenValueKind = 'guid' | 'url' | 'text';

export interface ITokenDefinition {
  /** Lowercase letters only; the token is written as {name} or {name:arg}. */
  name: string;
  hasArgument: boolean;
  kind: TokenValueKind;
  /** Whether tokenize() replaces occurrences of the value. Off for values too generic to replace blindly. */
  autoTokenize: boolean;
  /** Custom resolution; by default the value stored in the TokenContext is used. */
  resolve?: (arg: string | undefined, ctx: TokenContext) => string | undefined;
}

const registry: { [name: string]: ITokenDefinition } = {};

export function registerToken(def: ITokenDefinition): void {
  if (!/^[a-z]+$/.test(def.name)) {
    throw new Error(`Invalid token name "${def.name}": lowercase letters only.`);
  }
  registry[def.name] = def;
}

export function getTokenDefinition(name: string): ITokenDefinition | undefined {
  return Object.prototype.hasOwnProperty.call(registry, name) ? registry[name] : undefined;
}

// Built-in tokens (phase 0). Providers register further ones (viewid, groupkey, principal ...) as they come.
registerToken({ name: 'site', hasArgument: false, kind: 'url', autoTokenize: true });
registerToken({ name: 'siterelative', hasArgument: false, kind: 'url', autoTokenize: true });
registerToken({ name: 'sitename', hasArgument: false, kind: 'text', autoTokenize: false });
registerToken({ name: 'listkey', hasArgument: true, kind: 'guid', autoTokenize: true });
// Site-relative list URL (e.g. Lists/Projektek); absolute list URLs are already covered by {site}.
registerToken({ name: 'listurl', hasArgument: true, kind: 'text', autoTokenize: false });
registerToken({ name: 'fieldid', hasArgument: true, kind: 'guid', autoTokenize: true });
// SharePoint groups: values are group IDs on the target. Never auto-tokenized (plain numbers).
registerToken({ name: 'groupkey', hasArgument: true, kind: 'text', autoTokenize: false });
registerToken({ name: 'associatedownergroup', hasArgument: false, kind: 'text', autoTokenize: false });
registerToken({ name: 'associatedmembergroup', hasArgument: false, kind: 'text', autoTokenize: false });
registerToken({ name: 'associatedvisitorgroup', hasArgument: false, kind: 'text', autoTokenize: false });
// People in item values and system fields: {principal:key} → the target login name, set by the PrincipalMapper.
registerToken({ name: 'principal', hasArgument: true, kind: 'text', autoTokenize: false });
