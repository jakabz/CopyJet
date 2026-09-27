import type { FieldValue } from '../model';
import type { TokenContext } from '../tokenizer';
import { resolveKnown, tokenize } from '../tokenizer';
import { formatSpDateTime, formatSpNumber, type IWebLocale } from './locale';

/**
 * FieldValueSerializer: item values between the source REST shape, the template and the strings
 * AddValidateUpdateItemUsingPath / ValidateUpdateListItem accept (spike 08).
 *
 * Template shapes (language-independent): string, number, boolean, ISO UTC date, string[] (MultiChoice),
 * { lookup: [sourceIds] }, { principals: ['{principal:key}'] }, { url, description }.
 */

export type ItemFieldKind =
  | 'text'
  | 'note'
  | 'number'
  | 'boolean'
  | 'choice'
  | 'multiChoice'
  | 'dateTime'
  | 'url'
  | 'user'
  | 'userMulti'
  | 'lookup'
  | 'lookupMulti';

const KINDS: { [typeAsString: string]: ItemFieldKind } = {
  Text: 'text',
  Note: 'note',
  Number: 'number',
  Currency: 'number',
  Integer: 'number',
  Boolean: 'boolean',
  Choice: 'choice',
  MultiChoice: 'multiChoice',
  DateTime: 'dateTime',
  URL: 'url',
  User: 'user',
  UserMulti: 'userMulti',
  Lookup: 'lookup',
  LookupMulti: 'lookupMulti'
};

/** Types whose value CopyJet does not copy yet (warned once per column); everything else unknown is ignored. */
export const NOT_YET_COPIED_TYPES = ['TaxonomyFieldType', 'TaxonomyFieldTypeMulti', 'Thumbnail', 'Location', 'Geolocation'];

/** Writable columns that are not item content (or are handled separately: content type, system values). */
const SKIPPED_FIELDS = ['ContentType', 'Attachments', '_ColorTag', 'ComplianceAssetId', 'Author', 'Editor', 'Created', 'Modified', 'FileLeafRef'];

export interface IItemField {
  internalName: string;
  typeAsString: string;
  kind: ItemFieldKind;
}

export function itemFieldKind(typeAsString: string): ItemFieldKind | undefined {
  return Object.prototype.hasOwnProperty.call(KINDS, typeAsString) ? KINDS[typeAsString] : undefined;
}

/** A list column as REST returns it → the item field CopyJet copies, or undefined. */
export function toItemField(f: { InternalName: string; TypeAsString: string; Hidden?: boolean; ReadOnlyField?: boolean }): IItemField | undefined {
  if (f.Hidden || f.ReadOnlyField || SKIPPED_FIELDS.indexOf(f.InternalName) >= 0) return undefined;
  const kind = itemFieldKind(f.TypeAsString);
  return kind ? { internalName: f.InternalName, typeAsString: f.TypeAsString, kind } : undefined;
}

export const isLookupKind = (kind: ItemFieldKind): boolean => kind === 'lookup' || kind === 'lookupMulti';
export const isUserKind = (kind: ItemFieldKind): boolean => kind === 'user' || kind === 'userMulti';

/** The REST property holding the value: person and lookup columns are read by ID (no $expand, no lookup threshold). */
export function restPropertyOf(field: IItemField): string {
  return isLookupKind(field.kind) || isUserKind(field.kind) ? `${field.internalName}Id` : field.internalName;
}

/** Multi-value REST values: a plain array (nometadata) or { results: [] } (verbose). */
function asArray<T>(raw: unknown): T[] {
  if (Array.isArray(raw)) return raw as T[];
  if (raw && typeof raw === 'object' && Array.isArray((raw as { results?: unknown }).results)) return (raw as { results: T[] }).results;
  return raw === null || raw === undefined ? [] : [raw as T];
}

export interface IToTemplateContext {
  tokens: TokenContext;
  /** Site user ID → '{principal:key}', or undefined when the principal cannot be carried (SharePoint group …). */
  principal(userId: number): string | undefined;
}

/** A REST item value → template value; undefined means "nothing to copy" (empty). */
export function toTemplateValue(field: IItemField, raw: unknown, ctx: IToTemplateContext): FieldValue | undefined {
  if (raw === null || raw === undefined) return undefined;
  switch (field.kind) {
    case 'text':
    case 'choice':
      return String(raw) === '' ? undefined : String(raw);
    case 'note':
      // Rich text may link to the source site.
      return String(raw) === '' ? undefined : tokenize(String(raw), ctx.tokens);
    case 'number':
      return typeof raw === 'number' ? raw : Number(raw);
    case 'boolean':
      return !!raw;
    case 'dateTime':
      return String(raw);
    case 'multiChoice': {
      const values = asArray<string>(raw).map(String);
      return values.length ? values : undefined;
    }
    case 'url': {
      const link = raw as { Url?: string; Description?: string };
      if (!link.Url) return undefined;
      const out: { url: string; description?: string } = { url: tokenize(link.Url, ctx.tokens) };
      if (link.Description && link.Description !== link.Url) out.description = link.Description;
      return out;
    }
    case 'user':
    case 'userMulti': {
      const principals = asArray<number>(raw)
        .map((id) => ctx.principal(Number(id)))
        .filter((p): p is string => !!p);
      return principals.length ? { principals } : undefined;
    }
    case 'lookup':
    case 'lookupMulti': {
      const ids = asArray<number>(raw).map(Number).filter((id) => id > 0);
      return ids.length ? { lookup: ids } : undefined;
    }
    default:
      return undefined;
  }
}

export interface IToTargetContext {
  locale: IWebLocale;
  tokens: TokenContext;
  /** UTC ISO → the target web's local wall-clock time ("2026-03-02T09:15:00"). */
  localTime(utcIso: string): string;
  /** Source lookup IDs → target IDs; unknown ones are dropped. */
  lookupIds?(field: IItemField, sourceIds: number[]): number[];
}

/** Target lookup IDs → form value. Multi lookups need the empty value part: "1;#;#2" ("1;#2" silently keeps only the first, spike 08 E). */
export function lookupValue(field: IItemField, targetIds: number[]): string | undefined {
  if (!targetIds.length) return undefined;
  return field.kind === 'lookup' ? String(targetIds[0]) : targetIds.join(';#;#');
}

/** Login names of '{principal:key}' tokens; unresolved ones (not on the target) are dropped. */
export function principalLogins(tokens: string[], ctx: TokenContext): string[] {
  return tokens.map((t) => resolveKnown(t, ctx)).filter((login) => !/^\{[a-z]+(:[^{}]+)?\}$/.test(login));
}

/** People picker value for ValidateUpdateListItem. */
export function peopleValue(logins: string[]): string {
  return JSON.stringify(logins.map((Key) => ({ Key })));
}

/**
 * A template value → the string AddValidateUpdateItemUsingPath accepts on the target web, or undefined when
 * nothing should be sent (empty, or every person / lookup unresolved).
 */
export function toTargetValue(field: IItemField, value: FieldValue, ctx: IToTargetContext): string | undefined {
  if (value === null || value === undefined) return undefined;
  switch (field.kind) {
    case 'text':
    case 'choice':
      return String(value);
    case 'note':
      return resolveKnown(String(value), ctx.tokens);
    case 'number':
      return typeof value === 'number' ? formatSpNumber(value, ctx.locale) : undefined;
    case 'boolean':
      return value ? '1' : '0';
    case 'dateTime':
      return typeof value === 'string' ? formatSpDateTime(ctx.localTime(value), ctx.locale) : undefined;
    case 'multiChoice': {
      const values = Array.isArray(value) ? value.map(String) : [String(value)];
      return values.length ? `;#${values.join(';#')};#` : undefined;
    }
    case 'url': {
      const link = value as { url?: string; description?: string };
      if (!link.url) return undefined;
      // A comma inside the URL is doubled; the first single ", " separates the description (spike 08 E).
      // Without a description SharePoint shows the URL itself, as on the source.
      const url = resolveKnown(link.url, ctx.tokens);
      return `${url.replace(/,/g, ',,')}, ${link.description === undefined ? url : link.description}`;
    }
    case 'user':
    case 'userMulti': {
      const tokens = (value as { principals?: string[] }).principals || [];
      const logins = principalLogins(tokens, ctx.tokens);
      if (!logins.length) return undefined;
      return peopleValue(field.kind === 'user' ? logins.slice(0, 1) : logins);
    }
    case 'lookup':
    case 'lookupMulti': {
      const ids = (value as { lookup?: number[] }).lookup || [];
      return lookupValue(field, ctx.lookupIds ? ctx.lookupIds(field, ids) : []);
    }
    default:
      return undefined;
  }
}
