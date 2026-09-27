import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/regional-settings/web';
import type { IListItemFormUpdateValue } from '@pnp/sp/lists';
import { CopyJetError } from '../errors';
import { runBatched, type BatchOp } from '../http/batch';
import { limitConcurrency, type Settled } from '../http/concurrency';
import type { FieldValue, IArtifactRef, IInstallContext, IItemsFile } from '../model';
import type { TokenContext } from '../tokenizer';
import { toItemField, type IItemField } from './fieldValues';
import { LocalTimeConverter } from './localTime';
import { lcidToTag, type IWebLocale } from './locale';

export type TemplateItem = IItemsFile['items'][number];

/** A writable target column; lookups know the template key of the list they point to. */
export interface ITargetField extends IItemField {
  lookupListKey?: string;
}

/** One entry of the AddValidateUpdateItemUsingPath / ValidateUpdateListItem answer. */
export type IFormResult = IListItemFormUpdateValue;

export interface IFormValue {
  FieldName: string;
  FieldValue: string;
}

/** Template key of the list with this GUID, from the {listkey:K} values of the install. */
function listKeyById(listId: string, tokens: TokenContext): string | undefined {
  const id = listId.replace(/[{}]/g, '').toLowerCase();
  const e = tokens.entries().filter((x) => x.name === 'listkey' && x.value === id)[0];
  return e ? e.arg : undefined;
}

/** Writable columns of the target list by internal name. */
export async function readTargetFields(sp: SPFI, listUrl: string, listKey: string, tokens: TokenContext): Promise<{ [name: string]: ITargetField }> {
  const infos = await sp.web
    .getList(listUrl)
    .fields.select('InternalName', 'TypeAsString', 'Hidden', 'ReadOnlyField', 'SchemaXml')<
    Array<{ InternalName: string; TypeAsString: string; Hidden: boolean; ReadOnlyField: boolean; SchemaXml: string }>
  >();
  const out: { [name: string]: ITargetField } = {};
  infos.forEach((info) => {
    const f: ITargetField | undefined = toItemField(info);
    if (!f) return;
    if (f.kind === 'lookup' || f.kind === 'lookupMulti') {
      const m = /\sList="([^"]+)"/.exec(info.SchemaXml || '');
      if (m) f.lookupListKey = m[1] === 'Self' ? listKey : listKeyById(m[1], tokens);
    }
    out[f.internalName] = f;
  });
  return out;
}

/** The target web's regional settings (spike 08 D). */
export async function readWebLocale(sp: SPFI): Promise<IWebLocale> {
  const rs = await sp.web.regionalSettings.select('LocaleId', 'DecimalSeparator', 'Time24')<{ LocaleId: number; DecimalSeparator: string; Time24: boolean }>();
  const tag = lcidToTag(rs.LocaleId);
  if (!tag) {
    throw new CopyJetError('LOCALE_UNSUPPORTED', `The target site's regional setting (LCID ${rs.LocaleId}) is not supported for item values.`, { lcid: rs.LocaleId });
  }
  return { localeId: rs.LocaleId, tag, decimalSeparator: rs.DecimalSeparator || '.', time24: !!rs.Time24 };
}

/** SharePoint's time zone conversion of the target web. */
export function webLocalTime(sp: SPFI): LocalTimeConverter {
  return new LocalTimeConverter(async (iso) => {
    const r: unknown = await sp.web.regionalSettings.timeZone.utcToLocalTime(iso);
    return typeof r === 'string' ? r : String((r as { value?: string }).value);
  });
}

/** '{principal:key}' keys used by the given values and system values. */
export function principalKeysOf(items: TemplateItem[], fields: { [name: string]: ITargetField }): string[] {
  const keys: { [k: string]: boolean } = {};
  const add = (token: string | undefined): void => {
    const m = token ? /^\{principal:([^{}]+)\}$/.exec(token) : null;
    if (m) keys[m[1]] = true;
  };
  items.forEach((item) => {
    Object.keys(item.values).forEach((name) => {
      const f = fields[name];
      const v = item.values[name] as { principals?: string[] } | null;
      if (f && (f.kind === 'user' || f.kind === 'userMulti') && v && v.principals) v.principals.forEach(add);
    });
    if (item.system) {
      add(item.system.author);
      add(item.system.editor);
    }
  });
  return Object.keys(keys);
}

/** UTC values of date columns and system dates, for LocalTimeConverter.prepare(). */
export function dateValuesOf(items: TemplateItem[], fields: { [name: string]: ITargetField }, withSystem: 'all' | 'modified'): string[] {
  const out: string[] = [];
  items.forEach((item) => {
    Object.keys(item.values).forEach((name) => {
      const f = fields[name];
      const v: FieldValue = item.values[name];
      if (f && f.kind === 'dateTime' && typeof v === 'string') out.push(v);
    });
    if (item.system) {
      if (withSystem === 'all' && item.system.created) out.push(item.system.created);
      if (item.system.modified) out.push(item.system.modified);
    }
  });
  return out;
}

export interface IWriteOutcome {
  /** Per operation: the new/updated item's ID, or the reason it failed. */
  results: Array<{ id?: number; errors?: string[] }>;
}

/**
 * Runs item writes in $batch requests of 100 (or, `batched` off, one by one with 4 in flight – the test
 * mock has no $batch). A field-level error (HasException) fails the item: SharePoint does not create it
 * (spike 08 C).
 */
export async function runItemWrites(sp: SPFI, ops: Array<BatchOp<IFormResult[]>>, batched: boolean, signal?: AbortSignal): Promise<IWriteOutcome> {
  const settled: Array<Settled<IFormResult[]>> = batched
    ? await runBatched(sp, ops, 100, signal)
    : await limitConcurrency(ops.map((op) => () => op(sp)), 4, signal);
  return {
    results: settled.map((r) => {
      if (!r.ok) return { errors: [r.error instanceof Error ? r.error.message : String(r.error)] };
      const values = r.value || [];
      const errors = values.filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage || ''}`);
      if (errors.length) return { errors };
      const id = values.filter((v) => v.FieldName === 'Id')[0];
      return { id: id && id.FieldValue ? Number(id.FieldValue) : undefined };
    })
  };
}

/** Logs at most `limit` item failures in detail and one summary line for the rest. */
export function logItemFailures(ctx: IInstallContext, ref: IArtifactRef, failures: Array<{ sourceId: number; errors: string[] }>, limit: number = 20): void {
  failures.slice(0, limit).forEach((f) =>
    ctx.log.warn(`Item ${f.sourceId} could not be written: ${f.errors.join('; ')}`, { artifact: ref, code: 'ITEM_FAILED', detail: f })
  );
  if (failures.length > limit) {
    ctx.log.warn(`${failures.length - limit} more items could not be written.`, { artifact: ref, code: 'ITEM_FAILED_MORE', detail: failures.length - limit });
  }
}
