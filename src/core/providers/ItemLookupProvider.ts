import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { CopyJetError, throwIfAborted } from '../errors';
import { isHttpStatus } from '../http/status';
import {
  dateValuesOf,
  isLookupKind,
  itemLookupsKey,
  logItemFailures,
  readTargetFields,
  readWebLocale,
  runItemWrites,
  lookupValue,
  webLocalTime,
  type IFormValue,
  type IListItemsDef,
  type ITargetField
} from '../items';
import type { ConflictMode, IApplyResult, IArtifactRef, IDiffResult, IInstallContext, IItemsFile, IProvider } from '../model';
import { contentContext, itemsListUrl, systemFormValues } from './ItemProvider';
import type { IItemProviderOptions } from './ItemProvider';

/**
 * Lookup values, second round (rendszerterv §6): runs after the items of the list and of every list its
 * lookups point to, and maps source item IDs through their ID maps. Editor and Modified are sent again so
 * the update keeps the source values.
 */
export class ItemLookupProvider implements IProvider<IListItemsDef> {
  public readonly kind = 'itemLookups' as const;
  private readonly _batched: boolean;

  constructor(options: IItemProviderOptions = {}) {
    this._batched = options.batched !== false;
  }

  public async diff(sp: SPFI, def: IListItemsDef, ctx: IInstallContext): Promise<IDiffResult> {
    throwIfAborted(ctx.signal);
    const ref: IArtifactRef = { kind: this.kind, key: itemLookupsKey(def.listKey) };
    try {
      await sp.web.getList(itemsListUrl(def, ctx)).select('Id')();
    } catch (e) {
      if (isHttpStatus(e, 404)) return { ref, status: 'unsupported', changes: ['listMissing'] };
      throw e;
    }
    return { ref, status: 'new' };
  }

  public async apply(sp: SPFI, def: IListItemsDef, _mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;
    if (diff.status !== 'new') {
      ctx.log.warn(`Lookup values skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'ITEMS_UNSUPPORTED', detail: diff.changes });
      return { ref, outcome: 'skipped' };
    }
    const content = contentContext(ctx);
    const idMap = content.idMaps[def.listKey];
    if (!idMap) {
      ctx.log.info('The items of this list were not written in this run; lookup values skipped.', { artifact: ref });
      return { ref, outcome: 'skipped' };
    }

    const file = await content.reader.getJson<IItemsFile>(def.source, ctx.signal);
    const listUrl = itemsListUrl(def, ctx);
    const fields = await readTargetFields(sp, listUrl, def.listKey, ctx.tokens);
    const unresolved: { [name: string]: boolean } = {};
    const lookupIds = (f: ITargetField, ids: number[]): number[] => {
      const map = f.lookupListKey ? content.idMaps[f.lookupListKey] : undefined;
      if (!map) {
        unresolved[f.internalName] = true;
        return [];
      }
      return ids.map((id) => map[id]).filter((id) => !!id);
    };

    // Items with at least one lookup value, their lookup form values computed up front.
    const pending = file.items
      .filter((item) => !!idMap[item.sourceId])
      .map((item) => {
        const values: IFormValue[] = [];
        Object.keys(item.values).forEach((name) => {
          const f = fields[name];
          if (!f || !isLookupKind(f.kind)) return;
          const v = lookupValue(f, lookupIds(f, (item.values[name] as { lookup?: number[] } | null)?.lookup || []));
          if (v !== undefined) values.push({ FieldName: name, FieldValue: v });
        });
        return { item, values };
      })
      .filter((p) => p.values.length > 0);
    Object.keys(unresolved).forEach((name) =>
      ctx.log.warn(`Column ${name}: the items of its lookup list were not written in this run; its values are skipped.`, {
        artifact: ref,
        code: 'ITEM_LOOKUP_UNRESOLVED',
        detail: name
      })
    );
    if (!pending.length) {
      ctx.log.info('No lookup values to set.', { artifact: ref });
      return { ref, outcome: 'skipped' };
    }

    const locale = await readWebLocale(sp);
    const times = webLocalTime(sp);
    await times.prepare(dateValuesOf(pending.map((p) => p.item), {}, 'modified'), ctx.signal);
    const ops = pending.map((p) => {
      const formValues = p.values.concat(systemFormValues(p.item, 'modified', locale, times, ctx));
      const targetId = idMap[p.item.sourceId];
      return (s: SPFI) => s.web.getList(listUrl).items.getById(targetId).validateUpdateListItem(formValues, true);
    });
    const { results } = await runItemWrites(sp, ops, this._batched, ctx.signal);
    const failures = results
      .map((r, i) => ({ sourceId: pending[i].item.sourceId, errors: r.errors || [] }))
      .filter((f) => f.errors.length > 0);
    logItemFailures(ctx, ref, failures);
    if (failures.length === pending.length) {
      throw new CopyJetError('ITEMS_FAILED', `Lookup values could not be set on any of the ${pending.length} items.`, failures.slice(0, 5));
    }
    ctx.log.info(`Lookup values set on ${pending.length - failures.length} of ${pending.length} items.`, { artifact: ref });
    return { ref, outcome: 'updated' };
  }
}
