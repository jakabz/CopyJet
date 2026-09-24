import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import { throwIfAborted } from '../errors';
import { templateFieldFrom } from '../fields';
import { limitConcurrency } from '../http/concurrency';
import {
  LIST_FIELD_SELECT,
  isCopiedListField,
  isSupportedTemplate,
  listFieldKey,
  loadSourceSite,
  type IListFieldDef,
  type IListFieldInfoLike
} from '../lists';
import { listFieldDependencies } from '../planner/dependencies';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, ITemplateWriter } from '../model';

function listKeyOf(lookupList: string | undefined): string | undefined {
  const m = lookupList ? /\{listkey:([^{}]+)\}/.exec(lookupList) : null;
  return m ? m[1] : undefined;
}

/**
 * List columns: the list's own columns (full definition) and site column instances (same ID as the site
 * column, spike 05). Written into the template's list entry, which the ListExtractor must have created.
 */
export class ListFieldExtractor implements IExtractor<IListFieldDef> {
  public readonly kind = 'listField' as const;

  public async discover(sp: SPFI, signal?: AbortSignal): Promise<IDiscoveredArtifact[]> {
    const site = await loadSourceSite(sp, signal);
    const lists = site.lists.filter((l) => isSupportedTemplate(l.info.BaseTemplate));
    const results = await limitConcurrency(lists.map((l) => () => this._fields(sp, l.info.RootFolder.ServerRelativeUrl)), 4, signal);
    const out: IDiscoveredArtifact[] = [];
    results.forEach((r, i) => {
      if (!r.ok) return;
      r.value
        .filter((f) => isCopiedListField(f, lists[i].info.Id))
        .forEach((f) => out.push({ ref: { kind: this.kind, key: listFieldKey(lists[i].key, f.InternalName) }, title: f.Title, parentKey: `list:${lists[i].key}` }));
    });
    return out;
  }

  public dependencies(def: IListFieldDef): IArtifactRef[] {
    return listFieldDependencies(def);
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) {
      return;
    }
    const site = await loadSourceSite(sp, opts.signal);
    const lists = site.lists.filter((l) => wanted.some((k) => k.indexOf(`listField:${l.key}/`) === 0));
    const found: string[] = [];

    for (const l of lists) {
      throwIfAborted(opts.signal);
      const listDef = out.manifest.lists.filter((x) => x.key === l.key)[0];
      if (!listDef) {
        opts.log.warn(`List ${l.info.Title} is not in the template; its columns are skipped.`, {
          artifact: { kind: 'list', key: `list:${l.key}` },
          code: 'LISTFIELD_LIST_MISSING'
        });
        continue;
      }
      const fields = (await this._fields(sp, l.info.RootFolder.ServerRelativeUrl)).filter((f) => isCopiedListField(f, l.info.Id));
      fields
        .filter((f) => wanted.indexOf(listFieldKey(l.key, f.InternalName)) >= 0)
        .forEach((f) => {
          const ref: IArtifactRef = { kind: this.kind, key: listFieldKey(l.key, f.InternalName) };
          found.push(ref.key);
          const def = templateFieldFrom(f, opts.tokens, opts.log, ref);
          const target = listKeyOf(def.lookupList);
          if ((def.type === 'Lookup' || def.type === 'LookupMulti') && !target) {
            opts.log.warn(`Lookup column ${def.internalName} points to a list outside the site.`, { artifact: ref, code: 'LOOKUP_TARGET_UNKNOWN' });
          }
          const listFields = (listDef.fields = listDef.fields || []);
          const at = listFields.findIndex((x) => x.internalName === def.internalName);
          if (at >= 0) {
            listFields[at] = def;
          } else {
            listFields.push(def);
          }
        });
    }

    wanted
      .filter((key) => found.indexOf(key) < 0)
      .forEach((key) => opts.log.warn('List column not found on the source site.', { artifact: { kind: this.kind, key }, code: 'LISTFIELD_NOT_FOUND' }));
  }

  private _fields(sp: SPFI, listServerRelativeUrl: string): Promise<IListFieldInfoLike[]> {
    return sp.web.getList(listServerRelativeUrl).fields.select(...LIST_FIELD_SELECT)<IListFieldInfoLike[]>();
  }
}
