import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { AddFieldOptions } from '@pnp/sp/fields';
import { CopyJetError, throwIfAborted } from '../errors';
import { compareFields, fieldUpdate, isTaxonomyType, resolveFieldDef, sanitizeFieldXml, taxonomyFieldXml, toFieldDef, type IFieldInfoLike } from '../fields';
import { csomAddExistingField, csomCreateTaxonomyField, type FetchLike } from '../http/csom';
import { resolveTargetTermSet, termStoreFor, type ITargetTermSet } from '../taxonomy';
import { isHttpStatus } from '../http/status';
import { LIST_FIELD_SELECT, listFieldKey, toServerRelativeUrl, type IListFieldDef } from '../lists';
import type { ConflictMode, IApplyResult, IDiffResult, IInstallContext, IProvider } from '../model';
import { resolve } from '../tokenizer';

/** Keep the internal name and add the column to all list content types, so it shows in the forms (= 12). */
const ADD_OPTIONS = AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddToAllContentTypes;

interface IListFieldDiff extends IDiffResult {
  target?: IFieldInfoLike;
  /** A new Managed Metadata column: the site column to add, or the term set a list column binds to. */
  taxonomy?: { siteColumnId?: string; termSet?: ITargetTermSet };
}

const odataString = (s: string): string => s.replace(/'/g, "''");

export class ListFieldProvider implements IProvider<IListFieldDef> {
  public readonly kind = 'listField' as const;
  private readonly _fetch?: FetchLike;

  /** `fetchImpl` is used for the CSOM requests of Managed Metadata columns (injected in tests). */
  constructor(fetchImpl?: FetchLike) {
    this._fetch = fetchImpl;
  }

  public async diff(sp: SPFI, def: IListFieldDef, ctx: IInstallContext): Promise<IListFieldDiff> {
    throwIfAborted(ctx.signal);
    const ref = { kind: this.kind, key: listFieldKey(def.listKey, def.field.internalName) };
    const listUrl = this._listUrl(def, ctx);
    if (!(await this._listExists(sp, listUrl))) {
      return { ref, status: 'unsupported', changes: ['listMissing'] };
    }
    const fields = sp.web.getList(listUrl).fields;
    const byName = await fields.filter(`InternalName eq '${odataString(def.field.internalName)}'`).select(...LIST_FIELD_SELECT)<IFieldInfoLike[]>();
    throwIfAborted(ctx.signal);

    if (byName.length === 0) {
      if (def.field.id) {
        const byId = await fields.filter(`Id eq guid'${def.field.id}'`).select('Id')<Array<{ Id: string }>>();
        if (byId.length) return { ref, status: 'unsupported', changes: ['idConflict'] };
      }
      if (isTaxonomyType(def.field.type)) return this._taxonomyDiff(sp, def, ref, ctx);
      return { ref, status: 'new' };
    }
    const target = byName[0];
    const changes = compareFields(resolveFieldDef(def.field, ctx.tokens), toFieldDef(target, target.SchemaXml));
    return { ref, status: changes.length ? 'different' : 'same', changes: changes.length ? changes : undefined, target };
  }

  public async apply(sp: SPFI, def: IListFieldDef, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;

    switch (diff.status) {
      case 'new':
        return diff.taxonomy ? this._createTaxonomy(sp, def, diff.taxonomy, ctx) : this._create(sp, def, ctx);
      case 'same':
        ctx.log.info('List column already present; skipped.', { artifact: ref });
        return { ref, outcome: 'skipped' };
      case 'different':
        if (mode === 'update') {
          return this._update(sp, def, diff, ctx);
        }
        if (mode === 'rename') {
          ctx.log.warn('List columns cannot be installed renamed; the existing column is kept.', { artifact: ref, code: 'LISTFIELD_RENAME_UNSUPPORTED' });
        }
        ctx.log.info(`List column differs (${(diff.changes || []).join(', ')}); kept as it is.`, { artifact: ref });
        return { ref, outcome: 'skipped' };
      default:
        ctx.log.warn(`List column skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'LISTFIELD_UNSUPPORTED', detail: diff.changes });
        return { ref, outcome: 'skipped' };
    }
  }

  /**
   * A site column instance is created from the target site column's own SchemaXml (keeps the ID and the
   * SourceID link, spike 05); the list's own columns from the template's sanitized, resolved SchemaXml.
   */
  private async _create(sp: SPFI, def: IListFieldDef, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = { kind: this.kind, key: listFieldKey(def.listKey, def.field.internalName) };
    const siteColumn = def.field.id
      ? (await sp.web.availablefields.filter(`Id eq guid'${def.field.id}'`).select('SchemaXml')<Array<{ SchemaXml: string }>>())[0]
      : undefined;
    let schemaXml: string;
    if (siteColumn) {
      schemaXml = siteColumn.SchemaXml;
    } else {
      const sanitized = sanitizeFieldXml(resolve(def.field.schemaXml, ctx.tokens));
      if (sanitized.removedAttributes.length || sanitized.removedElements.length) {
        ctx.log.warn('Removed disallowed parts of SchemaXml before install.', {
          artifact: ref,
          code: 'FIELD_XML_SANITIZED',
          detail: { attributes: sanitized.removedAttributes, elements: sanitized.removedElements }
        });
      }
      schemaXml = sanitized.xml;
    }
    const created = await sp.web.getList(this._listUrl(def, ctx)).fields.createFieldAsXml({ SchemaXml: schemaXml, Options: ADD_OPTIONS });
    if (!created.Id) {
      throw new CopyJetError('LISTFIELD_CREATE_FAILED', `SharePoint returned no Id for ${def.field.internalName}.`);
    }
    if (created.InternalName && created.InternalName !== def.field.internalName) {
      ctx.log.warn(`Column was created as ${created.InternalName}.`, { artifact: ref, code: 'FIELD_NAME_CHANGED' });
    }
    // SchemaXml carries the site's default-language DisplayName; the template title is what the user saw.
    const title = resolve(def.field.title, ctx.tokens);
    if (created.Title !== undefined && created.Title !== title) {
      await sp.web.getList(this._listUrl(def, ctx)).fields.getById(created.Id).update({ Title: title }, 'SP.Field');
    }
    ctx.log.info(siteColumn ? 'Site column added to the list.' : 'List column created.', { artifact: ref });
    return { ref, outcome: 'created' };
  }

  /**
   * A new Managed Metadata column: an instance of a target site column with the same ID is added as it is;
   * a list column needs the term set on the target (by ID, else by "Group/Set" path, spike 11).
   */
  private async _taxonomyDiff(sp: SPFI, def: IListFieldDef, ref: IListFieldDiff['ref'], ctx: IInstallContext): Promise<IListFieldDiff> {
    if (def.field.id) {
      const site = await sp.web.availablefields.filter(`Id eq guid'${def.field.id}'`).select('Id')<Array<{ Id: string }>>();
      if (site.length) return { ref, status: 'new', taxonomy: { siteColumnId: def.field.id } };
    }
    const termSet = def.field.termSet ? await resolveTargetTermSet(termStoreFor(sp, ctx), def.field.termSet) : undefined;
    return termSet ? { ref, status: 'new', taxonomy: { termSet } } : { ref, status: 'unsupported', changes: ['termSetMissing'] };
  }

  /**
   * Managed Metadata through CSOM only: a SchemaXml with its term store binding fails with HTTP 500 and leaves
   * the list unusable (spike 11 B/C).
   */
  private async _createTaxonomy(sp: SPFI, def: IListFieldDef, taxonomy: NonNullable<IListFieldDiff['taxonomy']>, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = { kind: this.kind, key: listFieldKey(def.listKey, def.field.internalName) };
    const listUrl = this._listUrl(def, ctx);
    const title = resolve(def.field.title, ctx.tokens);
    if (taxonomy.siteColumnId) {
      await csomAddExistingField(sp, listUrl, taxonomy.siteColumnId, ctx.signal, this._fetch);
      ctx.log.info('Site column added to the list (Managed Metadata).', { artifact: ref });
      return { ref, outcome: 'created' };
    }
    const termSet = taxonomy.termSet!;
    const created = await csomCreateTaxonomyField(
      sp,
      listUrl,
      {
        schemaXml: taxonomyFieldXml(def.field, title),
        termStoreId: termSet.termStoreId,
        termSetId: termSet.termSetId,
        anchorId: def.field.termSet && def.field.termSet.anchorId,
        open: def.field.termSet && def.field.termSet.isOpen
      },
      ADD_OPTIONS,
      ctx.signal,
      this._fetch
    );
    if (created.internalName && created.internalName !== def.field.internalName) {
      ctx.log.warn(`Column was created as ${created.internalName}.`, { artifact: ref, code: 'FIELD_NAME_CHANGED' });
    }
    ctx.log.info(`List column created (Managed Metadata, term set found by ${termSet.matchedBy === 'id' ? 'ID' : 'path'}).`, { artifact: ref });
    return { ref, outcome: 'created' };
  }

  private async _update(sp: SPFI, def: IListFieldDef, diff: IListFieldDiff, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = diff.ref;
    const target = diff.target!;
    const update = fieldUpdate(resolveFieldDef(def.field, ctx.tokens), target, diff.changes || []);
    if (update.notUpdatable.length) {
      ctx.log.warn(`Not updatable on an existing column: ${update.notUpdatable.join(', ')}.`, { artifact: ref, code: 'FIELD_PARTIAL_UPDATE', detail: update.notUpdatable });
    }
    if (Object.keys(update.props).length === 0) {
      return { ref, outcome: 'skipped' };
    }
    await sp.web.getList(this._listUrl(def, ctx)).fields.getById(target.Id).update(update.props, update.fieldType);
    ctx.log.info(`List column updated: ${Object.keys(update.props).join(', ')}.`, { artifact: ref });
    return { ref, outcome: 'updated' };
  }

  /** Server-relative URL of the installed list – the renamed copy's URL when the ListProvider made one. */
  private _listUrl(def: IListFieldDef, ctx: IInstallContext): string {
    const web = ctx.tokens.get('siterelative');
    if (web === undefined) {
      throw new CopyJetError('TOKEN_UNRESOLVED', 'The install context has no {siterelative} value.');
    }
    return toServerRelativeUrl(ctx.tokens.get('listurl', def.listKey) || def.listUrl, web);
  }

  private async _listExists(sp: SPFI, listUrl: string): Promise<boolean> {
    try {
      await sp.web.getList(listUrl).select('Id')();
      return true;
    } catch (e) {
      if (isHttpStatus(e, 404)) return false;
      throw e;
    }
  }
}
