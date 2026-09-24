import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { AddFieldOptions } from '@pnp/sp/fields';
import { CopyJetError, throwIfAborted } from '../errors';
import { compareFields, fieldUpdate, resolveFieldDef, sanitizeFieldXml, toFieldDef, type IFieldInfoLike } from '../fields';
import { isHttpStatus } from '../http/status';
import { LIST_FIELD_SELECT, listFieldKey, toServerRelativeUrl, type IListFieldDef } from '../lists';
import type { ConflictMode, IApplyResult, IDiffResult, IInstallContext, IProvider } from '../model';
import { resolve } from '../tokenizer';

/** Keep the internal name and add the column to all list content types, so it shows in the forms (= 12). */
const ADD_OPTIONS = AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddToAllContentTypes;

interface IListFieldDiff extends IDiffResult {
  target?: IFieldInfoLike;
}

const odataString = (s: string): string => s.replace(/'/g, "''");

function isTaxonomy(type: string): boolean {
  return type === 'TaxonomyFieldType' || type === 'TaxonomyFieldTypeMulti';
}

export class ListFieldProvider implements IProvider<IListFieldDef> {
  public readonly kind = 'listField' as const;

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
      // Taxonomy columns need term store mapping (phase 2).
      return isTaxonomy(def.field.type) ? { ref, status: 'unsupported', changes: ['taxonomy'] } : { ref, status: 'new' };
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
        return this._create(sp, def, ctx);
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
    ctx.log.info(siteColumn ? 'Site column added to the list.' : 'List column created.', { artifact: ref });
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
