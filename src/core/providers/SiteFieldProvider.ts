import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/fields';
import { CopyJetError, throwIfAborted } from '../errors';
import { compareFields, sanitizeFieldXml, toFieldDef, type IFieldInfoLike } from '../fields';
import type { ConflictMode, IApplyResult, IDiffResult, IField, IInstallContext, IProvider } from '../model';
import { resolve } from '../tokenizer';

const SELECT = ['Id', 'InternalName', 'Title', 'TypeAsString', 'Group', 'Description', 'Required', 'Hidden', 'SchemaXml'];

/** SP.XmlSchemaFieldCreationInformation Options: AddFieldInternalNameHint – use the Name attribute as-is. */
const ADD_FIELD_INTERNAL_NAME_HINT = 8;

/** Changes the provider can apply to an existing field in 'update' mode. */
const UPDATABLE = ['title', 'group', 'description', 'required', 'choices', 'defaultValue', 'customFormatter'];

interface ISiteFieldDiff extends IDiffResult {
  target?: IFieldInfoLike;
}

const odataString = (s: string): string => s.replace(/'/g, "''");

const refOf = (def: IField): IDiffResult['ref'] => ({ kind: 'siteField', key: `field:${def.internalName}` });

function isTaxonomy(def: IField): boolean {
  return def.type === 'TaxonomyFieldType' || def.type === 'TaxonomyFieldTypeMulti';
}

function union(a: string[] | undefined, b: string[] | undefined): string[] {
  const out = (a || []).slice();
  (b || []).forEach((c) => {
    if (out.indexOf(c) < 0) out.push(c);
  });
  return out;
}

export class SiteFieldProvider implements IProvider<IField> {
  public readonly kind = 'siteField' as const;

  public async diff(sp: SPFI, def: IField, ctx: IInstallContext): Promise<ISiteFieldDiff> {
    throwIfAborted(ctx.signal);
    const ref = refOf(def);
    // availablefields: a column defined on a parent web also blocks creating the same name here.
    const byName = await sp.web.availablefields
      .filter(`InternalName eq '${odataString(def.internalName)}'`)
      .select(...SELECT)<IFieldInfoLike[]>();
    throwIfAborted(ctx.signal);

    if (byName.length === 0) {
      if (def.id && (await this._idTaken(sp, def.id))) {
        return { ref, status: 'unsupported', changes: ['idConflict'] };
      }
      // Taxonomy columns need term store mapping (phase 2); creating them unbound would leave them unusable.
      return { ref, status: isTaxonomy(def) ? 'unsupported' : 'new', changes: isTaxonomy(def) ? ['taxonomy'] : undefined };
    }

    const target = byName[0];
    const expected = toFieldDef(target, target.SchemaXml);
    const changes = compareFields(this._resolved(def, ctx), expected);
    return { ref, status: changes.length ? 'different' : 'same', changes: changes.length ? changes : undefined, target };
  }

  public async apply(sp: SPFI, def: IField, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;

    switch (diff.status) {
      case 'new':
        return this._create(sp, def, ctx);
      case 'same':
        return this._done(ref, 'skipped', def, diff.target!.Id, ctx);
      case 'different':
        if (mode === 'update') {
          return this._update(sp, def, diff, ctx);
        }
        if (mode === 'rename') {
          // Renaming would give the column a new internal name and ID, breaking content type and view references.
          ctx.log.warn('Site columns cannot be installed renamed; the existing column is kept.', { artifact: ref, code: 'FIELD_RENAME_UNSUPPORTED' });
        }
        return this._done(ref, 'skipped', def, diff.target!.Id, ctx);
      default:
        ctx.log.warn(`Site column skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'FIELD_UNSUPPORTED', detail: diff.changes });
        return { ref, outcome: 'skipped' };
    }
  }

  private async _create(sp: SPFI, def: IField, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = refOf(def);
    // Resolve, then sanitize again: a template may have been edited by hand.
    const sanitized = sanitizeFieldXml(resolve(def.schemaXml, ctx.tokens));
    if (sanitized.removedAttributes.length || sanitized.removedElements.length) {
      ctx.log.warn('Removed disallowed parts of SchemaXml before install.', {
        artifact: ref,
        code: 'FIELD_XML_SANITIZED',
        detail: { attributes: sanitized.removedAttributes, elements: sanitized.removedElements }
      });
    }
    const created = await sp.web.fields.createFieldAsXml({ SchemaXml: sanitized.xml, Options: ADD_FIELD_INTERNAL_NAME_HINT });
    if (!created.Id) {
      throw new CopyJetError('FIELD_CREATE_FAILED', `SharePoint returned no Id for ${def.internalName}.`);
    }
    if (created.InternalName && created.InternalName !== def.internalName) {
      ctx.log.warn(`Column was created as ${created.InternalName}.`, { artifact: ref, code: 'FIELD_NAME_CHANGED' });
    }
    ctx.log.info('Site column created.', { artifact: ref });
    return this._done(ref, 'created', def, created.Id, ctx);
  }

  private async _update(sp: SPFI, def: IField, diff: ISiteFieldDiff, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = diff.ref;
    const target = diff.target!;
    const changes = diff.changes || [];
    const skipped = changes.filter((c) => UPDATABLE.indexOf(c) < 0);
    if (skipped.length) {
      ctx.log.warn(`Not updatable on an existing column: ${skipped.join(', ')}.`, { artifact: ref, code: 'FIELD_PARTIAL_UPDATE', detail: skipped });
    }

    const d = this._resolved(def, ctx);
    const props: { [k: string]: unknown } = {};
    if (changes.indexOf('title') >= 0) props.Title = d.title;
    if (changes.indexOf('group') >= 0) props.Group = d.group || '';
    if (changes.indexOf('description') >= 0) props.Description = d.description || '';
    if (changes.indexOf('required') >= 0) props.Required = !!d.required;
    if (changes.indexOf('defaultValue') >= 0) props.DefaultValue = d.defaultValue === undefined ? null : d.defaultValue;
    if (changes.indexOf('customFormatter') >= 0) props.CustomFormatter = d.customFormatter || '';
    if (changes.indexOf('choices') >= 0) {
      // Never remove choices from the target: existing items may use them.
      props.Choices = union(toFieldDef(target, target.SchemaXml).choices, d.choices);
    }

    if (Object.keys(props).length === 0) {
      return this._done(ref, 'skipped', def, target.Id, ctx);
    }
    await sp.web.availablefields.getById(target.Id).update(props, `SP.Field${d.type === 'Choice' || d.type === 'MultiChoice' ? d.type : ''}`);
    ctx.log.info(`Site column updated: ${Object.keys(props).join(', ')}.`, { artifact: ref });
    return this._done(ref, 'updated', def, target.Id, ctx);
  }

  /** The definition with tokens in comparable properties resolved against the target. */
  private _resolved(def: IField, ctx: IInstallContext): IField {
    const r = (v: string | undefined): string | undefined => (v === undefined ? v : resolve(v, ctx.tokens));
    return { ...def, title: r(def.title)!, defaultValue: def.defaultValue === null ? null : r(def.defaultValue), formula: r(def.formula) };
  }

  private async _idTaken(sp: SPFI, id: string): Promise<boolean> {
    const found = await sp.web.availablefields.filter(`Id eq guid'${id}'`).select('Id')<Array<{ Id: string }>>();
    return found.length > 0;
  }

  private _done(ref: IApplyResult['ref'], outcome: IApplyResult['outcome'], def: IField, id: string, ctx: IInstallContext): IApplyResult {
    const fieldId = id.replace(/^\{|\}$/g, '').toLowerCase();
    ctx.tokens.set('fieldid', def.internalName, fieldId);
    if (outcome === 'skipped') {
      ctx.log.info('Site column already present; skipped.', { artifact: ref });
    }
    return { ref, outcome, tokens: { fieldid: { [def.internalName]: fieldId } } };
  }
}
