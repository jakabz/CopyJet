import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/fields';
import { throwIfAborted } from '../errors';
import { isCustomField, templateFieldFrom, type IFieldInfoLike } from '../fields';
import { siteFieldDependencies } from '../planner/dependencies';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, IField, ITemplateWriter } from '../model';

/**
 * Visible site columns; the custom ones are picked client-side by SourceID (isCustomField), because the
 * REST API has no server-side property that separates them from built-in columns (spike 01).
 * Hidden columns (e.g. taxonomy note companions) are managed by SharePoint.
 */
const VISIBLE_FIELDS_FILTER = 'Hidden eq false';
const SELECT = ['Id', 'InternalName', 'Title', 'TypeAsString', 'Group', 'Description', 'Required', 'Hidden', 'SchemaXml'];

export const siteFieldKey = (internalName: string): string => `field:${internalName}`;

function listKeyOf(lookupList: string | undefined): string | undefined {
  const m = lookupList ? /^\{listkey:([^{}]+)\}$/.exec(lookupList) : null;
  return m ? m[1] : undefined;
}

export class SiteFieldExtractor implements IExtractor<IField> {
  public readonly kind = 'siteField' as const;

  public async discover(sp: SPFI, signal?: AbortSignal): Promise<IDiscoveredArtifact[]> {
    throwIfAborted(signal);
    const fields = await this._customFields(sp, ['InternalName', 'Title', 'Group', 'SchemaXml']);
    return fields.map((f) => ({ ref: { kind: this.kind, key: siteFieldKey(f.InternalName) }, title: f.Title, group: f.Group }));
  }

  public dependencies(def: IField): IArtifactRef[] {
    return siteFieldDependencies(def);
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) {
      return;
    }
    const fields = await this._customFields(sp, SELECT);
    throwIfAborted(opts.signal);

    fields
      .filter((f) => wanted.indexOf(siteFieldKey(f.InternalName)) >= 0)
      .forEach((f) => {
        const ref: IArtifactRef = { kind: this.kind, key: siteFieldKey(f.InternalName) };
        const def = templateFieldFrom(f, opts.tokens, opts.log, ref);
        this._warnings(def, ref, opts, out);

        const list = out.manifest.siteFields;
        const existing = list.findIndex((x) => x.internalName === def.internalName);
        if (existing >= 0) {
          list[existing] = def;
        } else {
          list.push(def);
        }
      });

    wanted
      .filter((key) => !fields.some((f) => siteFieldKey(f.InternalName) === key))
      .forEach((key) => opts.log.warn('Site column not found on the source site.', { artifact: { kind: this.kind, key }, code: 'FIELD_NOT_FOUND' }));
  }

  private async _customFields(sp: SPFI, select: string[]): Promise<IFieldInfoLike[]> {
    const fields = await sp.web.fields.filter(VISIBLE_FIELDS_FILTER).select(...select)<IFieldInfoLike[]>();
    return fields.filter((f) => isCustomField(f.SchemaXml));
  }

  private _warnings(def: IField, ref: IArtifactRef, opts: IExtractOptions, out: ITemplateWriter): void {
    const warn = (code: string, message: string): void => {
      opts.log.warn(message, { artifact: ref, code });
      const warnings = out.manifest.meta.warnings || [];
      // Re-extracting the same column must not repeat its warnings.
      if (!warnings.some((w) => w.code === code && w.artifact === ref.key)) {
        out.manifest.meta.warnings = warnings.concat({ code, message, artifact: ref.key });
      }
    };
    if ((def.type === 'Lookup' || def.type === 'LookupMulti') && !listKeyOf(def.lookupList)) {
      warn('LOOKUP_TARGET_UNKNOWN', `Lookup column ${def.internalName} points to a list outside the template.`);
    }
    if (def.type === 'Other') {
      warn('FIELD_TYPE_UNSUPPORTED', `Column ${def.internalName} has an unsupported type; it is copied as-is.`);
    }
  }
}
