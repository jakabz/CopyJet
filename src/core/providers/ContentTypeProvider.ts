import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/content-types';
import { ContentTypes, type IContentTypes } from '@pnp/sp/content-types';
import {
  compareContentTypes,
  contentTypeKey,
  missingFieldRefs,
  normalizeContentTypeId,
  type IContentTypeInfoLike,
  type IFieldLinkInfoLike
} from '../contentTypes';
import { CopyJetError, throwIfAborted } from '../errors';
import { csomCreateContentType, csomUpdateFieldLinks, type FetchLike } from '../http/csom';
import type { IFieldLinkFlags } from '../http/csomXml';
import { isHttpStatus } from '../http/status';
import type { ConflictMode, IApplyResult, IContentType, IDiffResult, IInstallContext, IProvider } from '../model';

const SELECT = ['StringId', 'Name', 'Group', 'Description', 'Hidden', 'ReadOnly', 'Sealed', 'SchemaXml'];
const LINK_SELECT = ['Id', 'FieldInternalName', 'Name', 'Required', 'Hidden'];

/** Changes 'update' mode can apply. */
const UPDATABLE = ['name', 'group', 'description', 'fieldRefs', 'fieldRefFlags'];

const linkName = (l: IFieldLinkInfoLike): string => l.FieldInternalName || l.Name;

interface IContentTypeDiff extends IDiffResult {
  target?: IContentTypeInfoLike;
  targetLinks?: IFieldLinkInfoLike[];
}

const odataString = (s: string): string => s.replace(/'/g, "''");

// Content types visible on the web, including those inherited from parent webs.
const available = (sp: SPFI): IContentTypes => ContentTypes(sp.web, 'availablecontenttypes');

export class ContentTypeProvider implements IProvider<IContentType> {
  public readonly kind = 'contentType' as const;
  private readonly _fetch?: FetchLike;

  /** `fetchImpl` is used for CSOM calls (injected in tests). */
  constructor(fetchImpl?: FetchLike) {
    this._fetch = fetchImpl;
  }

  public async diff(sp: SPFI, def: IContentType, ctx: IInstallContext): Promise<IContentTypeDiff> {
    throwIfAborted(ctx.signal);
    const ref = { kind: this.kind, key: contentTypeKey(def.id) };
    const target = await this._get(sp, def.id);
    throwIfAborted(ctx.signal);

    if (!target) {
      // A same-named type with another ID blocks creation, and would not be the parent our children expect.
      const sameName = await available(sp).filter(`Name eq '${odataString(def.name)}'`).select('StringId')<Array<{ StringId: string }>>();
      return sameName.length ? { ref, status: 'unsupported', changes: ['nameConflict'] } : { ref, status: 'new' };
    }
    const targetLinks = await this._links(sp, target.StringId);
    const changes = compareContentTypes(def, target, targetLinks);
    return { ref, status: changes.length ? 'different' : 'same', changes: changes.length ? changes : undefined, target, targetLinks };
  }

  public async apply(sp: SPFI, def: IContentType, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult> {
    const diff = await this.diff(sp, def, ctx);
    throwIfAborted(ctx.signal);
    const ref = diff.ref;

    switch (diff.status) {
      case 'new':
        return this._create(sp, def, ctx);
      case 'same':
        ctx.log.info('Content type already present; skipped.', { artifact: ref });
        return { ref, outcome: 'skipped' };
      case 'different':
        if (mode === 'update') {
          return this._update(sp, def, diff, ctx);
        }
        if (mode === 'rename') {
          // A renamed copy needs a new ID, which would cut it off from the children and lists that expect this one.
          ctx.log.warn('Content types cannot be installed renamed; the existing one is kept.', { artifact: ref, code: 'CT_RENAME_UNSUPPORTED' });
        }
        ctx.log.info(`Content type differs (${(diff.changes || []).join(', ')}); kept as it is.`, { artifact: ref });
        return { ref, outcome: 'skipped' };
      default:
        ctx.log.warn(`Content type skipped: ${(diff.changes || []).join(', ')}.`, { artifact: ref, code: 'CT_UNSUPPORTED', detail: diff.changes });
        return { ref, outcome: 'skipped' };
    }
  }

  private async _create(sp: SPFI, def: IContentType, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = { kind: this.kind, key: contentTypeKey(def.id) };
    // CSOM, not REST: REST ignores the requested ID and even the parent it implies (spike 02).
    await csomCreateContentType(sp, { id: def.id, name: def.name, group: def.group, description: def.description }, ctx.signal, this._fetch);
    // The ID is what children and lists rely on; verify SharePoint kept it.
    const created = await this._get(sp, def.id);
    if (!created) {
      throw new CopyJetError('CT_ID_NOT_KEPT', `Content type ${def.name} was created without the requested ID ${def.id}.`, { id: def.id });
    }
    // The new type already has its parent's links; add only the missing ones and align overridden flags.
    await this._syncLinks(sp, def.id, def, await this._links(sp, def.id), ctx);
    ctx.log.info('Content type created.', { artifact: ref });
    return { ref, outcome: 'created' };
  }

  private async _update(sp: SPFI, def: IContentType, diff: IContentTypeDiff, ctx: IInstallContext): Promise<IApplyResult> {
    const ref = diff.ref;
    const changes = diff.changes || [];
    const notUpdatable = changes.filter((c) => UPDATABLE.indexOf(c) < 0);
    if (notUpdatable.length) {
      ctx.log.warn(`Not updatable on an existing content type: ${notUpdatable.join(', ')}.`, { artifact: ref, code: 'CT_PARTIAL_UPDATE', detail: notUpdatable });
    }
    const props: { [k: string]: string } = {};
    if (changes.indexOf('name') >= 0) props.Name = def.name;
    if (changes.indexOf('group') >= 0) props.Group = def.group || '';
    if (changes.indexOf('description') >= 0) props.Description = def.description || '';
    if (Object.keys(props).length) {
      await sp.web.contentTypes.getById(diff.target!.StringId).update(props);
    }
    // Only adds links: removing one from the target could hide data in lists using the type.
    const linkChanges = await this._syncLinks(sp, diff.target!.StringId, def, diff.targetLinks || [], ctx);
    if (!Object.keys(props).length && linkChanges === 0) {
      return { ref, outcome: 'skipped' };
    }
    ctx.log.info('Content type updated.', { artifact: ref, detail: { properties: Object.keys(props), fieldLinkChanges: linkChanges } });
    return { ref, outcome: 'updated' };
  }

  private _links(sp: SPFI, id: string): Promise<IFieldLinkInfoLike[]> {
    return available(sp).getById(id).fieldLinks.select(...LINK_SELECT)<IFieldLinkInfoLike[]>();
  }

  /**
   * Adds missing field links and aligns Required/Hidden on existing ones, in one CSOM request (REST can do
   * neither, spike 02). Never removes a link. Returns the number of links added or changed.
   */
  private async _syncLinks(sp: SPFI, id: string, def: IContentType, links: IFieldLinkInfoLike[], ctx: IInstallContext): Promise<number> {
    const add = missingFieldRefs(def, links).map((r) => ({ internalName: r.internalName, required: !!r.required, hidden: !!r.hidden }));
    const flags: IFieldLinkFlags[] = [];
    def.fieldRefs.forEach((r) => {
      const l = links.filter((x) => linkName(x) === r.internalName)[0];
      if (l && l.Id && (!!l.Required !== !!r.required || !!l.Hidden !== !!r.hidden)) {
        flags.push({ id: l.Id, required: !!r.required, hidden: !!r.hidden });
      }
    });
    try {
      await csomUpdateFieldLinks(sp, id, add, flags, ctx.signal, this._fetch);
    } catch (e) {
      const names = add.map((a) => a.internalName).join(', ');
      throw new CopyJetError('CT_FIELDLINK_FAILED', `Could not update column links (${names || 'flags'}) of content type ${id}; are the site columns installed?`, e);
    }
    return add.length + flags.length;
  }

  /**
   * The content type, or undefined when missing. SharePoint Online answers a missing ID with
   * 200 {"odata.null": true} rather than 404 (spike 02); both are handled.
   */
  private async _get(sp: SPFI, id: string): Promise<IContentTypeInfoLike | undefined> {
    try {
      const ct = await available(sp).getById(normalizeContentTypeId(id)).select(...SELECT)<IContentTypeInfoLike>();
      return ct && ct.StringId ? ct : undefined;
    } catch (e) {
      if (isHttpStatus(e, 404)) {
        return undefined;
      }
      throw e;
    }
  }
}
