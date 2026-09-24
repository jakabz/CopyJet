import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/content-types';
import { ContentTypes } from '@pnp/sp/content-types';
import { contentTypeKey, isCustomContentType, parentIdOf, toContentTypeDef, type IContentTypeInfoLike, type IFieldLinkInfoLike } from '../contentTypes';
import { throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import type { IArtifactRef, IContentType, IDiscoveredArtifact, IExtractOptions, IExtractor, ITemplateWriter } from '../model';

const SELECT = ['StringId', 'Name', 'Group', 'Description', 'Hidden', 'ReadOnly', 'Sealed', 'SchemaXml'];
const LINK_SELECT = ['FieldInternalName', 'Name', 'Required', 'Hidden'];

export class ContentTypeExtractor implements IExtractor<IContentType> {
  public readonly kind = 'contentType' as const;

  public async discover(sp: SPFI, signal?: AbortSignal): Promise<IDiscoveredArtifact[]> {
    throwIfAborted(signal);
    const cts = await this._customTypes(sp);
    return cts.map((ct) => ({ ref: { kind: this.kind, key: contentTypeKey(ct.StringId) }, title: ct.Name, group: ct.Group }));
  }

  /**
   * The parent (unless it is a built-in type, which exists everywhere – the Setup keeps only dependencies it
   * discovered) and the site columns the type links. Built-in columns are filtered the same way.
   */
  public dependencies(def: IContentType): IArtifactRef[] {
    const refs: IArtifactRef[] = def.fieldRefs.map((r) => ({ kind: 'siteField' as const, key: `field:${r.internalName}` }));
    if (def.parentId && def.parentId !== '0x') {
      refs.unshift({ kind: 'contentType', key: contentTypeKey(def.parentId) });
    }
    return refs;
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) {
      return;
    }
    const cts = (await this._customTypes(sp)).filter((ct) => wanted.indexOf(contentTypeKey(ct.StringId)) >= 0);
    // Parents come first so a child never precedes its parent in the template.
    cts.sort((a, b) => a.StringId.length - b.StringId.length);

    const available = ContentTypes(sp.web, 'availablecontenttypes');
    const linksOf = (id: string): Promise<IFieldLinkInfoLike[]> =>
      available.getById(id).fieldLinks.select(...LINK_SELECT)<IFieldLinkInfoLike[]>();

    const results = await limitConcurrency(
      cts.map((ct) => async () => {
        const parentId = parentIdOf(ct.StringId);
        const [links, parentLinks] = await Promise.all([linksOf(ct.StringId), parentId && parentId !== '0x' ? linksOf(parentId) : Promise.resolve([])]);
        return toContentTypeDef(ct, links, parentLinks);
      }),
      4,
      opts.signal
    );

    results.forEach((r, i) => {
      const ref: IArtifactRef = { kind: this.kind, key: contentTypeKey(cts[i].StringId) };
      if (!r.ok) {
        opts.log.error('Could not read content type field links.', { artifact: ref, code: 'CT_READ_FAILED', detail: r.error });
        return;
      }
      const list = out.manifest.contentTypes;
      const existing = list.findIndex((x) => x.id === r.value.id);
      if (existing >= 0) {
        list[existing] = r.value;
      } else {
        list.push(r.value);
      }
    });

    wanted
      .filter((key) => !cts.some((ct) => contentTypeKey(ct.StringId) === key))
      .forEach((key) => opts.log.warn('Content type not found on the source site.', { artifact: { kind: this.kind, key }, code: 'CT_NOT_FOUND' }));
  }

  private async _customTypes(sp: SPFI): Promise<IContentTypeInfoLike[]> {
    const cts = await sp.web.contentTypes.select(...SELECT)<IContentTypeInfoLike[]>();
    return cts.filter(isCustomContentType);
  }
}
