import type { SPFI } from '@pnp/sp';
import { throwIfAborted } from '../errors';
import { isSupportedTemplate, loadSourceSite, toListDef } from '../lists';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, IList, ITemplateWriter } from '../model';

export const listRefKey = (key: string): string => `list:${key}`;

/**
 * Lists and libraries: settings only for now. Content type bindings, folders, list columns and views are
 * added by later steps of phase 1 (their own extractor/provider pairs).
 */
export class ListExtractor implements IExtractor<IList> {
  public readonly kind = 'list' as const;

  public async discover(sp: SPFI, signal?: AbortSignal): Promise<IDiscoveredArtifact[]> {
    const site = await loadSourceSite(sp, signal);
    return site.lists.map((l) => {
      const found: IDiscoveredArtifact = { ref: { kind: this.kind, key: listRefKey(l.key) }, title: l.info.Title, itemCount: l.info.ItemCount };
      if (!isSupportedTemplate(l.info.BaseTemplate)) found.unsupported = 'LIST_TEMPLATE_UNSUPPORTED';
      return found;
    });
  }

  /** Content types the list uses (by ID); built-in ones are ignored by the Setup as they are not discovered. */
  public dependencies(def: IList): IArtifactRef[] {
    return (def.contentTypes || []).filter((ct) => /^0x/i.test(ct)).map((ct) => ({ kind: 'contentType' as const, key: `contentType:${ct}` }));
  }

  public async extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void> {
    throwIfAborted(opts.signal);
    const wanted = refs.filter((r) => r.kind === this.kind).map((r) => r.key);
    if (wanted.length === 0) {
      return;
    }
    const site = await loadSourceSite(sp, opts.signal);

    site.lists
      .filter((l) => wanted.indexOf(listRefKey(l.key)) >= 0)
      .forEach((l) => {
        const ref: IArtifactRef = { kind: this.kind, key: listRefKey(l.key) };
        if (!isSupportedTemplate(l.info.BaseTemplate)) {
          const message = `List ${l.info.Title} uses template ${l.info.BaseTemplate}, which CopyJet does not copy yet.`;
          opts.log.warn(message, { artifact: ref, code: 'LIST_TEMPLATE_UNSUPPORTED' });
          const warnings = out.manifest.meta.warnings || [];
          if (!warnings.some((w) => w.code === 'LIST_TEMPLATE_UNSUPPORTED' && w.artifact === ref.key)) {
            out.manifest.meta.warnings = warnings.concat({ code: 'LIST_TEMPLATE_UNSUPPORTED', message, artifact: ref.key });
          }
          return;
        }
        const def = toListDef(l.info, l.key, site.webServerRelativeUrl);
        const i = out.manifest.lists.findIndex((x) => x.key === def.key);
        if (i >= 0) {
          out.manifest.lists[i] = def;
        } else {
          out.manifest.lists.push(def);
        }
      });

    wanted
      .filter((key) => !site.lists.some((l) => listRefKey(l.key) === key))
      .forEach((key) => opts.log.warn('List not found on the source site.', { artifact: { kind: this.kind, key }, code: 'LIST_NOT_FOUND' }));
  }
}
