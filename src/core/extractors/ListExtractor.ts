import type { SPFI } from '@pnp/sp';
import { throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import { isSupportedTemplate, listFolderPaths, loadSourceSite, readListContentTypes, siteContentTypeIdOf, toListDef, type IListInfoLike } from '../lists';
import type { IArtifactRef, IDiscoveredArtifact, IExtractOptions, IExtractor, IList, ITemplateWriter } from '../model';

export const listRefKey = (key: string): string => `list:${key}`;

/**
 * Lists and libraries: settings, content type bindings (site content type IDs, default first) and folders.
 * List columns and views have their own extractor/provider pairs.
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

    const selected = site.lists.filter((l) => wanted.indexOf(listRefKey(l.key)) >= 0);
    const supported = selected.filter((l) => {
      if (isSupportedTemplate(l.info.BaseTemplate)) return true;
      const ref: IArtifactRef = { kind: this.kind, key: listRefKey(l.key) };
      const message = `List ${l.info.Title} uses template ${l.info.BaseTemplate}, which CopyJet does not copy yet.`;
      opts.log.warn(message, { artifact: ref, code: 'LIST_TEMPLATE_UNSUPPORTED' });
      const warnings = out.manifest.meta.warnings || [];
      if (!warnings.some((w) => w.code === 'LIST_TEMPLATE_UNSUPPORTED' && w.artifact === ref.key)) {
        out.manifest.meta.warnings = warnings.concat({ code: 'LIST_TEMPLATE_UNSUPPORTED', message, artifact: ref.key });
      }
      return false;
    });

    const results = await limitConcurrency(
      supported.map((l) => () => this._definition(sp, l.info, l.key, site.webServerRelativeUrl, opts.signal)),
      4,
      opts.signal
    );
    results.forEach((r, i) => {
      const ref: IArtifactRef = { kind: this.kind, key: listRefKey(supported[i].key) };
      if (!r.ok) {
        opts.log.error('Could not read the list structure.', { artifact: ref, code: 'LIST_READ_FAILED', detail: r.error });
        return;
      }
      const at = out.manifest.lists.findIndex((x) => x.key === r.value.key);
      if (at >= 0) {
        out.manifest.lists[at] = r.value;
      } else {
        out.manifest.lists.push(r.value);
      }
    });

    wanted
      .filter((key) => !site.lists.some((l) => listRefKey(l.key) === key))
      .forEach((key) => opts.log.warn('List not found on the source site.', { artifact: { kind: this.kind, key }, code: 'LIST_NOT_FOUND' }));
  }

  private async _definition(sp: SPFI, info: IListInfoLike, key: string, webServerRelativeUrl: string, signal?: AbortSignal): Promise<IList> {
    const def = toListDef(info, key, webServerRelativeUrl);
    const listUrl = info.RootFolder.ServerRelativeUrl;
    const [cts, folders] = await Promise.all([
      info.ContentTypesEnabled ? readListContentTypes(sp, listUrl) : Promise.resolve(undefined),
      listFolderPaths(sp, listUrl, signal)
    ]);
    if (cts && cts.ordered.length) {
      // Site-level IDs, default first; the list-level IDs are regenerated on the target.
      def.contentTypes = cts.ordered.map(siteContentTypeIdOf);
    }
    if (folders.length) {
      def.folders = folders.map((path) => ({ path }));
    }
    return def;
  }
}
