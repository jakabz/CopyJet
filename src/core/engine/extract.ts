import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import { throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import { loadSourceSite } from '../lists/sourceTokens';
import type { Logger } from '../logger/Logger';
import type { ArtifactKind, IArtifactRef, ICopyJetTemplate, IDiscoveredArtifact, IExtractOptions, ITemplateWriter } from '../model';
import { JsonTemplateWriter, ZipTemplateWriter, createEmptyTemplate } from '../packager';
import { templateNodes } from '../planner/planner';
import { createExtractors } from './registry';

export interface IDiscovery {
  artifacts: IDiscoveredArtifact[];
  /** Extractors that failed; the others' artifacts are still returned. */
  errors: Array<{ kind: ArtifactKind; error: unknown }>;
}

/** Runs every extractor's discover() (the Setup tree); one failing kind does not hide the others. */
export async function discoverSite(sp: SPFI, signal?: AbortSignal): Promise<IDiscovery> {
  const extractors = createExtractors();
  const results = await limitConcurrency(
    extractors.map((e) => () => e.discover(sp, signal)),
    3,
    signal
  );
  const out: IDiscovery = { artifacts: [], errors: [] };
  results.forEach((r, i) => {
    if (r.ok) {
      out.artifacts = out.artifacts.concat(r.value);
    } else {
      out.errors.push({ kind: extractors[i].kind, error: r.error });
    }
  });
  return out;
}

export interface IExtractRequest {
  refs: IArtifactRef[];
  /** The discovery the selection was made from; used to offer missing dependencies. */
  discovered: IDiscoveredArtifact[];
  name: string;
  description?: string;
  /** Login or e-mail of the user making the template (meta.createdBy). */
  createdBy: string;
  /** Items keep their author, editor and dates (default true). */
  preserveAuthors?: boolean;
  log: Logger;
  signal?: AbortSignal;
  /** Called before each extractor runs (with its kind) and once at the end (without). */
  onProgress?: (done: number, total: number, kind?: ArtifactKind) => void;
}

export interface IExtractResult {
  writer: ITemplateWriter;
  /** 'zip' when the template carries content (items, files), else 'json'. */
  format: 'json' | 'zip';
  /** Discovered, copyable artifacts the template depends on but that were not selected. */
  missing: IDiscoveredArtifact[];
}

/** Artifacts the template needs that the Setup could still add (discovered, supported, not selected). */
export function missingDependencies(template: ICopyJetTemplate, discovered: IDiscoveredArtifact[]): IDiscoveredArtifact[] {
  const nodes = templateNodes(template);
  const inTemplate: { [key: string]: boolean } = {};
  nodes.forEach((n) => (inTemplate[n.ref.key] = true));
  const needed: { [key: string]: boolean } = {};
  nodes.forEach((n) => n.deps.forEach((d) => (needed[d] = !inTemplate[d])));
  return discovered.filter((a) => needed[a.ref.key] && !a.unsupported);
}

/**
 * Extracts the selected artifacts into a template (.zip when it carries content, else .json). Extractors run
 * in registry order (lists before their columns, views and items). Returns the writer – finalize() validates and produces the file – and the
 * dependencies still missing from the selection.
 */
export async function extractTemplate(sp: SPFI, req: IExtractRequest): Promise<IExtractResult> {
  throwIfAborted(req.signal);
  const [web, site] = await Promise.all([
    sp.web.select('Url', 'Language')<{ Url: string; Language: number }>(),
    loadSourceSite(sp, req.signal)
  ]);
  const format = req.refs.some((r) => r.kind === 'items' || r.kind === 'files') ? 'zip' : 'json';
  const manifest = createEmptyTemplate({
    name: req.name,
    description: req.description || undefined,
    createdBy: req.createdBy,
    createdAt: new Date().toISOString(),
    sourceSiteUrl: web.Url,
    sourceTenant: web.Url.replace(/^https?:\/\//i, '').split('/')[0],
    sourceLcid: web.Language,
    includesContent: false
  });
  const writer: ITemplateWriter = format === 'zip' ? new ZipTemplateWriter(manifest) : new JsonTemplateWriter(manifest);
  if (!writer.manifest.meta.description) {
    delete writer.manifest.meta.description;
  }
  const extractors = createExtractors();
  for (let i = 0; i < extractors.length; i++) {
    throwIfAborted(req.signal);
    const e = extractors[i];
    if (req.onProgress) req.onProgress(i, extractors.length, e.kind);
    const refs = req.refs.filter((r) => r.kind === e.kind);
    if (refs.length === 0) continue;
    const opts: IExtractOptions = {
      includeContent: format === 'zip',
      includeVersions: false,
      includeMembers: false,
      preserveAuthors: req.preserveAuthors !== false,
      tokens: site.tokens,
      log: req.log,
      signal: req.signal
    };
    await e.extract(sp, refs, opts, writer);
    req.log.info(`Extracted: ${refs.length} × ${e.kind}.`, { step: e.kind });
  }
  if (req.onProgress) req.onProgress(extractors.length, extractors.length);
  return { writer, format, missing: missingDependencies(writer.manifest, req.discovered) };
}
