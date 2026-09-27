import { CopyJetError, throwIfAborted } from '../errors';
import type { ICopyJetTemplate, IMeta, ITemplateReader, ITemplateWriter } from '../model';
import { CURRENT_SCHEMA_VERSION, migrate, validateTemplate } from '../schema';
import { COPYJET_VERSION } from '../version';
import { loadZipPackage } from './zip';

export type NewTemplateMeta = Omit<IMeta, 'generator' | 'checksum' | 'estimatedSizeBytes'>;

/** A manifest with all required collections present and empty. */
export function createEmptyTemplate(meta: NewTemplateMeta): ICopyJetTemplate {
  return {
    schemaVersion: CURRENT_SCHEMA_VERSION,
    meta: { ...meta, generator: { name: 'CopyJet', version: COPYJET_VERSION } },
    siteFields: [],
    contentTypes: [],
    lists: [],
    groups: [],
    pages: [],
    principals: []
  };
}

/** Structure-only template: a single manifest.json. Content (items, files, pages) needs the zip format. */
export class JsonTemplateWriter implements ITemplateWriter {
  public readonly manifest: ICopyJetTemplate;

  constructor(manifest: ICopyJetTemplate) {
    this.manifest = manifest;
  }

  public addJson(path: string, _value: unknown): void {
    throw new CopyJetError('PACKAGE_JSON_NO_ENTRIES', `A .json template cannot hold package entries (${path}); use .zip.`);
  }

  public addBlob(path: string, _blob: Blob): void {
    throw new CopyJetError('PACKAGE_JSON_NO_ENTRIES', `A .json template cannot hold package entries (${path}); use .zip.`);
  }

  public async finalize(signal?: AbortSignal): Promise<Blob> {
    throwIfAborted(signal);
    const result = validateTemplate(this.manifest);
    if (!result.valid) {
      throw new CopyJetError('TEMPLATE_INVALID', 'The template does not match the schema.', result.errors);
    }
    return new Blob([JSON.stringify(this.manifest, null, 2)], { type: 'application/json' });
  }
}

class JsonTemplateReader implements ITemplateReader {
  public readonly manifest: ICopyJetTemplate;

  constructor(manifest: ICopyJetTemplate) {
    this.manifest = manifest;
  }

  public has(): boolean {
    return false;
  }

  public getJson<T>(path: string): Promise<T> {
    return Promise.reject(new CopyJetError('PACKAGE_ENTRY_NOT_FOUND', `Entry not found: ${path}`, { path }));
  }

  public getBlob(path: string): Promise<Blob> {
    return Promise.reject(new CopyJetError('PACKAGE_ENTRY_NOT_FOUND', `Entry not found: ${path}`, { path }));
  }
}

function isZip(head: Uint8Array): boolean {
  return head.length >= 2 && head[0] === 0x50 && head[1] === 0x4b; // "PK"
}

/** Migrates and validates a parsed manifest. */
function checkedManifest(json: unknown): ICopyJetTemplate {
  if (json === null || typeof json !== 'object' || Array.isArray(json)) {
    throw new CopyJetError('TEMPLATE_PARSE', 'The file does not contain a JSON object.');
  }
  const migrated = migrate(json as Record<string, unknown>);
  const result = validateTemplate(migrated);
  if (!result.valid) {
    throw new CopyJetError('TEMPLATE_INVALID', 'The template does not match the schema.', result.errors);
  }
  return migrated;
}

/** Content entries the manifest points to must be in the package. */
function checkContentEntries(reader: ITemplateReader): ITemplateReader {
  const missing = reader.manifest.lists
    .filter((l) => l.content && l.content.mode !== 'none' && l.content.sourceMode !== 'reference')
    .map((l) => l.content.source || '')
    .filter((source) => !source || (!reader.has(source) && !reader.has(`${source.replace(/\/+$/, '')}/_meta.json`)));
  if (missing.length) {
    throw new CopyJetError('TEMPLATE_CONTENT_MISSING', 'The template refers to content that is not in the file.', missing);
  }
  return reader;
}

/**
 * Opens a template file (.json or .zip): parses, migrates to the current schema version and validates it.
 * Errors (CopyJetError.code): TEMPLATE_PARSE, SCHEMA_*, TEMPLATE_INVALID, TEMPLATE_CONTENT_MISSING.
 */
export async function openTemplate(file: Blob, signal?: AbortSignal): Promise<ITemplateReader> {
  throwIfAborted(signal);
  const head = new Uint8Array(await file.slice(0, 2).arrayBuffer());
  if (isZip(head)) {
    const pkg = await loadZipPackage(file, signal);
    return checkContentEntries(pkg.open(checkedManifest(pkg.manifest)));
  }
  const text = (await file.text()).replace(/^\uFEFF/, '');
  throwIfAborted(signal);
  let json: unknown;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new CopyJetError('TEMPLATE_PARSE', 'The file is not valid JSON.', e instanceof Error ? e.message : e);
  }
  return checkContentEntries(new JsonTemplateReader(checkedManifest(json)));
}
