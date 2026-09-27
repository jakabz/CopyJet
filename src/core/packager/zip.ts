import JSZip from 'jszip';
import { CopyJetError, throwIfAborted } from '../errors';
import type { ICopyJetTemplate, ITemplateReader, ITemplateWriter } from '../model';
import { validate, validateTemplate, type SchemaTarget } from '../schema';

export const MANIFEST_ENTRY = 'manifest.json';

/** Package paths are relative, forward-slashed and never leave the package. */
export function checkEntryPath(path: string): string {
  if (!path || /^[\\/]/.test(path) || /\\/.test(path) || path.split('/').some((p) => p === '..' || p === '.' || p === '') || path === MANIFEST_ENTRY) {
    throw new CopyJetError('PACKAGE_PATH_INVALID', `Invalid package path: ${path}`, { path });
  }
  return path;
}

/** Which schema $def a JSON entry must match (items/<key>.json, files/<key>/_meta.json, pages/<name>.json). */
export function schemaTargetOf(path: string): SchemaTarget | undefined {
  if (/^items\/[^/]+\.json$/.test(path)) return 'itemsFile';
  if (/^files\/[^/]+\/_meta\.json$/.test(path)) return 'filesMetaFile';
  if (/^pages\/[^/]+\.json$/.test(path)) return 'pageFile';
  return undefined;
}

function validated<T>(path: string, value: T): T {
  const target = schemaTargetOf(path);
  if (target) {
    const result = validate(target, value);
    if (!result.valid) {
      throw new CopyJetError('TEMPLATE_INVALID', `Package entry ${path} does not match the schema.`, result.errors);
    }
  }
  return value;
}

/** Template with content: manifest.json plus item/file entries in one .zip. */
export class ZipTemplateWriter implements ITemplateWriter {
  public readonly manifest: ICopyJetTemplate;
  private readonly _json: { [path: string]: unknown } = {};
  private readonly _blobs: { [path: string]: Blob } = {};

  constructor(manifest: ICopyJetTemplate) {
    this.manifest = manifest;
  }

  public addJson(path: string, value: unknown): void {
    this._json[checkEntryPath(path)] = value;
  }

  public addBlob(path: string, blob: Blob): void {
    this._blobs[checkEntryPath(path)] = blob;
  }

  public async finalize(signal?: AbortSignal): Promise<Blob> {
    throwIfAborted(signal);
    const result = validateTemplate(this.manifest);
    if (!result.valid) {
      throw new CopyJetError('TEMPLATE_INVALID', 'The template does not match the schema.', result.errors);
    }
    const zip = new JSZip();
    zip.file(MANIFEST_ENTRY, JSON.stringify(this.manifest, null, 2));
    Object.keys(this._json).forEach((path) => zip.file(path, JSON.stringify(validated(path, this._json[path]))));
    for (const path of Object.keys(this._blobs)) {
      zip.file(path, new Uint8Array(await this._blobs[path].arrayBuffer()));
    }
    throwIfAborted(signal);
    // uint8array rather than blob: JSZip's blob output needs browser detection that fails under Node (tests).
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    throwIfAborted(signal);
    return new Blob([bytes], { type: 'application/zip' });
  }
}

class ZipTemplateReader implements ITemplateReader {
  public readonly manifest: ICopyJetTemplate;
  private readonly _zip: JSZip;
  private readonly _cache: { [path: string]: unknown } = {};

  constructor(manifest: ICopyJetTemplate, zip: JSZip) {
    this.manifest = manifest;
    this._zip = zip;
  }

  public has(path: string): boolean {
    return !!this._zip.file(path);
  }

  public async getJson<T>(path: string, signal?: AbortSignal): Promise<T> {
    throwIfAborted(signal);
    if (Object.prototype.hasOwnProperty.call(this._cache, path)) return this._cache[path] as T;
    const text = await this._entry(path).async('string');
    let value: T;
    try {
      value = JSON.parse(text.replace(/^﻿/, '')) as T;
    } catch (e) {
      throw new CopyJetError('TEMPLATE_PARSE', `Package entry ${path} is not valid JSON.`, e instanceof Error ? e.message : e);
    }
    this._cache[path] = validated(path, value);
    return value;
  }

  public async getBlob(path: string, signal?: AbortSignal): Promise<Blob> {
    throwIfAborted(signal);
    return new Blob([await this._entry(path).async('uint8array')]);
  }

  private _entry(path: string): JSZip.JSZipObject {
    const entry = this._zip.file(checkEntryPath(path));
    if (!entry) {
      throw new CopyJetError('PACKAGE_ENTRY_NOT_FOUND', `Entry not found: ${path}`, { path });
    }
    return entry;
  }
}

/** Reads a .zip package: its manifest.json (returned for migration and validation) and the reader over it. */
export async function loadZipPackage(file: Blob, signal?: AbortSignal): Promise<{ manifest: unknown; open: (manifest: ICopyJetTemplate) => ITemplateReader }> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(new Uint8Array(await file.arrayBuffer()));
  } catch (e) {
    throw new CopyJetError('TEMPLATE_PARSE', 'The file is not a valid .zip package.', e instanceof Error ? e.message : e);
  }
  throwIfAborted(signal);
  const entry = zip.file(MANIFEST_ENTRY);
  if (!entry) {
    throw new CopyJetError('TEMPLATE_PARSE', 'The package has no manifest.json.');
  }
  let manifest: unknown;
  try {
    manifest = JSON.parse((await entry.async('string')).replace(/^﻿/, ''));
  } catch (e) {
    throw new CopyJetError('TEMPLATE_PARSE', 'manifest.json is not valid JSON.', e instanceof Error ? e.message : e);
  }
  return { manifest, open: (m) => new ZipTemplateReader(m, zip) };
}
