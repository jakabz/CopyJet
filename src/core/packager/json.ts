import { CopyJetError, throwIfAborted } from '../errors';
import type { ICopyJetTemplate, IMeta, ITemplateReader, ITemplateWriter } from '../model';
import { CURRENT_SCHEMA_VERSION, migrate, validateTemplate } from '../schema';
import { COPYJET_VERSION } from '../version';

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

/**
 * Opens a template file: parses, migrates to the current schema version and validates it.
 * Errors (CopyJetError.code): TEMPLATE_ZIP_UNSUPPORTED, TEMPLATE_PARSE, SCHEMA_*, TEMPLATE_INVALID.
 */
export async function openTemplate(file: Blob, signal?: AbortSignal): Promise<ITemplateReader> {
  throwIfAborted(signal);
  const head = new Uint8Array(await file.slice(0, 2).arrayBuffer());
  if (isZip(head)) {
    throw new CopyJetError('TEMPLATE_ZIP_UNSUPPORTED', '.zip templates are not supported yet.');
  }
  const text = (await file.text()).replace(/^﻿/, '');
  throwIfAborted(signal);
  let json: unknown;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new CopyJetError('TEMPLATE_PARSE', 'The file is not valid JSON.', e instanceof Error ? e.message : e);
  }
  if (json === null || typeof json !== 'object' || Array.isArray(json)) {
    throw new CopyJetError('TEMPLATE_PARSE', 'The file does not contain a JSON object.');
  }
  const migrated = migrate(json as Record<string, unknown>);
  const result = validateTemplate(migrated);
  if (!result.valid) {
    throw new CopyJetError('TEMPLATE_INVALID', 'The template does not match the schema.', result.errors);
  }
  return new JsonTemplateReader(migrated);
}
