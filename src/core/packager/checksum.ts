import { throwIfAborted } from '../errors';
import type { ICopyJetTemplate, ITemplateReader } from '../model';

/**
 * Package checksum (meta.checksum, "sha256:<hex>"; rendszerterv §9). Defined so that it can be computed one
 * entry at a time – a 200 MB package never has to be hashed as one buffer (Web Crypto has no streaming):
 *
 *   entry hash = SHA-256 of the entry's bytes; for manifest.json the manifest without meta.checksum,
 *                serialized with JSON.stringify (no indentation);
 *   checksum   = SHA-256 of the lines "<path>\n<entry hash>\n", sorted by path.
 */

export type ChecksumStatus = 'ok' | 'mismatch' | 'missing';

const MANIFEST = 'manifest.json';

function hex(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let out = '';
  for (let i = 0; i < bytes.length; i++) out += (bytes[i] < 16 ? '0' : '') + bytes[i].toString(16);
  return out;
}

export async function sha256Hex(data: Uint8Array): Promise<string> {
  return hex(await crypto.subtle.digest('SHA-256', data));
}

const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);

/** The manifest's bytes as hashed: without meta.checksum, compact JSON. */
export function manifestChecksumBytes(manifest: ICopyJetTemplate): Uint8Array {
  const meta = { ...manifest.meta };
  delete meta.checksum;
  return utf8(JSON.stringify({ ...manifest, meta }));
}

export interface IChecksumEntry {
  path: string;
  /** Called once, when the entry is hashed. */
  bytes: () => Promise<Uint8Array>;
}

/** "sha256:<hex>" of the manifest and the given package entries (manifest.json itself is added here). */
export async function packageChecksum(manifest: ICopyJetTemplate, entries: IChecksumEntry[], signal?: AbortSignal): Promise<string> {
  const all: IChecksumEntry[] = [{ path: MANIFEST, bytes: async () => manifestChecksumBytes(manifest) }].concat(entries.filter((e) => e.path !== MANIFEST));
  all.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
  let lines = '';
  for (const entry of all) {
    throwIfAborted(signal);
    lines += `${entry.path}\n${await sha256Hex(await entry.bytes())}\n`;
  }
  return `sha256:${await sha256Hex(utf8(lines))}`;
}

/** Recomputes a loaded package's checksum and compares it with meta.checksum. */
export async function verifyChecksum(reader: ITemplateReader, signal?: AbortSignal): Promise<ChecksumStatus> {
  const expected = reader.storedManifest.meta.checksum;
  if (!expected) return 'missing';
  const entries = reader.entries().map((path) => ({ path, bytes: async () => new Uint8Array(await (await reader.getBlob(path, signal)).arrayBuffer()) }));
  return (await packageChecksum(reader.storedManifest, entries, signal)) === expected ? 'ok' : 'mismatch';
}
