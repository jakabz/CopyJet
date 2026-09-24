import { CopyJetError } from '../errors';
import type { ICopyJetTemplate } from '../model';

export const CURRENT_SCHEMA_VERSION = '1.0';

/**
 * Migration chain keyed by the version it upgrades FROM; each step returns the next minor version.
 * Example for a future 1.1: '1.0': (t) => ({ ...t, schemaVersion: '1.1', newField: [] })
 */
const MIGRATIONS: Record<string, (t: Record<string, unknown>) => Record<string, unknown>> = {};

function parseVersion(v: unknown): { major: number; minor: number } {
  const m = typeof v === 'string' ? /^(\d+)\.(\d+)$/.exec(v) : null;
  if (!m) {
    throw new CopyJetError('SCHEMA_VERSION_INVALID', `Invalid schemaVersion: ${String(v)}`);
  }
  return { major: parseInt(m[1], 10), minor: parseInt(m[2], 10) };
}

/**
 * Upgrades an older 1.x template to CURRENT_SCHEMA_VERSION. Rejects another major version and
 * newer minors (a template made by a newer CopyJet). Does not mutate the input.
 */
export function migrate(json: Record<string, unknown>): ICopyJetTemplate {
  const current = parseVersion(CURRENT_SCHEMA_VERSION);
  const from = parseVersion(json.schemaVersion);
  if (from.major !== current.major) {
    throw new CopyJetError('SCHEMA_MAJOR_MISMATCH', `Unsupported major schema version ${from.major}.`);
  }
  if (from.minor > current.minor) {
    throw new CopyJetError('SCHEMA_TOO_NEW', `Template schema ${String(json.schemaVersion)} is newer than ${CURRENT_SCHEMA_VERSION}.`);
  }
  let t = json;
  while (t.schemaVersion !== CURRENT_SCHEMA_VERSION) {
    const step = MIGRATIONS[t.schemaVersion as string];
    if (!step) {
      throw new CopyJetError('SCHEMA_NO_MIGRATION', `No migration from schema ${String(t.schemaVersion)}.`);
    }
    t = step(t);
  }
  return t as unknown as ICopyJetTemplate;
}
