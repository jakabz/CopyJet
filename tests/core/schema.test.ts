import { readdirSync, readFileSync } from 'fs';
import { join } from 'path';
import { validate, validateTemplate, migrate, CURRENT_SCHEMA_VERSION } from '../../src/core/schema';
import { CopyJetError } from '../../src/core/errors';

const root = join(__dirname, '..', '..');
const readJson = (path: string): Record<string, unknown> => JSON.parse(readFileSync(path, 'utf8'));
const examples = join(root, 'schema', 'examples');
const invalid = join(root, 'tests', 'fixtures', 'invalid');

describe('schema examples', () => {
  const files = readdirSync(examples).filter((f) => f.endsWith('.json'));

  it('has examples to check', () => {
    expect(files.length).toBeGreaterThan(0);
  });

  it.each(files.filter((f) => f.startsWith('manifest.')))('%s is a valid manifest', (f) => {
    const r = validateTemplate(readJson(join(examples, f)));
    expect(r.errors).toEqual([]);
    expect(r.valid).toBe(true);
  });

  it.each(files.filter((f) => f.startsWith('items.')))('%s is a valid items file', (f) => {
    const r = validate('itemsFile', readJson(join(examples, f)));
    expect(r.errors).toEqual([]);
  });
});

describe('package part validators', () => {
  it('accepts a minimal page file and rejects one without canvas', () => {
    expect(validate('pageFile', { name: 'Home.aspx', canvasContent: [] }).valid).toBe(true);
    expect(validate('pageFile', { name: 'Home.aspx' }).valid).toBe(false);
  });

  it('accepts a minimal files meta file and rejects a bad version label', () => {
    const ok = { listKey: 'Dokumentumok', files: [{ path: '2026/a.docx', sizeBytes: 10, blob: 'files/Dokumentumok/2026/a.docx' }] };
    expect(validate('filesMetaFile', ok).valid).toBe(true);
    const bad = { listKey: 'Dokumentumok', files: [{ path: 'a.docx', sizeBytes: 1, versions: [{ label: 'v1', blob: 'x' }] }] };
    expect(validate('filesMetaFile', bad).errors.map((e) => e.path)).toContain('/files/0/versions/0/label');
  });
});

describe('invalid templates are rejected', () => {
  it('missing meta', () => {
    const r = validateTemplate(readJson(join(invalid, 'missing-meta.manifest.json')));
    expect(r.valid).toBe(false);
    expect(r.errors).toContainEqual(expect.objectContaining({ path: '', keyword: 'required', params: { missingProperty: 'meta' } }));
  });

  it('Choice field without choices', () => {
    const r = validateTemplate(readJson(join(invalid, 'choice-without-choices.manifest.json')));
    expect(r.valid).toBe(false);
    expect(r.errors).toEqual([
      expect.objectContaining({ path: '/siteFields/0', keyword: 'required', params: { missingProperty: 'choices' } })
    ]);
  });

  it('content source pointing outside the package', () => {
    const r = validateTemplate(readJson(join(invalid, 'content-path-traversal.manifest.json')));
    expect(r.valid).toBe(false);
    expect(r.errors).toEqual([expect.objectContaining({ path: '/lists/0/content/source', keyword: 'pattern' })]);
  });

  it('every fixture in tests/fixtures/invalid is rejected', () => {
    for (const f of readdirSync(invalid)) {
      expect(validateTemplate(readJson(join(invalid, f))).valid).toBe(false);
    }
  });
});

describe('migrate', () => {
  const example = (): Record<string, unknown> => readJson(join(examples, 'manifest.example.json'));

  it('returns a current-version template unchanged', () => {
    const t = example();
    expect(migrate(t)).toEqual(t);
    expect(CURRENT_SCHEMA_VERSION).toBe('1.0');
  });

  it.each([
    ['2.0', 'SCHEMA_MAJOR_MISMATCH'],
    ['1.99', 'SCHEMA_TOO_NEW'],
    ['abc', 'SCHEMA_VERSION_INVALID']
  ])('rejects schemaVersion %s with %s', (v, code) => {
    const t = { ...example(), schemaVersion: v };
    expect(() => migrate(t)).toThrow(CopyJetError);
    expect(() => migrate(t)).toThrow(expect.objectContaining({ code }));
  });
});
