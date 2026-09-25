import { readFileSync } from 'fs';
import { join } from 'path';
import JSZip from 'jszip';
import { JsonTemplateWriter, ZipTemplateWriter, checkEntryPath, createEmptyTemplate, openTemplate } from '../../src/core/packager';
import { CopyJetError, AbortError } from '../../src/core/errors';

const exampleText = readFileSync(join(__dirname, '..', '..', 'schema', 'examples', 'manifest.example.json'), 'utf8');
const exampleItems = JSON.parse(readFileSync(join(__dirname, '..', '..', 'schema', 'examples', 'items.Projektek.example.json'), 'utf8').replace(/^\uFEFF/, ''));
/** The example manifest without its content, as a structure-only .json template. */
const structureOnly = (): string => {
  const t = JSON.parse(exampleText.replace(/^\uFEFF/, ''));
  t.lists.forEach((l: { content: unknown }) => (l.content = { mode: 'none' }));
  return JSON.stringify(t);
};

const meta = {
  name: 'Teszt sablon',
  createdBy: 'anna@contoso.com',
  createdAt: '2026-09-24T10:00:00Z',
  sourceSiteUrl: 'https://contoso.sharepoint.com/sites/projekt',
  sourceTenant: 'contoso.onmicrosoft.com',
  sourceLcid: 1038,
  includesContent: false
};

describe('JsonTemplateWriter', () => {
  it('writes a valid manifest that openTemplate reads back', async () => {
    const writer = new JsonTemplateWriter(createEmptyTemplate(meta));
    writer.manifest.lists.push({ key: 'Ugyfelek', url: 'Lists/Ugyfelek', title: 'Ügyfelek', template: 100, content: { mode: 'none' } });
    const blob = await writer.finalize();
    expect(blob.type).toBe('application/json');

    const reader = await openTemplate(blob);
    expect(reader.manifest).toEqual(writer.manifest);
    expect(reader.manifest.meta.generator.name).toBe('CopyJet');
  });

  it('refuses package entries and invalid manifests', async () => {
    const writer = new JsonTemplateWriter(createEmptyTemplate(meta));
    expect(() => writer.addJson('items/x.json', {})).toThrow(expect.objectContaining({ code: 'PACKAGE_JSON_NO_ENTRIES' }));
    writer.manifest.meta.name = '';
    await expect(writer.finalize()).rejects.toMatchObject({ code: 'TEMPLATE_INVALID' });
  });

  it('honours an aborted signal', async () => {
    const ac = new AbortController();
    ac.abort();
    await expect(new JsonTemplateWriter(createEmptyTemplate(meta)).finalize(ac.signal)).rejects.toBeInstanceOf(AbortError);
  });
});

describe('openTemplate', () => {
  it('opens a structure-only manifest (with BOM)', async () => {
    const reader = await openTemplate(new Blob(['\uFEFF' + structureOnly()]));
    expect(reader.manifest.lists.map((l) => l.key)).toEqual(['Ugyfelek', 'Projektek', 'Dokumentumok']);
    expect(reader.has('items/Projektek.json')).toBe(false);
    await expect(reader.getJson('items/Projektek.json')).rejects.toMatchObject({ code: 'PACKAGE_ENTRY_NOT_FOUND' });
  });

  it.each([
    ['PK\u0003\u0004zipdata', 'TEMPLATE_PARSE'],
    // A manifest that refers to items/files needs its package.
    [exampleText, 'TEMPLATE_CONTENT_MISSING'],
    ['{ not json', 'TEMPLATE_PARSE'],
    ['[1,2]', 'TEMPLATE_PARSE'],
    [JSON.stringify({ ...JSON.parse(exampleText), schemaVersion: '2.0' }), 'SCHEMA_MAJOR_MISMATCH'],
    [JSON.stringify({ ...JSON.parse(exampleText), lists: 'nope' }), 'TEMPLATE_INVALID']
  ])('rejects %#: %s', async (content, code) => {
    const p = openTemplate(new Blob([content]));
    await expect(p).rejects.toBeInstanceOf(CopyJetError);
    await expect(openTemplate(new Blob([content]))).rejects.toMatchObject({ code });
  });
});

describe('ZipTemplateWriter', () => {
  const withItems = (): ZipTemplateWriter => {
    const writer = new ZipTemplateWriter(createEmptyTemplate({ ...meta, includesContent: true }));
    writer.manifest.lists.push({ key: 'Projektek', url: 'Lists/Projektek', title: 'Projektek', template: 100, content: { mode: 'items', source: 'items/Projektek.json', itemCount: 2 } });
    writer.addJson('items/Projektek.json', exampleItems);
    return writer;
  };

  it('packs the manifest and its entries into a .zip that openTemplate reads back', async () => {
    const writer = withItems();
    writer.addBlob('attachments/Projektek/1/specifikacio.pdf', new Blob([new Uint8Array([1, 2, 3])]));
    const blob = await writer.finalize();
    expect(blob.type).toBe('application/zip');

    const reader = await openTemplate(blob);
    expect(reader.manifest).toEqual(writer.manifest);
    expect(reader.has('items/Projektek.json')).toBe(true);
    expect(await reader.getJson('items/Projektek.json')).toEqual(exampleItems);
    expect(new Uint8Array(await (await reader.getBlob('attachments/Projektek/1/specifikacio.pdf')).arrayBuffer())).toEqual(new Uint8Array([1, 2, 3]));
  });

  it('validates item entries against the schema', async () => {
    const writer = withItems();
    writer.addJson('items/Projektek.json', { listKey: 'Projektek', items: [{ values: {} }] });
    await expect(writer.finalize()).rejects.toMatchObject({ code: 'TEMPLATE_INVALID' });
  });

  it.each(['../x.json', '/items/x.json', 'items//x.json', 'items\\x.json', 'manifest.json', ''])('refuses the package path %j', (path) => {
    expect(() => checkEntryPath(path)).toThrow(expect.objectContaining({ code: 'PACKAGE_PATH_INVALID' }));
  });

  it('rejects packages without a manifest or with missing content entries', async () => {
    const empty = new JSZip();
    empty.file('readme.txt', 'x');
    await expect(openTemplate(new Blob([await empty.generateAsync({ type: 'uint8array' })]))).rejects.toMatchObject({ code: 'TEMPLATE_PARSE' });

    const writer = withItems();
    const zip = await JSZip.loadAsync(new Uint8Array(await (await writer.finalize()).arrayBuffer()));
    zip.remove('items/Projektek.json');
    await expect(openTemplate(new Blob([await zip.generateAsync({ type: 'uint8array' })]))).rejects.toMatchObject({ code: 'TEMPLATE_CONTENT_MISSING' });
  });
});
