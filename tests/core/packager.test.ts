import { readFileSync } from 'fs';
import { join } from 'path';
import { JsonTemplateWriter, createEmptyTemplate, openTemplate } from '../../src/core/packager';
import { CopyJetError, AbortError } from '../../src/core/errors';

const exampleText = readFileSync(join(__dirname, '..', '..', 'schema', 'examples', 'manifest.example.json'), 'utf8');

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
  it('opens the example manifest (with BOM)', async () => {
    const reader = await openTemplate(new Blob(['﻿' + exampleText]));
    expect(reader.manifest.lists.map((l) => l.key)).toEqual(['Ugyfelek', 'Projektek', 'Dokumentumok']);
    expect(reader.has('items/Projektek.json')).toBe(false);
    await expect(reader.getJson('items/Projektek.json')).rejects.toMatchObject({ code: 'PACKAGE_ENTRY_NOT_FOUND' });
  });

  it.each([
    ['PK\u0003\u0004zipdata', 'TEMPLATE_ZIP_UNSUPPORTED'],
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
