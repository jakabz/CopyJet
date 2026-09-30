import { toItemField, toTemplateValue } from '../../src/core/items';
import type { IField, ITermRef } from '../../src/core/model';
import { TermCollector, TermStoreClient, fillTermSetPath, termPaths } from '../../src/core/taxonomy';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';

// The Forrás term store of spike 11 A (IDs shortened to readable GUIDs).
const STORE = '755cfac8-49b6-4850-b75a-2d387c7676ac';
const SET = 'caaf4dd7-a354-4085-b190-05c20a16d9e2';
const PENZUGY = 'dae0f266-1da5-41e6-a7bd-896ae0a8cf92';
const LOGISZTIKA = '3483fadf-9a8c-4b7e-b8d4-bb0cd59c4830';
const RAKTAR = 'e8da65f6-be0e-4e28-b281-43096676e544';

const label = (name: string): Array<{ name: string; languageTag: string; isDefault: boolean }> => [{ name, languageTag: 'en-US', isDefault: true }];

function termStore(): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  return createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (!/\/_api\/v2\.1\/termStore/i.test(req.url)) return undefined;
    if (/\/termStore\?\$select=id$/i.test(req.url)) return { body: { id: STORE.toUpperCase() } };
    if ((m = /\/sets\/([0-9a-f-]+)\?\$expand=parentGroup$/i.exec(req.url))) {
      return m[1] === SET ? { body: { id: SET, localizedNames: label('Terület'), parentGroup: { id: 'g1', name: 'CopyJet teszt' } } } : { status: 404, body: {} };
    }
    if (new RegExp(`/sets/${SET}/children\\?`, 'i').test(req.url)) {
      return { body: { value: [{ id: LOGISZTIKA, labels: label('Logisztika'), childrenCount: 1 }, { id: PENZUGY, labels: label('Pénzügy'), childrenCount: 0 }] } };
    }
    if (new RegExp(`/sets/${SET}/terms/${LOGISZTIKA}/children\\?`, 'i').test(req.url)) return { body: { value: [{ id: RAKTAR, labels: label('Raktár'), childrenCount: 0 }] } };
    if (/\/groups\?\$select=id,name$/i.test(req.url)) return { body: { value: [{ id: 'g0', name: 'People' }, { id: 'g1', name: 'CopyJet teszt' }] } };
    if (/\/groups\/g1\/sets\?/i.test(req.url)) return { body: { value: [{ id: SET, localizedNames: label('Terület') }] } };
    return undefined;
  });
}

describe('TermStoreClient', () => {
  it('walks a term set into terms with parents and label paths', async () => {
    const { sp, requests } = termStore();
    const client = new TermStoreClient(sp);
    expect(await client.storeId()).toBe(STORE);
    const set = (await client.set(SET.toUpperCase()))!;
    expect(set).toEqual({ id: SET, name: 'Terület', groupName: 'CopyJet teszt' });
    const terms = await client.terms(SET);
    expect(terms).toEqual([
      { id: LOGISZTIKA, label: 'Logisztika', parentId: undefined },
      { id: PENZUGY, label: 'Pénzügy', parentId: undefined },
      { id: RAKTAR, label: 'Raktár', parentId: LOGISZTIKA }
    ]);
    expect(termPaths(set, terms)[RAKTAR]).toBe('CopyJet teszt/Terület/Logisztika/Raktár');
    // Cached: a second read asks nothing.
    const before = requests.length;
    await client.terms(SET);
    expect(requests.length).toBe(before);
    expect(await client.set('00000000-0000-0000-0000-000000000001')).toBeUndefined();
  });

  it('finds a term set by group and set name (cross-tenant)', async () => {
    const client = new TermStoreClient(termStore().sp);
    expect(await client.findSet('CopyJet teszt', 'Terület')).toEqual({ id: SET, name: 'Terület', groupName: 'CopyJet teszt' });
    expect(await client.findSet('CopyJet teszt', 'Nincs')).toBeUndefined();
    expect(await client.findSet('Nincs', 'Terület')).toBeUndefined();
  });
});

describe('Managed Metadata extraction', () => {
  it('turns REST values into term keys with labels from the store (the single value "Label" is a WssId)', async () => {
    const terms: ITermRef[] = [];
    const collector = new TermCollector(terms, new TermStoreClient(termStore().sp));
    expect(await collector.prepare([SET, '00000000-0000-0000-0000-000000000009'])).toEqual(['00000000-0000-0000-0000-000000000009']);
    const ctx = { tokens: new TokenContext(), principal: () => undefined, term: (id: string) => collector.key(id) };
    const single = toItemField({ InternalName: 'Terulet', TypeAsString: 'TaxonomyFieldType' })!;
    const multi = toItemField({ InternalName: 'Cimkek', TypeAsString: 'TaxonomyFieldTypeMulti' })!;
    expect(toTemplateValue(single, { Label: '8', TermGuid: PENZUGY, WssId: 8 }, ctx)).toEqual({ terms: ['Penzugy'] });
    expect(
      toTemplateValue(multi, [{ Label: 'Logisztika', TermGuid: LOGISZTIKA, WssId: 9 }, { Label: 'Raktár', TermGuid: RAKTAR, WssId: 10 }, { Label: '?', TermGuid: '00000000-0000-0000-0000-00000000000a', WssId: 11 }], ctx)
    ).toEqual({ terms: ['Logisztika', 'Raktar'] });
    expect(toTemplateValue(single, null, ctx)).toBeUndefined();
    expect(toTemplateValue(multi, [], ctx)).toBeUndefined();
    expect(terms).toEqual([
      { key: 'Penzugy', termId: PENZUGY, termSetId: SET, label: 'Pénzügy', path: 'CopyJet teszt/Terület/Pénzügy' },
      { key: 'Logisztika', termId: LOGISZTIKA, termSetId: SET, label: 'Logisztika', path: 'CopyJet teszt/Terület/Logisztika' },
      { key: 'Raktar', termId: RAKTAR, termSetId: SET, label: 'Raktár', path: 'CopyJet teszt/Terület/Logisztika/Raktár' }
    ]);
  });

  it('fills a column\'s term set path, or warns when the set is gone', async () => {
    const client = new TermStoreClient(termStore().sp);
    const def = { internalName: 'Terulet', type: 'TaxonomyFieldType', title: 'Terület', schemaXml: '<Field />', termSet: { termSetId: SET, termStoreId: STORE, path: '' } } as IField;
    const warnings: string[] = [];
    await fillTermSetPath(def, client, (m) => warnings.push(m));
    expect(def.termSet!.path).toBe('CopyJet teszt/Terület');
    const gone = { ...def, termSet: { termSetId: 'ffffffff-0000-4000-8000-000000000001', path: '' } } as IField;
    await fillTermSetPath(gone, client, (m) => warnings.push(m));
    expect(gone.termSet!.path).toBe('');
    expect(warnings.length).toBe(1);
  });
});
