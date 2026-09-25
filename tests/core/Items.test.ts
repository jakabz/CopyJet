import { ItemExtractor, ITEM_PAGE_SIZE } from '../../src/core/extractors/ItemExtractor';
import { ItemLookupProvider } from '../../src/core/providers/ItemLookupProvider';
import { ItemProvider } from '../../src/core/providers/ItemProvider';
import { listItemsDefs, lookupTargetsWithoutContent, type IListItemsDef } from '../../src/core/items';
import { loadSourceSite, toListDef } from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import { PrincipalMapper } from '../../src/core/mapping';
import type { IInstallContext, IItemsFile, ITemplateReader } from '../../src/core/model';
import { ZipTemplateWriter, createEmptyTemplate, openTemplate } from '../../src/core/packager';
import { buildPlan } from '../../src/core/planner';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { SOURCE_WEB, lookupForras, sourceLists, tesztLista } from '../fixtures/lists';

const TEST_LIST = tesztLista.RootFolder.ServerRelativeUrl; // /sites/Forras/Lists/Teszt lista
const LOOKUP_LIST = lookupForras.RootFolder.ServerRelativeUrl;
const ANNA = 'i:0#.f|membership|anna@contoso.com';
// A custom site content type and its list instances on the source and the target (parent + "00" + GUID).
const SITE_CT = `0x0100${'A'.repeat(32)}`;
const SOURCE_LIST_CT = `${SITE_CT}00${'B'.repeat(32)}`;
const TARGET_LIST_CT = `${SITE_CT}00${'C'.repeat(32)}`;
// Item on a list can be 0x01 + 00 + GUID + 00 + GUID; only the REST Parent tells it is 0x01 (spike 08 F).
const SOURCE_ITEM_CT = `0x0100${'E'.repeat(32)}00${'F'.repeat(32)}`;
const TARGET_ITEM_CT = `0x0100${'1'.repeat(32)}00${'2'.repeat(32)}`;
const listCts = (pairs: Array<[string, string]>): Array<{ StringId: string; Parent: { StringId: string } }> => pairs.map(([StringId, parent]) => ({ StringId, Parent: { StringId: parent } }));

type Raw = { [k: string]: unknown };

const sourceFields = [
  { InternalName: 'Title', TypeAsString: 'Text', Hidden: false, ReadOnlyField: false },
  { InternalName: 'CJNum', TypeAsString: 'Number', Hidden: false, ReadOnlyField: false },
  { InternalName: 'CJDate', TypeAsString: 'DateTime', Hidden: false, ReadOnlyField: false },
  { InternalName: 'Valaki', TypeAsString: 'User', Hidden: false, ReadOnlyField: false },
  { InternalName: 'ListaLookup', TypeAsString: 'Lookup', Hidden: false, ReadOnlyField: false },
  { InternalName: 'Terulet', TypeAsString: 'TaxonomyFieldType', Hidden: false, ReadOnlyField: false },
  { InternalName: 'Author', TypeAsString: 'User', Hidden: false, ReadOnlyField: true },
  { InternalName: 'ContentType', TypeAsString: 'Computed', Hidden: false, ReadOnlyField: false }
];

const sourceItems: { [listUrl: string]: Raw[] } = {
  [TEST_LIST]: [
    { ID: 1, FSObjType: 0, FileDirRef: TEST_LIST, ContentTypeId: SOURCE_LIST_CT, AuthorId: 7, EditorId: 9, Created: '2026-03-02T08:15:00Z', Modified: '2026-07-16T11:30:00Z', Title: 'Első', CJNum: 12.5, CJDate: '2026-03-02T08:15:00Z', ValakiId: 7, ListaLookupId: 2 },
    { ID: 2, FSObjType: 1, FileDirRef: TEST_LIST, Title: '2026' },
    { ID: 3, FSObjType: 0, FileDirRef: `${TEST_LIST}/2026/Q1`, ContentTypeId: SOURCE_ITEM_CT, AuthorId: 7, EditorId: 7, Created: '2026-01-15T10:00:00Z', Modified: '2026-01-15T10:00:00Z', Title: 'Mappában', CJNum: null, ValakiId: null, ListaLookupId: null }
  ],
  [LOOKUP_LIST]: [
    { ID: 1, FSObjType: 0, FileDirRef: LOOKUP_LIST, ContentTypeId: SOURCE_ITEM_CT, AuthorId: 7, EditorId: 7, Created: '2026-01-01T10:00:00Z', Modified: '2026-01-01T10:00:00Z', Title: 'Egy' },
    { ID: 2, FSObjType: 0, FileDirRef: LOOKUP_LIST, ContentTypeId: SOURCE_ITEM_CT, AuthorId: 7, EditorId: 7, Created: '2026-01-01T10:00:00Z', Modified: '2026-01-01T10:00:00Z', Title: 'Kettő' }
  ]
};

function sourceSp(items = sourceItems): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  return createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (req.method === 'GET' && /\/_api\/web\?\$select=/i.test(req.url)) return { body: SOURCE_WEB };
    if (req.method === 'GET' && /\/_api\/web\/lists\?\$select=/i.test(req.url)) return { body: sourceLists };
    if (req.method === 'GET' && /\/_api\/web\/siteusers\?\$select=/i.test(req.url)) {
      return {
        body: [
          { Id: 7, LoginName: ANNA, Title: 'Kiss Anna', Email: 'anna@contoso.com' },
          { Id: 9, LoginName: 'SHAREPOINT\\system', Title: 'Rendszerfiók' }
        ]
      };
    }
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/fields\?\$select=/i.test(req.url)) return { body: sourceFields };
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/rootFolder\?\$select=ContentTypeOrder$/i.test(req.url)) return { body: { ContentTypeOrder: [{ StringValue: SOURCE_ITEM_CT }] } };
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/contentTypes\?\$select=StringId,Parent\/StringId&\$expand=Parent$/i.test(req.url)) {
      return { body: listCts([[SOURCE_LIST_CT, SITE_CT], [SOURCE_ITEM_CT, '0x01']]) };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\/items\?.*\$filter=ID gt (\d+).*\$top=(\d+)/i.exec(req.url))) {
      const after = Number(m[2]);
      return { body: (items[m[1]] || []).filter((i) => (i.ID as number) > after).slice(0, Number(m[3])) };
    }
    return undefined;
  }, SOURCE_WEB.Url);
}

async function extractPackage(): Promise<ITemplateReader> {
  const { sp } = sourceSp();
  const site = await loadSourceSite(sp);
  const writer = new ZipTemplateWriter(
    createEmptyTemplate({
      name: 'Tartalom',
      createdBy: 'anna@contoso.com',
      createdAt: '2026-09-25T10:00:00Z',
      sourceSiteUrl: SOURCE_WEB.Url,
      sourceTenant: 'contoso.sharepoint.com',
      sourceLcid: 1038,
      includesContent: false
    })
  );
  const teszt = toListDef(tesztLista, 'Teszt_lista', SOURCE_WEB.ServerRelativeUrl);
  teszt.fields = [{ internalName: 'ListaLookup', type: 'Lookup', title: 'ListaLookup', lookupList: '{{listkey:Teszt_lookup_forrs}}', lookupField: 'Title', schemaXml: '<Field />' }];
  writer.manifest.lists.push(teszt, toListDef(lookupForras, 'Teszt_lookup_forrs', SOURCE_WEB.ServerRelativeUrl));
  const log = new Logger();
  await new ItemExtractor().extract(
    sp,
    [
      { kind: 'items', key: 'items:Teszt_lista' },
      { kind: 'items', key: 'items:Teszt_lookup_forrs' }
    ],
    { includeContent: true, includeVersions: false, includeMembers: false, tokens: site.tokens, log },
    writer
  );
  expect(log.entries.filter((e) => e.level === 'warn').map((e) => e.code)).toEqual(['ITEM_FIELD_NOT_COPIED', 'ITEM_FIELD_NOT_COPIED']);
  return openTemplate(await writer.finalize());
}

describe('ItemExtractor', () => {
  it('writes items/<key>.json, marks the list content and collects principals', async () => {
    const reader = await extractPackage();
    const teszt = reader.manifest.lists.filter((l) => l.key === 'Teszt_lista')[0];
    expect(teszt.content).toEqual({ mode: 'items', source: 'items/Teszt_lista.json', itemCount: 2, includeAttachments: false });
    expect(reader.manifest.meta.includesContent).toBe(true);
    expect(reader.manifest.principals).toEqual([{ key: 'anna', kind: 'user', loginName: ANNA, email: 'anna@contoso.com', displayName: 'Kiss Anna' }]);
    const file = await reader.getJson<IItemsFile>('items/Teszt_lista.json');
    expect(file.items).toEqual([
      {
        sourceId: 1,
        contentType: SITE_CT,
        values: { Title: 'Első', CJNum: 12.5, CJDate: '2026-03-02T08:15:00Z', Valaki: { principals: ['{principal:anna}'] }, ListaLookup: { lookup: [2] } },
        // The system account is not carried: the installing user becomes the editor.
        system: { author: '{principal:anna}', created: '2026-03-02T08:15:00Z', modified: '2026-07-16T11:30:00Z' }
      },
      {
        sourceId: 3,
        folder: '2026/Q1',
        contentType: '0x01',
        values: { Title: 'Mappában' },
        system: { author: '{principal:anna}', editor: '{principal:anna}', created: '2026-01-15T10:00:00Z', modified: '2026-01-15T10:00:00Z' }
      }
    ]);
  });

  it('reads large lists in ID-filtered pages', async () => {
    const many: Raw[] = [];
    for (let id = 1; id <= ITEM_PAGE_SIZE + 5; id++) many.push({ ID: id, FSObjType: 0, FileDirRef: TEST_LIST, Title: `T${id}` });
    const { sp, requests } = sourceSp({ [TEST_LIST]: many });
    const site = await loadSourceSite(sp);
    const writer = new ZipTemplateWriter(createEmptyTemplate({ name: 'N', createdBy: 'a', createdAt: '2026-09-25T10:00:00Z', sourceSiteUrl: SOURCE_WEB.Url, sourceTenant: 'c', sourceLcid: 1038, includesContent: false }));
    writer.manifest.lists.push(toListDef(tesztLista, 'Teszt_lista', SOURCE_WEB.ServerRelativeUrl));
    await new ItemExtractor().extract(sp, [{ kind: 'items', key: 'items:Teszt_lista' }], { includeContent: true, includeVersions: false, includeMembers: false, tokens: site.tokens, log: new Logger() }, writer);
    expect(writer.manifest.lists[0].content.itemCount).toBe(ITEM_PAGE_SIZE + 5);
    expect(requests.filter((r) => /\/items\?/.test(r.url)).map((r) => /ID gt (\d+)/.exec(r.url)![1])).toEqual(['0', String(ITEM_PAGE_SIZE)]);
  });
});

// ---------------------------------------------------------------------------------------------------------
// Target web: in-memory lists with items, folders and the calls the providers make.

const TARGET = { Url: 'https://fabrikam.sharepoint.com/sites/Cel', ServerRelativeUrl: '/sites/Cel', Title: 'Cél' };
const T_TEST = '/sites/Cel/Lists/Teszt lista';
const T_LOOKUP = '/sites/Cel/Lists/Teszt lookup forrs';
const T_LOOKUP_ID = 'aaaaaaaa-0000-4000-8000-000000000002';

interface ITargetList {
  items: Raw[];
  folders: string[];
  fields: Array<{ InternalName: string; TypeAsString: string; Hidden: boolean; ReadOnlyField: boolean; SchemaXml: string }>;
}

function budapest(iso: string): string {
  const d = new Date(Date.parse(iso) + (/-0[4-9]-|-10-(0|1|2[0-4])/.test(iso) ? 2 : 1) * 3600000);
  return d.toISOString().slice(0, 19);
}

function targetSp(
  lists: { [url: string]: ITargetList },
  knownUsers: string[] = [ANNA],
  regional = { LocaleId: 1038, DecimalSeparator: ',', Time24: true }
): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  let seq = 100;
  return createMockSp((req) => {
    let m: RegExpExecArray | null;
    const list = (url: string): ITargetList | undefined => lists[url];
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\?\$select=ItemCount$/i.exec(req.url))) {
      const l = list(m[1]);
      return l ? { body: { ItemCount: l.items.length + l.folders.length } } : { status: 404, body: {} };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\?\$select=Id$/i.exec(req.url))) {
      return list(m[1]) ? { body: { Id: 'x' } } : { status: 404, body: {} };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\/items\?\$filter=FSObjType eq 1&\$select=FileRef/i.exec(req.url))) {
      return { body: list(m[1])!.folders.map((f) => ({ FileRef: `${m![1]}/${f}` })) };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\/items\?\$filter=FSObjType eq 0/i.exec(req.url))) {
      return { body: list(m[1])!.items.slice(0, 1).map((i) => ({ Id: i.ID })) };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\/fields\?\$select=/i.exec(req.url))) return { body: list(m[1])!.fields };
    if (req.method === 'GET' && /\/_api\/web\/regionalsettings\?\$select=/i.test(req.url)) return { body: regional };
    if (req.method === 'POST' && (m = /\/regionalsettings\/timezone\/utctolocaltime\('([^']+)'\)/i.exec(req.url))) return { body: { value: budapest(m[1]) } };
    if (req.method === 'POST' && /\/_api\/web\/ensureuser$/i.test(req.url)) {
      const login = (req.body as { logonName: string }).logonName;
      return knownUsers.indexOf(login) >= 0 ? { body: { Id: 5, LoginName: login } } : { status: 500, body: { 'odata.error': { message: { value: 'not found' } } } };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\/rootFolder\?\$select=ContentTypeOrder$/i.exec(req.url))) {
      return { body: { ContentTypeOrder: [{ StringValue: TARGET_ITEM_CT }, { StringValue: TARGET_LIST_CT }] } };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\/contentTypes\?\$select=StringId,Parent\/StringId&\$expand=Parent$/i.exec(req.url))) {
      return { body: listCts([[TARGET_ITEM_CT, '0x01'], [TARGET_LIST_CT, SITE_CT], [`0x012000${'D'.repeat(32)}`, '0x0120']]) };
    }
    if (req.method === 'POST' && (m = /\/getList\('([^']+)'\)\/AddValidateUpdateItemUsingPath\(\)$/i.exec(req.url))) {
      const l = list(m[1])!;
      const body = req.body as {
        formValues: Array<{ FieldName: string; FieldValue: string }>;
        listItemCreateInfo: { FolderPath: { DecodedUrl: string }; LeafName?: { DecodedUrl: string }; UnderlyingObjectType?: number };
      };
      const folder = body.listItemCreateInfo.FolderPath.DecodedUrl;
      if (body.listItemCreateInfo.UnderlyingObjectType === 1) {
        // A list folder (spike 08 E): its parent must exist.
        const parent = folder === m[1] ? '' : folder.slice(m[1].length + 1);
        if (parent && l.folders.indexOf(parent) < 0) return { status: 500, body: { 'odata.error': { message: { value: 'folder missing' } } } };
        l.folders.push(parent ? `${parent}/${body.listItemCreateInfo.LeafName!.DecodedUrl}` : body.listItemCreateInfo.LeafName!.DecodedUrl);
        return { body: { value: [{ FieldName: 'Id', FieldValue: String(++seq), HasException: false }] } };
      }
      if (folder !== m[1] && l.folders.indexOf(folder.slice(m[1].length + 1)) < 0) return { status: 500, body: { 'odata.error': { message: { value: 'folder missing' } } } };
      const bad = body.formValues.filter((v) => v.FieldName === 'CJNum' && /\./.test(v.FieldValue));
      if (bad.length) return { body: { value: [{ FieldName: 'CJNum', FieldValue: bad[0].FieldValue, HasException: true, ErrorMessage: 'Itt csak számok szerepelhetnek.' }] } };
      const item: Raw = { ID: ++seq, folder };
      body.formValues.forEach((v) => (item[v.FieldName] = v.FieldValue));
      l.items.push(item);
      return { body: { value: body.formValues.map((v) => ({ ...v, HasException: false })).concat([{ FieldName: 'Id', FieldValue: String(item.ID), HasException: false }]) } };
    }
    if (req.method === 'POST' && (m = /\/getList\('([^']+)'\)\/items\((\d+)\)\/ValidateUpdateListItem(\(\))?$/i.exec(req.url))) {
      const item = list(m[1])!.items.filter((i) => i.ID === Number(m![2]))[0];
      const body = req.body as { formValues: Array<{ FieldName: string; FieldValue: string }> };
      body.formValues.forEach((v) => (item[v.FieldName] = v.FieldValue));
      return { body: { value: body.formValues.map((v) => ({ ...v, HasException: false })) } };
    }
    return undefined;
  }, TARGET.Url);
}

const textField = (InternalName: string, TypeAsString = 'Text'): ITargetList['fields'][number] => ({ InternalName, TypeAsString, Hidden: false, ReadOnlyField: false, SchemaXml: `<Field Name="${InternalName}" />` });

function emptyTarget(): { [url: string]: ITargetList } {
  return {
    [T_TEST]: {
      items: [],
      folders: [],
      fields: [
        textField('Title'),
        textField('CJNum', 'Number'),
        textField('CJDate', 'DateTime'),
        textField('Valaki', 'User'),
        { InternalName: 'ListaLookup', TypeAsString: 'Lookup', Hidden: false, ReadOnlyField: false, SchemaXml: `<Field Name="ListaLookup" List="{${T_LOOKUP_ID.toUpperCase()}}" ShowField="Title" />` }
      ]
    },
    [T_LOOKUP]: { items: [], folders: [], fields: [textField('Title')] }
  };
}

function installContext(reader: ITemplateReader, sp: ReturnType<typeof createMockSp>['sp']): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: TARGET.Url, serverRelativeUrl: TARGET.ServerRelativeUrl, title: TARGET.Title });
  tokens.set('listkey', 'Teszt_lista', 'aaaaaaaa-0000-4000-8000-000000000001');
  tokens.set('listkey', 'Teszt_lookup_forrs', T_LOOKUP_ID);
  return { targetSiteUrl: TARGET.Url, tokens, log: new Logger(), content: { reader, idMaps: {}, principals: new PrincipalMapper(sp, reader.manifest.principals) } };
}

describe('ItemProvider + ItemLookupProvider', () => {
  it('plans items after their lists and columns, lookups after the items they point to', async () => {
    const reader = await extractPackage();
    const plan = buildPlan(reader.manifest);
    const step = (key: string) => plan.steps.filter((s) => s.ref.key === key)[0];
    expect(step('items:Teszt_lista').dependsOn).toEqual(['list:Teszt_lista', 'listField:Teszt_lista/ListaLookup']);
    expect(step('itemLookups:Teszt_lista').dependsOn).toEqual(['items:Teszt_lista', 'items:Teszt_lookup_forrs']);
    expect(step('itemLookups:Teszt_lookup_forrs')).toBeUndefined();
  });

  it('writes items in the web locale, creates folders, then maps lookups through the ID maps', async () => {
    const reader = await extractPackage();
    const lists = emptyTarget();
    const { sp } = targetSp(lists);
    const ctx = installContext(reader, sp);
    const defs = listItemsDefs(reader.manifest);
    const def = (key: string): IListItemsDef => defs.filter((d) => d.listKey === key)[0];
    const items = new ItemProvider({ batched: false });

    expect((await items.diff(sp, def('Teszt_lista'), ctx)).status).toBe('new');
    expect(await items.apply(sp, def('Teszt_lookup_forrs'), 'skip', ctx)).toMatchObject({ outcome: 'created' });
    expect(await items.apply(sp, def('Teszt_lista'), 'skip', ctx)).toMatchObject({ outcome: 'created' });
    // The two folders (2026, 2026/Q1) are list items too and take IDs 103 and 104.
    expect(ctx.content!.idMaps).toEqual({ Teszt_lookup_forrs: { 1: 101, 2: 102 }, Teszt_lista: { 1: 105, 3: 106 } });

    const [first, inFolder] = lists[T_TEST].items;
    expect(first).toEqual({
      ID: 105,
      folder: T_TEST,
      Title: 'Első',
      CJNum: '12,5',
      CJDate: '2026. 03. 02. 9:15',
      Valaki: `[{"Key":"${ANNA}"}]`,
      ContentTypeId: TARGET_LIST_CT,
      Author: `[{"Key":"${ANNA}"}]`,
      Created: '2026. 03. 02. 9:15',
      Modified: '2026. 07. 16. 13:30'
    });
    expect(inFolder.folder).toBe(`${T_TEST}/2026/Q1`);
    expect(inFolder.ContentTypeId).toBe(TARGET_ITEM_CT);
    expect(lists[T_TEST].folders).toEqual(['2026', '2026/Q1']);

    const lookups = new ItemLookupProvider({ batched: false });
    expect(await lookups.apply(sp, def('Teszt_lista'), 'skip', ctx)).toMatchObject({ outcome: 'updated' });
    // Source item 2 of the lookup list is target item 102; Editor/Modified are sent again to keep them.
    expect(first.ListaLookup).toBe('102');
    expect(first.Modified).toBe('2026. 07. 16. 13:30');
    expect(inFolder.ListaLookup).toBeUndefined();
    expect(ctx.log.counts.warn).toBe(0);
  });

  it('never adds items to a list that already has some', async () => {
    const reader = await extractPackage();
    const lists = emptyTarget();
    lists[T_LOOKUP].items.push({ ID: 1, Title: 'Már itt' });
    const { sp, requests } = targetSp(lists);
    const ctx = installContext(reader, sp);
    const def = listItemsDefs(reader.manifest).filter((d) => d.listKey === 'Teszt_lookup_forrs')[0];
    expect(await new ItemProvider({ batched: false }).apply(sp, def, 'update', ctx)).toMatchObject({ outcome: 'skipped' });
    expect(ctx.log.entries.map((e) => e.code)).toEqual(['ITEMS_TARGET_NOT_EMPTY']);
    expect(requests.filter((r) => r.method === 'POST')).toEqual([]);
    expect(lists[T_LOOKUP].items.length).toBe(1);
  });

  it('treats a list holding only folders as empty', async () => {
    const reader = await extractPackage();
    const lists = emptyTarget();
    lists[T_TEST].folders.push('2026');
    const { sp } = targetSp(lists);
    const def = listItemsDefs(reader.manifest).filter((d) => d.listKey === 'Teszt_lista')[0];
    expect((await new ItemProvider({ batched: false }).diff(sp, def, installContext(reader, sp))).status).toBe('new');
  });

  it('leaves unknown people and lookups without ID maps empty, with one warning each', async () => {
    const reader = await extractPackage();
    const lists = emptyTarget();
    const { sp } = targetSp(lists, []);
    const ctx = installContext(reader, sp);
    const def = listItemsDefs(reader.manifest).filter((d) => d.listKey === 'Teszt_lista')[0];
    await new ItemProvider({ batched: false }).apply(sp, def, 'skip', ctx);
    await new ItemLookupProvider({ batched: false }).apply(sp, def, 'skip', ctx);
    expect(lists[T_TEST].items[0].Valaki).toBeUndefined();
    expect(lists[T_TEST].items[0].Author).toBeUndefined();
    expect(lists[T_TEST].items[0].ListaLookup).toBeUndefined();
    expect(ctx.log.entries.filter((e) => e.level === 'warn').map((e) => e.code)).toEqual(['PRINCIPAL_NOT_FOUND', 'ITEM_LOOKUP_UNRESOLVED']);
  });

  it('reports items SharePoint rejects and fails the step when none could be written', async () => {
    const reader = await extractPackage();
    const lists = emptyTarget();
    // The web says en-US ("12.5"), but the list rejects a dot – every item fails like on a hu-HU web (spike 08 C).
    const { sp } = targetSp(lists, [ANNA], { LocaleId: 1033, DecimalSeparator: '.', Time24: false });
    const ctx = installContext(reader, sp);
    const def = listItemsDefs(reader.manifest).filter((d) => d.listKey === 'Teszt_lista')[0];
    const file = await reader.getJson<IItemsFile>(def.source);
    file.items.forEach((i) => (i.values.CJNum = 1.5));
    await expect(new ItemProvider({ batched: false }).apply(sp, def, 'skip', ctx)).rejects.toMatchObject({ code: 'ITEMS_FAILED' });
    expect(ctx.log.entries.filter((e) => e.code === 'ITEM_FAILED').map((e) => e.message)).toEqual([
      'Item 1 could not be written: CJNum: Itt csak számok szerepelhetnek.',
      'Item 3 could not be written: CJNum: Itt csak számok szerepelhetnek.'
    ]);
    expect(lists[T_TEST].items).toEqual([]);
  });
});

describe('lookupTargetsWithoutContent', () => {
  it('names template lists that content lists look up to but that travel without items', async () => {
    const reader = await extractPackage();
    const template = { ...reader.manifest, lists: reader.manifest.lists.map((l) => ({ ...l, content: { mode: 'none' as const } })) };
    expect(lookupTargetsWithoutContent(template, ['Teszt_lista'])).toEqual(['Teszt_lookup_forrs']);
    expect(lookupTargetsWithoutContent(template, ['Teszt_lista', 'Teszt_lookup_forrs'])).toEqual([]);
    expect(lookupTargetsWithoutContent(template, ['Teszt_lookup_forrs'])).toEqual([]);
    // A target outside the template is the missing-dependency check's business, not this one's.
    expect(lookupTargetsWithoutContent({ ...template, lists: template.lists.filter((l) => l.key === 'Teszt_lista') }, ['Teszt_lista'])).toEqual([]);
  });
});
