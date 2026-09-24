import { ListExtractor } from '../../src/core/extractors/ListExtractor';
import { ListProvider } from '../../src/core/providers/ListProvider';
import {
  assignListKeys,
  compareLists,
  isUserList,
  listKeyFromUrl,
  listUpdateProps,
  loadSourceSite,
  toSiteRelativeUrl,
  type IListInfoLike
} from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import type { IInstallContext, IList } from '../../src/core/model';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { validateTemplate } from '../../src/core/schema';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { SOURCE_WEB, documents, lookupForras, sourceLists, systemLists, tesztLista } from '../fixtures/lists';

const TARGET_WEB = '/sites/Cel';

function sourceSp(): ReturnType<typeof createMockSp> {
  return createMockSp((req) => {
    if (req.method !== 'GET') return undefined;
    if (/\/_api\/web\/lists\?/i.test(req.url)) return { body: sourceLists };
    if (/\/_api\/web\?\$select=/i.test(req.url)) return { body: SOURCE_WEB };
    return undefined;
  }, SOURCE_WEB.Url);
}

/** In-memory target web; lists are keyed by server-relative URL, created the way SharePoint does (spike 03). */
function targetSite(initial: IListInfoLike[] = []): { sp: ReturnType<typeof createMockSp>['sp']; lists: IListInfoLike[]; requests: IMockRequest[] } {
  const lists = initial.map((l) => ({ ...l, RootFolder: { ServerRelativeUrl: l.RootFolder.ServerRelativeUrl.replace('/sites/Forras', TARGET_WEB) } }));
  let seq = 0;
  const { sp, requests } = createMockSp((req) => {
    const byUrl = /\/getList\('([^']+)'\)/i.exec(req.url);
    const find = (url: string): IListInfoLike | undefined => lists.find((l) => l.RootFolder.ServerRelativeUrl.toLowerCase() === url.toLowerCase());
    if (req.method === 'GET' && byUrl) {
      const l = find(byUrl[1]);
      return l ? { body: l } : { status: 404, body: { 'odata.error': { message: { value: 'A fájl nem található.' } } } };
    }
    if (req.method === 'POST' && byUrl) {
      Object.assign(find(byUrl[1])!, req.body);
      return { status: 204 };
    }
    const byTitle = /\/lists\?\$filter=Title eq '([^']+)'/i.exec(req.url);
    if (req.method === 'GET' && byTitle) {
      return { body: lists.filter((l) => l.Title === byTitle[1]).map((l) => ({ Id: l.Id })) };
    }
    if (req.method === 'POST' && /\/_api\/web\/lists$/i.test(req.url)) {
      const b = req.body as { Title: string; BaseTemplate: number };
      const url = b.BaseTemplate === 101 ? `${TARGET_WEB}/${b.Title}` : `${TARGET_WEB}/Lists/${b.Title}`;
      const created: IListInfoLike = { Id: `33333333-0000-4000-8000-${String(++seq).padStart(12, '0')}`, Title: b.Title, BaseTemplate: b.BaseTemplate, RootFolder: { ServerRelativeUrl: url } };
      lists.push(created);
      return { body: created };
    }
    return undefined;
  });
  return { sp, lists, requests };
}

function ctx(): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: `https://fabrikam.sharepoint.com${TARGET_WEB}`, serverRelativeUrl: TARGET_WEB, title: 'Cél' });
  return { targetSiteUrl: `https://fabrikam.sharepoint.com${TARGET_WEB}`, tokens, log: new Logger() };
}

async function extractAll(): Promise<{ lists: IList[]; writer: JsonTemplateWriter; log: Logger }> {
  const extractor = new ListExtractor();
  const { sp } = sourceSp();
  const writer = new JsonTemplateWriter(
    createEmptyTemplate({
      name: 'Teszt',
      createdBy: 'anna@contoso.com',
      createdAt: '2026-09-24T10:00:00Z',
      sourceSiteUrl: SOURCE_WEB.Url,
      sourceTenant: 'contoso.onmicrosoft.com',
      sourceLcid: 1038,
      includesContent: false
    })
  );
  const log = new Logger();
  const refs = (await extractor.discover(sp)).map((d) => d.ref);
  const site = await loadSourceSite(sp);
  await extractor.extract(sp, refs, { includeContent: false, includeVersions: false, includeMembers: false, tokens: site.tokens, log }, writer);
  expect(validateTemplate(writer.manifest).errors).toEqual([]);
  return { lists: writer.manifest.lists, writer, log };
}

describe('list model', () => {
  it('keeps only user lists (spike 03)', () => {
    expect(sourceLists.filter(isUserList).map((l) => l.Title)).toEqual(['Dokumentumok', 'Események', 'Teszt lista', 'Teszt lookup forrás']);
    expect(systemLists.filter(isUserList)).toEqual([]);
  });

  it('derives site-relative URLs and keys, unique per site', () => {
    expect(toSiteRelativeUrl('/sites/Forras/Lists/Teszt lista', '/sites/Forras/')).toBe('Lists/Teszt lista');
    expect(toSiteRelativeUrl('/Shared Documents', '/')).toBe('Shared Documents');
    expect(listKeyFromUrl('Shared Documents')).toBe('Shared_Documents');
    expect(listKeyFromUrl('Lists/Ügyfelek')).toBe('Ugyfelek');
    expect(listKeyFromUrl('Lists/Árvíztűrő tükörfúrógép')).toBe('Arvizturo_tukorfurogep');
    expect(assignListKeys(['Lists/Projektek', 'Projektek', 'Lists/Teszt lista'])).toEqual({
      'Lists/Projektek': 'Projektek',
      'Lists/Teszt lista': 'Teszt_lista',
      Projektek: 'Projektek_2'
    });
  });

  it('compares only settings present in the template and never sends MajorVersionLimit without versioning', () => {
    const def: IList = { key: 'A', url: 'Lists/A', title: 'A', template: 100, content: { mode: 'none' }, enableVersioning: false, majorVersionLimit: 50 };
    expect(compareLists(def, { Id: 'x', Title: 'A', BaseTemplate: 100, EnableVersioning: false, RootFolder: { ServerRelativeUrl: '/A' } })).toEqual(['majorVersionLimit']);
    expect(listUpdateProps(def)).toEqual({ Title: 'A', EnableVersioning: false });
  });
});

describe('source token context', () => {
  it('registers {listkey} and {listurl} for every user list', async () => {
    const site = await loadSourceSite(sourceSp().sp);
    expect(site.lists.map((l) => l.key)).toEqual(['Shared_Documents', 'Events', 'Teszt_lista', 'Teszt_lookup_forrs']);
    expect(site.tokens.get('listkey', 'Teszt_lookup_forrs')).toBe(lookupForras.Id);
    expect(site.tokens.get('listurl', 'Teszt_lista')).toBe('Lists/Teszt lista');
    expect(site.tokens.get('siterelative')).toBe('/sites/Forras');
  });
});

describe('ListExtractor', () => {
  it('discovers user lists and flags unsupported templates', async () => {
    const found = await new ListExtractor().discover(sourceSp().sp);
    expect(found.map((f) => [f.ref.key, f.unsupported])).toEqual([
      ['list:Shared_Documents', undefined],
      ['list:Events', 'LIST_TEMPLATE_UNSUPPORTED'],
      ['list:Teszt_lista', undefined],
      ['list:Teszt_lookup_forrs', undefined]
    ]);
  });

  it('extracts settings per list type and skips unsupported templates with a warning', async () => {
    const { lists, writer, log } = await extractAll();
    expect(lists.map((l) => l.key)).toEqual(['Shared_Documents', 'Teszt_lista', 'Teszt_lookup_forrs']);
    expect(lists[0]).toMatchObject({ url: 'Shared Documents', template: 101, enableVersioning: true, majorVersionLimit: 500, forceCheckout: false });
    expect(lists[0].enableAttachments).toBeUndefined();
    expect(lists[1]).toMatchObject({ url: 'Lists/Teszt lista', title: 'Teszt lista', description: 'Próba', enableAttachments: true, content: { mode: 'none' } });
    expect(lists[1].forceCheckout).toBeUndefined();
    expect(lists[2]).toMatchObject({ url: 'Lists/Teszt lookup forrs', title: 'Teszt lookup forrás' });
    expect(writer.manifest.meta.warnings).toEqual([expect.objectContaining({ code: 'LIST_TEMPLATE_UNSUPPORTED', artifact: 'list:Events' })]);
    expect(log.counts.warn).toBe(1);
  });
});

describe('ListProvider', () => {
  const provider = new ListProvider();

  it('creates a list at its URL, renames it to the real title and registers tokens; a rerun skips it', async () => {
    const { lists } = await extractAll();
    const target = targetSite();
    const c = ctx();
    const lookup = lists[2];

    expect((await provider.diff(target.sp, lookup, c)).status).toBe('new');
    const result = await provider.apply(target.sp, lookup, 'skip', c);
    expect(result).toMatchObject({ outcome: 'created' });
    expect(target.lists[0]).toMatchObject({ Title: 'Teszt lookup forrás', RootFolder: { ServerRelativeUrl: '/sites/Cel/Lists/Teszt lookup forrs' }, EnableAttachments: true });
    // Created with the URL leaf as title first (spike 03).
    expect(target.requests.find((r) => r.method === 'POST' && /\/lists$/i.test(r.url))!.body).toMatchObject({ Title: 'Teszt lookup forrs', BaseTemplate: 100 });
    expect(c.tokens.get('listkey', 'Teszt_lookup_forrs')).toBe(target.lists[0].Id);
    expect(c.tokens.get('listurl', 'Teszt_lookup_forrs')).toBe('Lists/Teszt lookup forrs');

    const posts = target.requests.filter((r) => r.method === 'POST').length;
    expect((await provider.diff(target.sp, lookup, c)).status).toBe('same');
    expect(await provider.apply(target.sp, lookup, 'update', c)).toMatchObject({ outcome: 'skipped' });
    expect(target.requests.filter((r) => r.method === 'POST').length).toBe(posts);
  });

  it('creates a library at the web root', async () => {
    const { lists } = await extractAll();
    const target = targetSite();
    const docs = { ...lists[0], key: 'Iratok', url: 'Iratok', title: 'Iratok tár' };
    expect(await provider.apply(target.sp, docs, 'skip', ctx())).toMatchObject({ outcome: 'created' });
    expect(target.lists[0]).toMatchObject({ Title: 'Iratok tár', RootFolder: { ServerRelativeUrl: '/sites/Cel/Iratok' }, MajorVersionLimit: 500 });
  });

  it('updates only differing settings of an existing list in update mode', async () => {
    const { lists } = await extractAll();
    const target = targetSite([{ ...documents, EnableVersioning: false, MajorVersionLimit: 0, OnQuickLaunch: true }]);
    const c = ctx();
    expect(await provider.diff(target.sp, lists[0], c)).toMatchObject({ status: 'different', changes: ['enableVersioning', 'majorVersionLimit'] });
    expect(await provider.apply(target.sp, lists[0], 'skip', c)).toMatchObject({ outcome: 'skipped' });
    expect(target.lists[0].EnableVersioning).toBe(false);
    expect(await provider.apply(target.sp, lists[0], 'update', c)).toMatchObject({ outcome: 'updated' });
    expect(target.requests.filter((r) => r.method === 'POST').pop()!.body).toEqual({ EnableVersioning: true, MajorVersionLimit: 500 });
    expect(c.tokens.get('listkey', 'Shared_Documents')).toBe(documents.Id);
  });

  it('rename mode creates a separate copy and points the key tokens at it', async () => {
    const { lists } = await extractAll();
    const target = targetSite([{ ...tesztLista, Title: 'Teszt lista', EnableAttachments: false }]);
    const c = ctx();
    expect(await provider.apply(target.sp, lists[1], 'rename', c)).toMatchObject({ outcome: 'created' });
    const copy = target.lists[1];
    expect(copy).toMatchObject({ Title: 'Teszt lista (copy)', RootFolder: { ServerRelativeUrl: '/sites/Cel/Lists/Teszt lista_copy' } });
    expect(c.tokens.get('listkey', 'Teszt_lista')).toBe(copy.Id);
    expect(target.lists[0].EnableAttachments).toBe(false);
  });

  it('refuses a title taken by another list, a different template and a non-creatable URL', async () => {
    const { lists } = await extractAll();
    const c = ctx();
    const taken = targetSite([{ ...lookupForras, Id: 'other', RootFolder: { ServerRelativeUrl: '/sites/Forras/Lists/Masik' } }]);
    expect(await provider.diff(taken.sp, lists[2], c)).toMatchObject({ status: 'unsupported', changes: ['titleConflict'] });

    const wrongType = targetSite([{ ...tesztLista, BaseTemplate: 101 }]);
    expect(await provider.diff(wrongType.sp, lists[1], c)).toMatchObject({ status: 'unsupported', changes: ['template'] });

    const nested = { ...lists[1], url: 'Lists/Archiv/Teszt' };
    expect(await provider.diff(targetSite().sp, nested, c)).toMatchObject({ status: 'unsupported', changes: ['url'] });
    expect(await provider.apply(targetSite().sp, nested, 'update', c)).toMatchObject({ outcome: 'skipped' });
    expect(c.tokens.get('listkey', 'Teszt_lista')).toBeUndefined();
  });

  it('fails with LIST_URL_MISMATCH when SharePoint puts the list elsewhere', async () => {
    const { lists } = await extractAll();
    // Simulate a site where the new list lands on another URL.
    const misplaced = createMockSp((req) => {
      if (req.method === 'POST' && /\/_api\/web\/lists$/i.test(req.url)) return { body: { Id: 'x' } };
      if (/\/getList\(/i.test(req.url)) return { status: 404, body: {} };
      if (/\/lists\?\$filter=Title/i.test(req.url)) return { body: [] };
      return undefined;
    });
    await expect(provider.apply(misplaced.sp, lists[1], 'skip', ctx())).rejects.toMatchObject({ code: 'LIST_URL_MISMATCH' });
  });
});
