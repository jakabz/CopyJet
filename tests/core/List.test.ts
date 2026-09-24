import { ListExtractor } from '../../src/core/extractors/ListExtractor';
import { ListProvider } from '../../src/core/providers/ListProvider';
import {
  assignListKeys,
  compareLists,
  isUserList,
  listKeyFromUrl,
  listUpdateProps,
  loadSourceSite,
  missingFolders,
  toSiteRelativeUrl,
  type IListInfoLike
} from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import type { IInstallContext, IList } from '../../src/core/model';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { validateTemplate } from '../../src/core/schema';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { PROJEKT_SITE_CT, SOURCE_WEB, documents, lookupForras, sourceFolders, sourceListCts, sourceLists, systemLists, tesztLista } from '../fixtures/lists';

const TARGET_WEB = '/sites/Cel';

interface ISiteState {
  web: { Url: string; ServerRelativeUrl: string; Title: string };
  lists: IListInfoLike[];
  /** Folder paths per list URL, relative to the list root. */
  folders: { [listUrl: string]: string[] };
  cts: { [listUrl: string]: { ordered: string[]; all: string[] } };
  /** Site content types that can be added to lists. */
  siteCts: string[];
}

const guid32 = (n: number): string => n.toString(16).toUpperCase().padStart(32, '0');

/**
 * In-memory SharePoint web covering the list REST calls CopyJet makes. Behaviour follows spikes 03/04:
 * missing list = 404, new list gets its URL from the title, nested folders need existing parents.
 */
function fakeSite(state: ISiteState): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[]; fetch: (url: string, init: RequestInit) => Promise<Response> } {
  let seq = 0;
  const findList = (url: string): IListInfoLike | undefined => state.lists.find((l) => l.RootFolder.ServerRelativeUrl.toLowerCase() === url.toLowerCase());
  const listOfPath = (path: string): string | undefined =>
    state.lists.map((l) => l.RootFolder.ServerRelativeUrl).filter((u) => path === u || path.indexOf(u + '/') === 0)[0];

  const { sp, requests } = createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (req.method === 'POST' && /\/_api\/contextinfo$/i.test(req.url)) {
      return { body: { FormDigestValue: 'digest', WebFullUrl: state.web.Url } };
    }
    if (req.method === 'GET' && /\/_api\/web\?\$select=/i.test(req.url)) return { body: state.web };
    if (req.method === 'GET' && /\/_api\/web\/lists\?\$select=/i.test(req.url)) return { body: state.lists };
    if ((m = /\/getFolderByServerRelativePath\(decodedUrl='([^']+)'\)\/folders/i.exec(req.url))) {
      const path = m[1];
      const listUrl = listOfPath(path)!;
      const rel = path === listUrl ? '' : path.slice(listUrl.length + 1);
      const children = (state.folders[listUrl] || []).filter((f) => (rel ? f.indexOf(rel + '/') === 0 && f.split('/').length === rel.split('/').length + 1 : f.indexOf('/') < 0));
      return { body: children.map((f) => ({ Name: f.split('/').pop() })) };
    }
    if (req.method === 'POST' && (m = /\/folders\/addUsingPath\(DecodedUrl='([^']+)'/i.exec(req.url))) {
      const path = m[1];
      const listUrl = listOfPath(path)!;
      const rel = path.slice(listUrl.length + 1);
      const parent = rel.split('/').slice(0, -1).join('/');
      const folders = (state.folders[listUrl] = state.folders[listUrl] || []);
      if (parent && folders.indexOf(parent) < 0) return { status: 500, body: { 'odata.error': { message: { value: `"${path}" nem található.` } } } };
      if (folders.indexOf(rel) < 0) folders.push(rel);
      return { body: { Name: rel.split('/').pop() } };
    }
    if ((m = /\/getList\('([^']+)'\)\/rootFolder\?/i.exec(req.url))) {
      return { body: { ContentTypeOrder: ((state.cts[m[1]] || { ordered: [] }).ordered).map((StringValue) => ({ StringValue })) } };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\/contentTypes\?/i.exec(req.url))) {
      return { body: ((state.cts[m[1]] || { all: [] }).all).map((StringId) => ({ StringId })) };
    }
    if (req.method === 'POST' && (m = /\/getList\('([^']+)'\)\/contentTypes\/addAvailableContentType\('([^']+)'\)$/i.exec(req.url))) {
      if (state.siteCts.indexOf(m[2]) < 0) return { status: 500, body: { 'odata.error': { message: { value: `A tartalomtípus nem található (azonosító: '${m[2]}').` } } } };
      const listCt = `${m[2]}00${guid32(++seq)}`;
      const cts = (state.cts[m[1]] = state.cts[m[1]] || { ordered: [], all: [] });
      cts.all.push(listCt);
      cts.ordered.push(listCt);
      return { body: { StringId: listCt } };
    }
    if ((m = /\/getList\('([^']+)'\)(\?|$)/i.exec(req.url))) {
      const l = findList(m[1]);
      if (req.method === 'POST') {
        Object.assign(l!, req.body);
        return { status: 204 };
      }
      return l ? { body: l } : { status: 404, body: { 'odata.error': { message: { value: 'A fájl nem található.' } } } };
    }
    if (req.method === 'GET' && (m = /\/lists\?\$filter=Title eq '([^']+)'/i.exec(req.url))) {
      return { body: state.lists.filter((l) => l.Title === m![1]).map((l) => ({ Id: l.Id })) };
    }
    if (req.method === 'POST' && /\/_api\/web\/lists$/i.test(req.url)) {
      const b = req.body as { Title: string; BaseTemplate: number };
      const url = b.BaseTemplate === 101 ? `${state.web.ServerRelativeUrl}/${b.Title}` : `${state.web.ServerRelativeUrl}/Lists/${b.Title}`;
      const created: IListInfoLike = { Id: `33333333-0000-4000-8000-${String(++seq).padStart(12, '0')}`, Title: b.Title, BaseTemplate: b.BaseTemplate, RootFolder: { ServerRelativeUrl: url } };
      state.lists.push(created);
      state.folders[url] = [b.BaseTemplate === 101 ? 'Forms' : 'Attachments'];
      const def = `${b.BaseTemplate === 101 ? '0x0101' : '0x01'}00${guid32(++seq)}`;
      state.cts[url] = { ordered: [def], all: [def, `0x012000${guid32(++seq)}`] };
      return { body: created };
    }
    return undefined;
  }, state.web.Url);

  // Raw REST (spike 04): MERGE UniqueContentTypeOrder on a list's root folder, odata=verbose.
  const rawFetch = async (url: string, init: RequestInit): Promise<Response> => {
    const m = /\/getList\('([^']+)'\)\/RootFolder$/i.exec(decodeURIComponent(url));
    expect(m).not.toBeNull();
    expect((init.headers as Record<string, string>)['Content-Type']).toBe('application/json;odata=verbose');
    const ids = (JSON.parse(String(init.body)) as { UniqueContentTypeOrder: { results: Array<{ StringValue: string }> } }).UniqueContentTypeOrder.results.map((r) => r.StringValue);
    state.cts[m![1]].ordered = ids;
    return new Response(null, { status: 204 });
  };
  return { sp, requests, fetch: rawFetch };
}

function sourceSp(): ReturnType<typeof fakeSite> {
  return fakeSite({ web: SOURCE_WEB, lists: sourceLists, folders: sourceFolders, cts: sourceListCts, siteCts: [] });
}

/** Target web; `initial` lists are taken from the source fixtures and moved to the target URL. */
function targetSite(initial: IListInfoLike[] = []): ReturnType<typeof fakeSite> & { lists: IListInfoLike[]; state: ISiteState } {
  const state: ISiteState = {
    web: { Url: `https://fabrikam.sharepoint.com${TARGET_WEB}`, ServerRelativeUrl: TARGET_WEB, Title: 'Cél' },
    lists: initial.map((l) => ({ ...l, RootFolder: { ServerRelativeUrl: l.RootFolder.ServerRelativeUrl.replace('/sites/Forras', TARGET_WEB) } })),
    folders: {},
    cts: {},
    siteCts: ['0x01', '0x0101', PROJEKT_SITE_CT]
  };
  state.lists.forEach((l) => {
    state.folders[l.RootFolder.ServerRelativeUrl] = [];
    const def = `${l.BaseTemplate === 101 ? '0x0101' : '0x01'}00${'D'.repeat(32)}`;
    state.cts[l.RootFolder.ServerRelativeUrl] = { ordered: [def], all: [def] };
  });
  return { ...fakeSite(state), lists: state.lists, state };
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

describe('list structure (content types, folders)', () => {
  it('extracts site-level content type IDs (default first) and user folders only', async () => {
    const { lists } = await extractAll();
    const teszt = lists.find((l) => l.key === 'Teszt_lista')!;
    expect(teszt.contentTypes).toEqual([PROJEKT_SITE_CT, '0x01']);
    expect(teszt.folders).toEqual([{ path: 'Mappa' }]);
    expect(lists.find((l) => l.key === 'Shared_Documents')!.folders).toEqual([{ path: '2026' }, { path: '2026/Q3' }]);
    // Content types are only recorded when the list has them enabled.
    expect(lists.find((l) => l.key === 'Teszt_lookup_forrs')!.contentTypes).toBeUndefined();
  });

  it('computes missing folders parent-first', () => {
    expect(missingFolders(['a/b/c', 'x'], ['a'])).toEqual(['x', 'a/b', 'a/b/c']);
    expect(missingFolders(['A/b'], ['a', 'a/B'])).toEqual([]);
  });

  it('creates a list with its content types in template order (first = default) and its folders', async () => {
    const { lists } = await extractAll();
    const target = targetSite();
    const provider = new ListProvider(target.fetch);
    const c = ctx();
    expect(await provider.apply(target.sp, lists.find((l) => l.key === 'Teszt_lista')!, 'skip', c)).toMatchObject({ outcome: 'created' });

    const url = '/sites/Cel/Lists/Teszt lista';
    const cts = target.state.cts[url];
    expect(cts.ordered.map((id) => id.slice(0, -34))).toEqual([PROJEKT_SITE_CT, '0x01']);
    expect(cts.all).toHaveLength(3); // Item, Folder (hidden), Projekt – nothing removed
    expect(target.state.folders[url]).toEqual(['Attachments', 'Mappa']);
    expect((await provider.diff(target.sp, lists.find((l) => l.key === 'Teszt_lista')!, c)).status).toBe('same');
  });

  it('diff reports a missing content type and a wrong default; update adds and reorders without removing', async () => {
    const { lists } = await extractAll();
    const teszt = lists.find((l) => l.key === 'Teszt_lista')!;
    const target = targetSite([{ ...tesztLista }]);
    const provider = new ListProvider(target.fetch);
    const c = ctx();
    const url = '/sites/Cel/Lists/Teszt lista';
    target.state.cts[url].all.push('0x0104007B48E5823A4D6E4C8C5BF88502FD1CDA');
    target.state.cts[url].ordered.push('0x0104007B48E5823A4D6E4C8C5BF88502FD1CDA');

    expect(await provider.diff(target.sp, teszt, c)).toMatchObject({ status: 'different', changes: ['contentTypes', 'folders'] });
    expect(await provider.apply(target.sp, teszt, 'update', c)).toMatchObject({ outcome: 'updated' });
    expect(target.state.cts[url].ordered.map((id) => id.slice(0, -34))).toEqual([PROJEKT_SITE_CT, '0x01', '0x0104']);

    // Only the default differs now.
    target.state.cts[url].ordered.reverse();
    expect(await provider.diff(target.sp, teszt, c)).toMatchObject({ status: 'different', changes: ['contentTypeOrder'] });
  });

  it('fails with LIST_CT_FAILED when a content type is missing from the target site', async () => {
    const { lists } = await extractAll();
    const target = targetSite();
    target.state.siteCts = ['0x01', '0x0101'];
    await expect(new ListProvider(target.fetch).apply(target.sp, lists.find((l) => l.key === 'Teszt_lista')!, 'skip', ctx())).rejects.toMatchObject({
      code: 'LIST_CT_FAILED'
    });
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
    expect(await provider.diff(target.sp, lists[0], c)).toMatchObject({ status: 'different', changes: ['enableVersioning', 'majorVersionLimit', 'folders'] });
    expect(await provider.apply(target.sp, lists[0], 'skip', c)).toMatchObject({ outcome: 'skipped' });
    expect(target.lists[0].EnableVersioning).toBe(false);
    expect(target.state.folders['/sites/Cel/Shared Documents']).toEqual([]);
    expect(await provider.apply(target.sp, lists[0], 'update', c)).toMatchObject({ outcome: 'updated' });
    expect(target.requests.find((r) => r.method === 'POST' && /\/getList\('[^']+'\)$/.test(r.url))!.body).toEqual({ EnableVersioning: true, MajorVersionLimit: 500 });
    expect(target.state.folders['/sites/Cel/Shared Documents']).toEqual(['2026', '2026/Q3']);
    expect(c.tokens.get('listkey', 'Shared_Documents')).toBe(documents.Id);
  });

  it('rename mode creates a separate copy and points the key tokens at it', async () => {
    const { lists } = await extractAll();
    const target = targetSite([{ ...tesztLista, Title: 'Teszt lista', EnableAttachments: false }]);
    const c = ctx();
    expect(await new ListProvider(target.fetch).apply(target.sp, lists[1], 'rename', c)).toMatchObject({ outcome: 'created' });
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
