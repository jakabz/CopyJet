import { ViewExtractor } from '../../src/core/extractors/ViewExtractor';
import { ViewProvider } from '../../src/core/providers/ViewProvider';
import { compareViews, listViewDefs, loadSourceSite, matchView, scopeName, toListDef, type IListViewDef, type IViewInfoLike } from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import type { IInstallContext } from '../../src/core/model';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { validateTemplate } from '../../src/core/schema';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { SOURCE_WEB, documents, sourceLists, tesztLista } from '../fixtures/lists';
import { allDocs, allItems, nyitott, racs, ugyfel, viewFields } from '../fixtures/views';

const TARGET_WEB = '/sites/Cel';

interface IViewState {
  views: { [listUrl: string]: IViewInfoLike[] };
  fields: { [viewId: string]: string[] };
}

/** In-memory views of a web covering the REST calls CopyJet makes (spike 06 behaviour). */
function fakeViews(state: IViewState, web = SOURCE_WEB): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  let seq = 0;
  return createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (req.method === 'GET' && /\/_api\/web\?\$select=/i.test(req.url)) return { body: web };
    if (req.method === 'GET' && /\/_api\/web\/lists\?\$select=/i.test(req.url)) return { body: sourceLists };
    if ((m = /\/views\('([^']+)'\)\/viewfields\/removeallviewfields$/i.exec(req.url))) {
      state.fields[m[1]] = [];
      return { body: {} };
    }
    if ((m = /\/views\('([^']+)'\)\/viewfields\/addviewfield\('([^']+)'\)$/i.exec(req.url))) {
      state.fields[m[1]].push(m[2]);
      return { body: {} };
    }
    if (req.method === 'GET' && (m = /\/views\('([^']+)'\)\/viewfields$/i.exec(req.url))) {
      return { body: { Items: state.fields[m[1]] || [] } };
    }
    if (req.method === 'POST' && (m = /\/getList\('([^']+)'\)\/views\('([^']+)'\)$/i.exec(req.url))) {
      const list = state.views[m[1]];
      const view = list.find((x) => x.Id === m![2])!;
      const body = req.body as Partial<IViewInfoLike>;
      if (body.DefaultView) list.forEach((x) => (x.DefaultView = false));
      Object.assign(view, body);
      return { status: 204 };
    }
    if (req.method === 'POST' && (m = /\/getList\('([^']+)'\)\/views$/i.exec(req.url))) {
      const body = req.body as Partial<IViewInfoLike> & { Title: string };
      const created: IViewInfoLike = { Id: `99999999-0000-4000-8000-${String(++seq).padStart(12, '0')}`, ViewType: 'HTML', DefaultView: false, Scope: 0, RowLimit: 30, Paged: false, ...body };
      state.views[m[1]].push(created);
      state.fields[created.Id] = ['LinkTitle'];
      return { body: created };
    }
    if (req.method === 'GET' && (m = /\/getList\('([^']+)'\)\/views\?\$filter=Hidden eq false and PersonalView eq false/i.exec(req.url))) {
      const list = state.views[m[1]];
      return list ? { body: list.filter((x) => !x.Hidden && !x.PersonalView) } : { status: 404, body: {} };
    }
    return undefined;
  }, web.Url);
}

function sourceSp(): ReturnType<typeof fakeViews> {
  const personal: IViewInfoLike = { ...ugyfel, Id: 'p', Title: 'Személyes nézet', PersonalView: true };
  return fakeViews({
    views: {
      [tesztLista.RootFolder.ServerRelativeUrl]: [allItems, nyitott, ugyfel, racs, personal],
      [documents.RootFolder.ServerRelativeUrl]: [allDocs, { ...allDocs, Id: 'h', Title: 'assetLibTemp', DefaultView: false, Hidden: true }]
    },
    fields: viewFields
  });
}

async function extract(): Promise<JsonTemplateWriter> {
  const { sp } = sourceSp();
  const site = await loadSourceSite(sp);
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
  writer.manifest.lists.push(toListDef(tesztLista, 'Teszt_lista', SOURCE_WEB.ServerRelativeUrl), toListDef(documents, 'Shared_Documents', SOURCE_WEB.ServerRelativeUrl));
  const extractor = new ViewExtractor();
  const refs = (await extractor.discover(sp)).map((d) => d.ref);
  await extractor.extract(sp, refs, { includeContent: false, includeVersions: false, includeMembers: false, tokens: site.tokens, log: new Logger() }, writer);
  expect(validateTemplate(writer.manifest).errors).toEqual([]);
  return writer;
}

/** Target with a fresh list: only its language-specific default view "All Items". */
function targetSite(): { state: IViewState } & ReturnType<typeof fakeViews> {
  const listUrl = `${TARGET_WEB}/Lists/Teszt lista`;
  const state: IViewState = {
    views: { [listUrl]: [{ Id: 'aaaaaaaa-0000-4000-8000-000000000001', Title: 'All Items', DefaultView: true, ViewType: 'HTML', RowLimit: 30, Paged: true, Scope: 0, ViewQuery: '' }] },
    fields: { 'aaaaaaaa-0000-4000-8000-000000000001': ['LinkTitle'] }
  };
  return { state, ...fakeViews(state, { Url: `https://fabrikam.sharepoint.com${TARGET_WEB}`, ServerRelativeUrl: TARGET_WEB, Title: 'Cél' }) };
}

function ctx(): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: `https://fabrikam.sharepoint.com${TARGET_WEB}`, serverRelativeUrl: TARGET_WEB, title: 'Cél' }).set(
    'listurl',
    'Teszt_lista',
    'Lists/Teszt lista'
  );
  return { targetSiteUrl: `https://fabrikam.sharepoint.com${TARGET_WEB}`, tokens, log: new Logger() };
}

const defOf = (writer: JsonTemplateWriter, title: string): IListViewDef => listViewDefs(writer.manifest).filter((d) => d.view.title === title)[0];

describe('view model', () => {
  it('matches the default view by flag and others by title', () => {
    const targets: IViewInfoLike[] = [{ Id: '1', Title: 'All Items', DefaultView: true }, { Id: '2', Title: 'Nyitott elemek' }];
    expect(matchView({ title: 'Minden elem', default: true, fields: [] }, targets)!.Id).toBe('1');
    expect(matchView({ title: 'Nyitott elemek', fields: [] }, targets)!.Id).toBe('2');
    expect(matchView({ title: 'Ügyfél nézet', fields: [] }, targets)).toBeUndefined();
  });

  it('compares settings and field order', () => {
    expect(scopeName(2)).toBe('RecursiveAll');
    const view = { title: 'A', fields: ['LinkTitle', 'X'], rowLimit: 30, paged: true, scope: 'Default' as const };
    expect(compareViews(view, { Id: '1', Title: 'A', RowLimit: 30, Paged: true, Scope: 0 }, ['LinkTitle', 'X'])).toEqual([]);
    expect(compareViews(view, { Id: '1', Title: 'A', RowLimit: 30, Paged: true, Scope: 0 }, ['X', 'LinkTitle'])).toEqual(['fields']);
  });
});

describe('ViewExtractor', () => {
  it('discovers public, visible views and flags non-HTML ones', async () => {
    const found = await new ViewExtractor().discover(sourceSp().sp);
    expect(found.map((f) => [f.ref.key, f.unsupported])).toEqual([
      ['view:Shared_Documents/Minden dokumentum', undefined],
      ['view:Teszt_lista/Minden elem', undefined],
      ['view:Teszt_lista/Nyitott elemek', undefined],
      ['view:Teszt_lista/Ügyfél nézet', undefined],
      ['view:Teszt_lista/Rács', 'VIEW_TYPE_UNSUPPORTED']
    ]);
  });

  it('writes views with their fields into the list entries', async () => {
    const writer = await extract();
    const views = writer.manifest.lists[0].views!;
    expect(views.map((v) => v.title)).toEqual(['Minden elem', 'Nyitott elemek', 'Ügyfél nézet', 'Rács']);
    expect(views[0]).toMatchObject({ default: true, fields: viewFields[allItems.Id], rowLimit: 30, paged: true, scope: 'Default', viewType: 'HTML' });
    expect(views[1].query).toBe(nyitott.ViewQuery);
    expect(views[2]).toMatchObject({ scope: 'Recursive', rowLimit: 100 });
    expect(writer.manifest.lists[1].views!.map((v) => v.title)).toEqual(['Minden dokumentum']);
  });

  it('depends on its list and the columns it shows', async () => {
    const writer = await extract();
    expect(new ViewExtractor().dependencies(defOf(writer, 'Nyitott elemek'))).toEqual([
      { kind: 'list', key: 'list:Teszt_lista' },
      { kind: 'listField', key: 'listField:Teszt_lista/LinkTitle' },
      { kind: 'listField', key: 'listField:Teszt_lista/ListaSzoveg' }
    ]);
  });
});

describe('ViewProvider', () => {
  const provider = new ViewProvider();
  const LIST = `${TARGET_WEB}/Lists/Teszt lista`;

  it('updates the language-specific default view in place and creates the others by title', async () => {
    const writer = await extract();
    const target = targetSite();
    const c = ctx();

    const def = defOf(writer, 'Minden elem');
    expect(await provider.diff(target.sp, def, c)).toMatchObject({ status: 'different', changes: ['title', 'fields'] });
    expect(await provider.apply(target.sp, def, 'skip', c)).toMatchObject({ outcome: 'skipped' });
    expect(await provider.apply(target.sp, def, 'update', c)).toMatchObject({ outcome: 'updated' });
    const allItemsTarget = target.state.views[LIST][0];
    expect(allItemsTarget).toMatchObject({ Title: 'Minden elem', DefaultView: true });
    expect(target.state.fields[allItemsTarget.Id]).toEqual(viewFields[allItems.Id]);

    expect(await provider.apply(target.sp, defOf(writer, 'Nyitott elemek'), 'skip', c)).toMatchObject({ outcome: 'created' });
    expect(await provider.apply(target.sp, defOf(writer, 'Ügyfél nézet'), 'skip', c)).toMatchObject({ outcome: 'created' });
    const [, open, client] = target.state.views[LIST];
    expect(open).toMatchObject({ Title: 'Nyitott elemek', ViewQuery: nyitott.ViewQuery, RowLimit: 30, Paged: true, DefaultView: false });
    expect(target.state.fields[open.Id]).toEqual(['LinkTitle', 'ListaSzoveg']);
    expect(client).toMatchObject({ Title: 'Ügyfél nézet', Scope: 1, RowLimit: 100 });
    // Scope was only sent where it is not the default.
    const creates = target.requests.filter((r) => r.method === 'POST' && /\/views\('[^']+'\)$/.test(r.url)).map((r) => r.body);
    expect(creates).toContainEqual({ Scope: 1 });

    const posts = target.requests.filter((r) => r.method === 'POST').length;
    for (const title of ['Minden elem', 'Nyitott elemek', 'Ügyfél nézet']) {
      expect((await provider.diff(target.sp, defOf(writer, title), c)).status).toBe('same');
    }
    expect(target.requests.filter((r) => r.method === 'POST').length).toBe(posts);
  });

  it('skips GRID views and reports a missing list', async () => {
    const writer = await extract();
    const c = ctx();
    expect(await provider.diff(targetSite().sp, defOf(writer, 'Rács'), c)).toMatchObject({ status: 'unsupported', changes: ['viewType'] });
    expect(await provider.apply(targetSite().sp, defOf(writer, 'Rács'), 'update', c)).toMatchObject({ outcome: 'skipped' });
    const missing = { ...defOf(writer, 'Nyitott elemek'), listKey: 'Nincs', listUrl: 'Lists/Nincs' };
    expect(await provider.diff(targetSite().sp, missing, c)).toMatchObject({ status: 'unsupported', changes: ['listMissing'] });
  });

  it('makes a template default view the default when it is created', async () => {
    const writer = await extract();
    const target = targetSite();
    target.state.views[LIST][0].DefaultView = false; // target has no default (edge case)
    const def = defOf(writer, 'Minden elem');
    expect(await provider.apply(target.sp, def, 'skip', ctx())).toMatchObject({ outcome: 'created' });
    expect(target.state.views[LIST].filter((v) => v.DefaultView).map((v) => v.Title)).toEqual(['Minden elem']);
  });
});
