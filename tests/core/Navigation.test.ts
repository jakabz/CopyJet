import { NavigationExtractor } from '../../src/core/extractors/NavigationExtractor';
import { toListDef } from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import type { IInstallContext, INavNode } from '../../src/core/model';
import { comparableUrl, isSystemNode, LINKLESS, resolveNavUrl, tokenizeNavUrl, type INavStepDef, type ITargetNavNode } from '../../src/core/navigation';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { buildPlan } from '../../src/core/planner';
import { NavigationProvider } from '../../src/core/providers/NavigationProvider';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { SOURCE_WEB, tesztLista } from '../fixtures/lists';

const AUDIENCE = 'a64a6140-20c7-44c7-ae78-754d724f63e8';
const SRC = { webUrl: SOURCE_WEB.Url, webServerRelativeUrl: '/sites/Forras', lists: [{ key: 'Teszt_lista', url: 'Lists/Teszt lista' }] };

describe('navigation model', () => {
  it.each([
    ['/sites/Forras', '{siterelative}'],
    ['/sites/Forras/SitePages/CopyJet-teszt.aspx', '{siterelative}/SitePages/CopyJet-teszt.aspx'],
    ['https://contoso.sharepoint.com/sites/Forras/SitePages/A.aspx', '{siterelative}/SitePages/A.aspx'],
    ['/sites/Forras/Lists/Teszt lista/AllItems.aspx', '{siterelative}/{listurl:Teszt_lista}/AllItems.aspx'],
    ['/sites/Forras/Lists/Teszt lista/', '{siterelative}/{listurl:Teszt_lista}/'],
    ['/sites/Forras/Lists/Teszt listak', '{siterelative}/Lists/Teszt listak'],
    ['/sites/Forras2/SitePages/X.aspx', '/sites/Forras2/SitePages/X.aspx'],
    ['https://microsoft.com', 'https://microsoft.com'],
    [LINKLESS, LINKLESS]
  ])('tokenizes %s as %s', (url, expected) => {
    expect(tokenizeNavUrl(url, SRC)).toBe(expected);
  });

  it('resolves for the target and compares URLs in one form', () => {
    const tokens = TokenContext.forSite({ absoluteUrl: 'https://fabrikam.sharepoint.com/sites/Cel', serverRelativeUrl: '/sites/Cel', title: 'Cél' });
    tokens.set('listurl', 'Teszt_lista', 'Lists/Teszt lista 2');
    expect(resolveNavUrl('{siterelative}/{listurl:Teszt_lista}/AllItems.aspx', tokens)).toBe('/sites/Cel/Lists/Teszt lista 2/AllItems.aspx');
    expect(comparableUrl('https://fabrikam.sharepoint.com/sites/Cel/Lists/Teszt%20lista/', 'https://fabrikam.sharepoint.com/sites/Cel')).toBe('/sites/cel/lists/teszt lista');
    expect(isSystemNode('1031')).toBe(true);
    expect(isSystemNode(2009)).toBe(false);
  });
});

// MenuState of the source as in spike 13 A: built-in nodes, an unlinked heading with a list link, an external
// link with an audience, a hidden node and Recent with its auto-filled children.
const menuState = {
  Nodes: [
    { Key: '1031', Title: 'Kezdőlap', SimpleUrl: '/sites/Forras', Nodes: [] },
    { Key: '2001', Title: 'Dokumentumok', SimpleUrl: '/sites/Forras/Shared Documents/Forms/AllItems.aspx', Nodes: [] },
    { Key: '1034', Title: 'Webhely tartalma', SimpleUrl: '/sites/Forras/_layouts/15/viewlsts.aspx', Nodes: [] },
    {
      Key: '2009',
      Title: 'CopyJet teszt',
      SimpleUrl: LINKLESS,
      Nodes: [
        { Key: '2010', Title: 'Teszt lista', SimpleUrl: '/sites/Forras/Lists/Teszt lista/', Nodes: [] },
        { Key: '2012', Title: 'Célközönség teszt', SimpleUrl: 'https://index.hu', AudienceIds: [AUDIENCE.toUpperCase()], Nodes: [] },
        { Key: '2013', Title: 'Rejtett', SimpleUrl: 'https://example.com', IsHidden: true, Nodes: [] }
      ]
    },
    { Key: '1033', Title: 'Legutóbbiak', SimpleUrl: '', Nodes: [{ Key: '2008', Title: 'CopyJetSpike11', SimpleUrl: '/sites/Forras/Lists/CopyJetSpike11/AllItems.aspx' }] }
  ]
};

function sourceSp(): ReturnType<typeof createMockSp>['sp'] {
  return createMockSp((req) => {
    if (/\/_api\/web\?\$select=Url,ServerRelativeUrl$/i.test(req.url)) return { body: SOURCE_WEB };
    if (/menuNodeKey='1025'/.test(req.url)) return { body: menuState };
    if (/menuNodeKey='1002'/.test(req.url)) return { body: { Nodes: [] } };
    if (/\/rootFolder\?\$select=WelcomePage$/i.test(req.url)) return { body: { WelcomePage: 'SitePages/CopyJet-teszt.aspx' } };
    return undefined;
  }, SOURCE_WEB.Url).sp;
}

async function extract(log = new Logger(), withPage = true): Promise<JsonTemplateWriter> {
  const writer = new JsonTemplateWriter(
    createEmptyTemplate({ name: 'Nav', createdBy: 'a', createdAt: '2026-10-05T10:00:00Z', sourceSiteUrl: SOURCE_WEB.Url, sourceTenant: 'contoso.sharepoint.com', sourceLcid: 1038, includesContent: false })
  );
  writer.manifest.lists.push(toListDef(tesztLista, 'Teszt_lista', SOURCE_WEB.ServerRelativeUrl));
  if (withPage) writer.manifest.pages.push({ name: 'CopyJet-teszt.aspx', title: 'CopyJet teszt', layout: 'Article', source: 'pages/CopyJet-teszt.json' });
  const refs = [
    { kind: 'navigation' as const, key: 'navigation:quickLaunch' },
    { kind: 'navigation' as const, key: 'navigation:homePage' }
  ];
  await new NavigationExtractor().extract(sourceSp(), refs, { includeContent: false, includeVersions: false, includeMembers: false, tokens: new TokenContext(), log }, writer);
  return writer;
}

describe('NavigationExtractor', () => {
  it('discovers the menus with their link count and the home page', async () => {
    const found = await new NavigationExtractor().discover(sourceSp());
    expect(found.map((a) => [a.ref.key, a.title, a.itemCount])).toEqual([
      ['navigation:quickLaunch', 'quickLaunch', 4],
      ['navigation:homePage', 'CopyJet-teszt.aspx', undefined]
    ]);
  });

  it('carries the user nodes tokenized, without built-in, Recent and hidden nodes', async () => {
    const log = new Logger();
    const nav = (await extract(log)).manifest.navigation!;
    expect(nav.mode).toBe('append');
    expect(nav.quickLaunch).toEqual([
      { title: 'Dokumentumok', url: '{siterelative}/Shared Documents/Forms/AllItems.aspx' },
      {
        title: 'CopyJet teszt',
        url: LINKLESS,
        isExternal: true,
        children: [
          { title: 'Teszt lista', url: '{siterelative}/{listurl:Teszt_lista}/' },
          { title: 'Célközönség teszt', url: 'https://index.hu', isExternal: true, audiences: [AUDIENCE] }
        ]
      }
    ]);
    expect(nav.homePage).toBe('CopyJet-teszt.aspx');
    expect(log.entries.map((e) => e.code).filter((c) => !!c)).toEqual(['NAV_HIDDEN_SKIPPED']);
  });

  it('warns when the home page is not in the template', async () => {
    const log = new Logger();
    await extract(log, false);
    expect(log.entries.filter((e) => e.level === 'warn').map((e) => e.code)).toEqual(['NAV_HOME_PAGE_NOT_IN_TEMPLATE']);
  });

  it('plans the menu and the home page last, after the pages, without blocking on them', async () => {
    const plan = buildPlan((await extract()).manifest);
    const keys = plan.steps.map((s) => s.ref.key);
    expect(keys.slice(-2).sort()).toEqual(['navigation:homePage', 'navigation:quickLaunch']);
    const nav = plan.steps.filter((s) => s.ref.key === 'navigation:quickLaunch')[0];
    expect(nav.dependsOn).toEqual([]);
    expect(nav.level).toBeGreaterThan(plan.steps.filter((s) => s.ref.key === 'page:CopyJet-teszt.aspx')[0].level);
  });
});

// ---------------------------------------------------------------------------------------------------------

const TARGET = 'https://contoso.sharepoint.com/sites/Cel';

interface ITarget {
  sp: ReturnType<typeof createMockSp>['sp'];
  requests: IMockRequest[];
  menu: ITargetNavNode[];
  welcome: { page: string };
}

/** A target with its own menu; adds create nodes, a link to a missing page fails as in spike 13 B. */
function targetSp(pages: string[] = ['TopicHome.aspx', 'CopyJet-teszt.aspx']): ITarget {
  const menu: ITargetNavNode[] = [
    { Id: 1031, Title: 'Kezdőlap', Url: '/sites/Cel', Children: [] },
    { Id: 2001, Title: 'Dokumentumok', Url: '/sites/Cel/Shared Documents/Forms/AllItems.aspx', Children: [] }
  ];
  const welcome = { page: 'SitePages/TopicHome.aspx' };
  let next = 3000;
  const byId = (id: number, nodes: ITargetNavNode[] = menu): ITargetNavNode | undefined => {
    for (const n of nodes) {
      if (n.Id === id) return n;
      const found = byId(id, n.Children || []);
      if (found) return found;
    }
    return undefined;
  };
  const { sp, requests } = createMockSp((req) => {
    let m: RegExpExecArray | null = null;
    if (req.method === 'GET' && /\/navigation\/QuickLaunch\?\$expand=/i.test(req.url)) return { body: { value: JSON.parse(JSON.stringify(menu)) } };
    if (req.method === 'POST' && (/\/navigation\/QuickLaunch$/i.test(req.url) || (m = /\/navigation\/GetNodeById\((\d+)\)\/Children$/i.exec(req.url)))) {
      const b = req.body as { Title: string; Url: string; IsExternal: boolean };
      if (/Nincs-ilyen/.test(b.Url)) {
        return { status: 500, body: { 'odata.error': { code: '-2130247147, Microsoft.SharePoint.SPException', message: { value: 'nincs ilyen fájl' } } } };
      }
      const node: ITargetNavNode = { Id: next++, Title: b.Title, Url: b.Url, Children: [] };
      const parent = m ? byId(parseInt(m[1], 10)) : undefined;
      (parent ? parent.Children! : menu).push(node);
      return { body: { Id: node.Id, Url: b.Url } };
    }
    if (req.method === 'POST' && /\/navigation\/GetNodeById\(\d+\)$/i.test(req.url)) return { body: {} };
    if (req.method === 'GET' && /\/rootFolder\?\$select=WelcomePage$/i.test(req.url)) return { body: { WelcomePage: welcome.page } };
    if (req.method === 'POST' && /\/web\/rootfolder$/i.test(req.url)) {
      welcome.page = (req.body as { WelcomePage: string }).WelcomePage;
      return { body: {} };
    }
    if (/\/sitepages\/pages\?\$select=/i.test(req.url)) return { body: { value: pages.map((f, i) => ({ Id: i + 1, FileName: f })) } };
    return undefined;
  }, TARGET);
  return { sp, requests, menu, welcome };
}

function installContext(): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: TARGET, serverRelativeUrl: '/sites/Cel', title: 'Cél' });
  tokens.set('listurl', 'Teszt_lista', 'Lists/Teszt lista');
  return { targetSiteUrl: TARGET, tokens, log: new Logger() };
}

const nodes: INavNode[] = [
  { title: 'Dokumentumok', url: '{siterelative}/Shared Documents/Forms/AllItems.aspx' },
  {
    title: 'CopyJet teszt',
    url: LINKLESS,
    isExternal: true,
    children: [
      { title: 'Teszt lista', url: '{siterelative}/{listurl:Teszt_lista}/' },
      { title: 'Célközönség teszt', url: 'https://index.hu', isExternal: true, audiences: [AUDIENCE] },
      { title: 'Törött', url: '{siterelative}/SitePages/Nincs-ilyen.aspx' },
      { title: 'Ismeretlen lista', url: '{siterelative}/{listurl:Mas_lista}/AllItems.aspx' }
    ]
  },
  { title: 'Törött fejléc', url: '{siterelative}/SitePages/Nincs-ilyen.aspx', children: [{ title: 'Külső', url: 'https://microsoft.com', isExternal: true }] }
];

const menuDef = (sourceTenant = 'contoso.sharepoint.com'): INavStepDef => ({ part: 'quickLaunch', nodes, mode: 'append', sourceTenant });

describe('NavigationProvider', () => {
  it('counts the links the target menu lacks', async () => {
    const t = targetSp();
    expect(await new NavigationProvider().diff(t.sp, menuDef(), installContext())).toMatchObject({ status: 'new', changes: ['newLinks:7'] });
  });

  it('appends the missing links, skips broken ones, keeps a broken heading without link, sets audiences', async () => {
    const t = targetSp();
    const ctx = installContext();
    expect(await new NavigationProvider().apply(t.sp, menuDef(), 'skip', ctx)).toMatchObject({ outcome: 'created' });
    const flat = (ns: ITargetNavNode[]): string[] => ns.reduce((a: string[], n) => a.concat(`${n.Title} → ${n.Url}`, flat(n.Children || []).map((c) => `  ${c}`)), []);
    expect(flat(t.menu)).toEqual([
      'Kezdőlap → /sites/Cel',
      'Dokumentumok → /sites/Cel/Shared Documents/Forms/AllItems.aspx',
      `CopyJet teszt → ${LINKLESS}`,
      '  Teszt lista → /sites/Cel/Lists/Teszt lista/',
      '  Célközönség teszt → https://index.hu',
      `Törött fejléc → ${LINKLESS}`,
      '  Külső → https://microsoft.com'
    ]);
    const merges = t.requests.filter((r) => r.method === 'POST' && /GetNodeById\(\d+\)$/i.test(r.url));
    expect(merges.map((r) => r.body)).toEqual([{ AudienceIds: [AUDIENCE] }]);
    const warns = ctx.log.entries.filter((e) => e.level === 'warn');
    expect(warns.map((e) => e.code)).toEqual(['NAV_LINK_BROKEN']);
    expect(warns[0].detail).toEqual(['Törött', 'Ismeretlen lista', 'Törött fejléc']);

    // A rerun finds everything (the broken links stay missing, nothing is duplicated).
    const before = t.menu.length;
    await new NavigationProvider().apply(t.sp, menuDef(), 'skip', installContext());
    expect(t.menu.length).toBe(before);
    expect(t.menu[2].Children!.length).toBe(2);
  });

  it('leaves audiences out in another tenant', async () => {
    const t = targetSp();
    const ctx = installContext();
    await new NavigationProvider().apply(t.sp, menuDef('fabrikam.sharepoint.com'), 'skip', ctx);
    expect(t.requests.some((r) => r.method === 'POST' && /GetNodeById\(\d+\)$/i.test(r.url))).toBe(false);
    expect(ctx.log.entries.map((e) => e.code)).toContain('NAV_AUDIENCE_OTHER_TENANT');
  });

  it('sets the home page when the page is on the target', async () => {
    const t = targetSp();
    const def: INavStepDef = { part: 'homePage', page: 'CopyJet-teszt.aspx', inTemplate: true };
    expect(await new NavigationProvider().diff(t.sp, def, installContext())).toMatchObject({ status: 'new' });
    expect(await new NavigationProvider().apply(t.sp, def, 'skip', installContext())).toMatchObject({ outcome: 'updated' });
    expect(t.welcome.page).toBe('SitePages/CopyJet-teszt.aspx');
    expect(await new NavigationProvider().diff(t.sp, def, installContext())).toMatchObject({ status: 'same' });
  });

  it('keeps the home page when the page is missing', async () => {
    const t = targetSp(['TopicHome.aspx']);
    const def: INavStepDef = { part: 'homePage', page: 'CopyJet-teszt.aspx', inTemplate: false };
    expect(await new NavigationProvider().diff(t.sp, def, installContext())).toMatchObject({ status: 'unsupported', changes: ['homePageMissing'] });
    const ctx = installContext();
    expect(await new NavigationProvider().apply(t.sp, def, 'skip', ctx)).toMatchObject({ outcome: 'skipped' });
    expect(t.welcome.page).toBe('SitePages/TopicHome.aspx');
    expect(ctx.log.entries.map((e) => e.code)).toContain('NAV_HOME_PAGE_MISSING');
  });
});
