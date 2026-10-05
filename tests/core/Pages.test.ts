import { PageExtractor } from '../../src/core/extractors/PageExtractor';
import { toListDef } from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import { PrincipalMapper } from '../../src/core/mapping';
import type { IInstallContext, IPageFile, ITemplateReader } from '../../src/core/model';
import { ZipTemplateWriter, createEmptyTemplate, openTemplate } from '../../src/core/packager';
import { siteAssetsIn, webPartsOf } from '../../src/core/pages';
import { buildPlan } from '../../src/core/planner';
import { PageProvider } from '../../src/core/providers/PageProvider';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { SOURCE_WEB, lookupForras, sourceLists, tesztLista } from '../fixtures/lists';

const WEB_ID = '8b98bd9e-6008-44fa-a2e2-784ff594eab3';
const SITE_ID = 'aaaaaaaa-1111-4222-8333-000000000001';
const VIEW_ID = '77777777-0000-4000-8000-000000000001';
const LIST_WP = 'f92bf067-bc19-489e-a556-7fe95f508720';
const CUSTOM_WP = 'dddddddd-0000-4000-8000-00000000000c';

// A canvas like spike 12 A's: a list web part on the template's list and view, a link and an image on the
// site, the web and site IDs (upper case, braces), and a list outside the template.
const canvas = [
  { controlType: 4, position: { zoneIndex: 1, sectionIndex: 1, controlIndex: 1 }, innerHTML: '<a href="/sites/Forras/Lists/Teszt%20lista/AllItems.aspx">Lista</a>' },
  {
    controlType: 3,
    webPartId: LIST_WP,
    webPartData: {
      id: LIST_WP,
      title: 'Lista',
      properties: { selectedListId: tesztLista.Id.toUpperCase(), selectedViewId: `{${VIEW_ID}}`, webId: WEB_ID, siteId: SITE_ID, other: lookupForras.Id },
      serverProcessedContent: { imageSources: { img: '/sites/Forras/SiteAssets/SitePages/CopyJet-teszt/k%C3%A9p%201.png' } }
    }
  },
  { controlType: 3, webPartId: CUSTOM_WP, webPartData: { id: CUSTOM_WP, title: 'Egyedi', properties: {} } }
];

function sourceSp(): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  return createMockSp((req) => {
    if (/\/_api\/web\?\$select=Id,Url,ServerRelativeUrl,Title$/i.test(req.url)) return { body: { ...SOURCE_WEB, Id: WEB_ID } };
    if (/\/_api\/web\?\$select=/i.test(req.url)) return { body: SOURCE_WEB };
    if (/\/_api\/site\?\$select=Id$/i.test(req.url)) return { body: { Id: SITE_ID } };
    if (/\/_api\/web\/rootFolder\?\$select=WelcomePage$/i.test(req.url)) return { body: { WelcomePage: 'SitePages/CopyJet-teszt.aspx' } };
    if (/\/_api\/web\/lists\?\$select=/i.test(req.url)) return { body: sourceLists };
    if (/\/views\?\$select=Id,Title$/i.test(req.url)) return { body: [{ Id: VIEW_ID, Title: 'Minden elem' }] };
    if (/\/sitepages\/pages\?\$select=/i.test(req.url)) return { body: { value: [{ Id: 4, FileName: 'CopyJet-teszt.aspx', Title: 'CopyJet teszt', PageLayoutType: 'Article', PromotedState: 0 }] } };
    if (/\/sitepages\/pages\(4\)\?\$select=/i.test(req.url)) {
      return {
        body: {
          Id: 4,
          FileName: 'CopyJet-teszt.aspx',
          Title: 'CopyJet teszt',
          PageLayoutType: 'Article',
          PromotedState: 0,
          CanvasContent1: JSON.stringify(canvas),
          LayoutWebpartsContent: '[{"id":"cbe7b0a9","properties":{"title":"CopyJet teszt"}}]',
          BannerImageUrl: 'https://contoso.sharepoint.com/sites/Forras/SiteAssets/SitePages/CopyJet-teszt/fejlec.jpeg',
          Description: 'Teszt lap'
        }
      };
    }
    if (/\/web\/GetClientSideWebParts$/i.test(req.url)) {
      return { body: { value: [{ Id: LIST_WP, Manifest: '{"isInternal":true}' }, { Id: CUSTOM_WP, Manifest: '{"isInternal":false}' }] } };
    }
    const file = /\/getFileByServerRelativePath\(decodedUrl='([^']+)'\)\/\$value$/i.exec(req.url);
    if (file) return { body: `bytes of ${file[1].split('/').pop()}` };
    return undefined;
  }, SOURCE_WEB.Url);
}

async function extract(log = new Logger()): Promise<ITemplateReader> {
  const { sp } = sourceSp();
  const writer = new ZipTemplateWriter(createEmptyTemplate({ name: 'Lap', createdBy: 'a', createdAt: '2026-09-30T10:00:00Z', sourceSiteUrl: SOURCE_WEB.Url, sourceTenant: 'contoso', sourceLcid: 1038, includesContent: false }));
  const teszt = toListDef(tesztLista, 'Teszt_lista', SOURCE_WEB.ServerRelativeUrl);
  teszt.views = [{ title: 'Minden elem', fields: ['LinkTitle'] }];
  writer.manifest.lists.push(teszt);
  await new PageExtractor().extract(sp, [{ kind: 'page', key: 'page:CopyJet-teszt.aspx' }], { includeContent: true, includeVersions: false, includeMembers: false, tokens: new TokenContext(), log }, writer);
  return openTemplate(await writer.finalize());
}

describe('page model', () => {
  it('finds SiteAssets paths in any URL form, decoded and unique', () => {
    expect(
      siteAssetsIn(['"/sites/Forras/SiteAssets/a%20b/k%C3%A9p.png" and https://x/sites/Forras/SiteAssets/a%20b/kép.png', null as unknown as string, '/sites/Forras/SiteAssets/c.jpg?w=2'], '/sites/Forras')
    ).toEqual(['a b/kép.png', 'c.jpg']);
    expect(webPartsOf(canvas)).toEqual([
      { id: LIST_WP, title: 'Lista' },
      { id: CUSTOM_WP, title: 'Egyedi' }
    ]);
  });
});

describe('PageExtractor', () => {
  it('tokenizes the canvas for the template\'s lists and views, packs SiteAssets images and lists the web parts', async () => {
    const log = new Logger();
    const reader = await extract(log);
    const page = reader.manifest.pages[0];
    expect(page).toEqual({
      name: 'CopyJet-teszt.aspx',
      title: 'CopyJet teszt',
      layout: 'Article',
      source: 'pages/CopyJet-teszt.json',
      promotedState: 0,
      isHomePage: true,
      assets: ['assets/SiteAssets/SitePages/CopyJet-teszt/kép 1.png', 'assets/SiteAssets/SitePages/CopyJet-teszt/fejlec.jpeg'],
      requiredWebParts: [
        { id: LIST_WP, title: 'Lista', isCustom: false },
        { id: CUSTOM_WP, title: 'Egyedi', isCustom: true }
      ]
    });
    const file = await reader.getJson<IPageFile>('pages/CopyJet-teszt.json');
    const text = JSON.stringify(file.canvasContent);
    expect(text).toContain('{siterelative}/Lists/Teszt%20lista/AllItems.aspx');
    expect(text).toContain('"selectedListId":"{listkey:Teszt_lista}"');
    expect(text).toContain('"selectedViewId":"{{viewid:Teszt_lista/Minden elem}}"');
    expect(text).toContain('"webId":"{webid}"');
    expect(text).toContain('"siteId":"{siteid}"');
    // A list outside the template keeps its source ID, with a warning.
    expect(text).toContain(lookupForras.Id);
    expect(log.entries.filter((e) => e.level === 'warn').map((e) => e.code)).toEqual(['PAGE_LIST_OUTSIDE_TEMPLATE']);
    expect(file.bannerImage).toBe('assets/SiteAssets/SitePages/CopyJet-teszt/fejlec.jpeg');
    expect(await (await reader.getBlob('assets/SiteAssets/SitePages/CopyJet-teszt/kép 1.png')).text()).toBe('"bytes of kép 1.png"');
  });

  it('plans pages after their lists and views without blocking on them', async () => {
    const plan = buildPlan((await extract()).manifest);
    const page = plan.steps.filter((s) => s.ref.key === 'page:CopyJet-teszt.aspx')[0];
    expect(page.dependsOn).toEqual([]);
    const listLevel = plan.steps.filter((s) => s.ref.key === 'view:Teszt_lista/Minden elem')[0].level;
    expect(page.level).toBeGreaterThan(listLevel);
  });
});

// ---------------------------------------------------------------------------------------------------------

const TARGET = 'https://fabrikam.sharepoint.com/sites/Cel';
const T_LIST = 'eeeeeeee-0000-4000-8000-000000000001';
const T_VIEW = 'eeeeeeee-0000-4000-8000-000000000002';

function targetSp(existing: string[] = []): { sp: ReturnType<typeof createMockSp>['sp']; calls: string[]; saved: Array<{ [k: string]: unknown }>; files: string[] } {
  const calls: string[] = [];
  const saved: Array<{ [k: string]: unknown }> = [];
  const files: string[] = [];
  const folders: string[] = [];
  let renamedTo = 'Page';
  const { sp } = createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (req.method === 'GET' && /\/sitepages\/pages\?\$select=/i.test(req.url)) return { body: { value: existing.map((f, i) => ({ Id: i + 1, FileName: f })) } };
    if (req.method === 'POST' && /\/_api\/sitepages\/pages$/i.test(req.url)) {
      calls.push(`create ${(req.body as { PageLayoutType: string }).PageLayoutType}`);
      return { body: { Id: 9, FileName: 'Page.aspx' } };
    }
    if (req.method === 'GET' && /\/sitepages\/pages\(9\)\?\$select=/i.test(req.url)) return { body: { Id: 9, FileName: `${renamedTo}.aspx` } };
    if (req.method === 'POST' && /\/getList\('\/sites\/Cel\/SitePages'\)\/items\(9\)\/validateupdatelistitem$/i.test(req.url)) {
      renamedTo = (req.body as { formValues: Array<{ FieldValue: string }> }).formValues[0].FieldValue;
      calls.push(`rename ${renamedTo}`);
      return { body: { value: [{ FieldName: 'FileLeafRef', HasException: false }] } };
    }
    if (req.method === 'POST' && (m = /\/sitepages\/pages\((\d+)\)\/(checkoutpage|savepageasdraft|publish)$/i.exec(req.url))) {
      calls.push(`${m[2]} ${m[1]}`);
      if (m[2] === 'savepageasdraft') saved.push(req.body as { [k: string]: unknown });
      return { body: {} };
    }
    if (req.method === 'POST' && /\/web\/lists\/ensuresiteassetslibrary$/i.test(req.url)) {
      calls.push('ensureSiteAssets');
      return { body: {} };
    }
    if (req.method === 'GET' && (m = /\/getFolderByServerRelativePath\(decodedUrl='([^']+)'\)\/folders\?\$select=Name$/i.exec(req.url))) {
      const rel = m[1].replace('/sites/Cel/SiteAssets', '').replace(/^\//, '');
      return { body: folders.filter((f) => (rel ? f.indexOf(`${rel}/`) === 0 && f.split('/').length === rel.split('/').length + 1 : f.indexOf('/') < 0)).map((f) => ({ Name: f.split('/').pop() })) };
    }
    if (req.method === 'POST' && (m = /\/_api\/web\/folders\/addUsingPath\(DecodedUrl='([^']+)'/i.exec(req.url))) {
      folders.push(m[1].replace('/sites/Cel/SiteAssets/', ''));
      return { body: {} };
    }
    if (req.method === 'POST' && (m = /\/getFolderByServerRelativePath\(decodedUrl='([^']+)'\)\/files\/AddUsingPath\(decodedurl='([^']+)'\)$/i.exec(req.url))) {
      files.push(`${m[1]}/${m[2]}`);
      return { body: {} };
    }
    if (/\/web\/GetClientSideWebParts$/i.test(req.url)) return { body: { value: [{ Id: LIST_WP, Manifest: '{"isInternal":true}' }] } };
    return undefined;
  }, TARGET);
  return { sp, calls, saved, files };
}

function installContext(reader: ITemplateReader, sp: ReturnType<typeof createMockSp>['sp']): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: TARGET, serverRelativeUrl: '/sites/Cel', title: 'Cél' });
  tokens.set('webid', undefined, 'bbbbbbbb-0000-4000-8000-000000000001');
  tokens.set('siteid', undefined, 'bbbbbbbb-0000-4000-8000-000000000002');
  tokens.set('listkey', 'Teszt_lista', T_LIST);
  tokens.set('viewid', 'Teszt_lista/Minden elem', T_VIEW);
  return { targetSiteUrl: TARGET, tokens, log: new Logger(), content: { reader, idMaps: {}, principals: new PrincipalMapper(sp, []) } };
}

describe('PageProvider', () => {
  it('creates, checks out, saves the resolved canvas, renames and publishes – images first', async () => {
    const reader = await extract();
    const target = targetSp();
    const ctx = installContext(reader, target.sp);
    const def = reader.manifest.pages[0];
    expect(await new PageProvider().apply(target.sp, def, 'skip', ctx)).toMatchObject({ outcome: 'created' });
    expect(target.calls).toEqual(['create Article', 'ensureSiteAssets', 'checkoutpage 9', 'savepageasdraft 9', 'rename CopyJet-teszt', 'publish 9']);
    expect(target.files).toEqual(['/sites/Cel/SiteAssets/SitePages/CopyJet-teszt/kép 1.png', '/sites/Cel/SiteAssets/SitePages/CopyJet-teszt/fejlec.jpeg']);
    const body = target.saved[0];
    const text = String(body.CanvasContent1);
    expect(text).toContain('/sites/Cel/Lists/Teszt%20lista/AllItems.aspx');
    expect(text).toContain(`"selectedListId":"${T_LIST}"`);
    expect(text).toContain(`"selectedViewId":"{${T_VIEW}}"`);
    expect(text).toContain('"webId":"bbbbbbbb-0000-4000-8000-000000000001"');
    expect(body.BannerImageUrl).toBe(`${TARGET}/SiteAssets/SitePages/CopyJet-teszt/fejlec.jpeg`);
    expect(body.Title).toBe('CopyJet teszt');
    // The custom web part is not on the target: a warning, the page still installs.
    expect(ctx.log.entries.filter((e) => e.level === 'warn').map((e) => e.code)).toEqual(['PAGE_WEBPART_MISSING']);
  });

  it('keeps an existing page unless the mode is update', async () => {
    const reader = await extract();
    const target = targetSp(['CopyJet-teszt.aspx']);
    const def = reader.manifest.pages[0];
    expect(await new PageProvider().apply(target.sp, def, 'skip', installContext(reader, target.sp))).toMatchObject({ outcome: 'skipped' });
    expect(target.calls).toEqual([]);
    expect(await new PageProvider().apply(target.sp, def, 'update', installContext(reader, target.sp))).toMatchObject({ outcome: 'updated' });
    expect(target.calls).toEqual(['ensureSiteAssets', 'checkoutpage 1', 'savepageasdraft 1', 'publish 1']);
  });
});
