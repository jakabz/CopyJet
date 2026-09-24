import { ContentTypeExtractor } from '../../src/core/extractors/ContentTypeExtractor';
import { ContentTypeProvider } from '../../src/core/providers/ContentTypeProvider';
import {
  compareContentTypes,
  isCustomContentType,
  parentIdOf,
  type IContentTypeInfoLike,
  type IFieldLinkInfoLike
} from '../../src/core/contentTypes';
import { Logger } from '../../src/core/logger';
import type { IContentType, IInstallContext } from '../../src/core/model';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { validateTemplate } from '../../src/core/schema';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { KIEMELT_ID, PROJEKT_ID, customPageCt, documentCt, hiddenCt, itemCt, kiemeltCt, link, projektCt, sitePageCt, sourceContentTypes, sourceLinks } from '../fixtures/contentTypes';

const ctId = (url: string, collection: string): string | undefined => {
  const m = new RegExp(`/${collection}\\('([^']+)'\\)`, 'i').exec(url);
  return m ? m[1] : undefined;
};

function sourceSp(): ReturnType<typeof createMockSp> {
  return createMockSp((req) => {
    if (req.method !== 'GET') return undefined;
    if (/\/_api\/web\/contenttypes\?/i.test(req.url)) return { body: sourceContentTypes };
    const id = ctId(req.url, 'availablecontenttypes');
    if (id && /\/fieldlinks/i.test(req.url)) return { body: sourceLinks[id] || [] };
    return undefined;
  }, 'https://contoso.sharepoint.com/sites/forras');
}

interface ITargetState {
  cts: IContentTypeInfoLike[];
  links: { [id: string]: IFieldLinkInfoLike[] };
}

/**
 * In-memory target. `keepIds: false` simulates SharePoint ignoring the requested ID; `failLinks` makes
 * adding field links fail as it does when the site column is missing.
 */
function targetSite(
  initial: Partial<ITargetState> = {},
  keepIds = true,
  failLinks = false
): { sp: ReturnType<typeof createMockSp>['sp']; state: ITargetState; requests: IMockRequest[]; provider: ContentTypeProvider; csomBodies: string[] } {
  const state: ITargetState = {
    cts: (initial.cts || [itemCt, documentCt]).map((c) => ({ ...c })),
    links: { '0x01': sourceLinks['0x01'].slice(), ...(initial.links || {}) }
  };
  const { sp, requests } = createMockSp((req) => {
    const avail = ctId(req.url, 'availablecontenttypes');
    const web = ctId(req.url, 'contenttypes');
    if (req.method === 'GET' && avail && /\/fieldlinks/i.test(req.url)) {
      return { body: state.links[avail] || [] };
    }
    if (req.method === 'GET' && avail) {
      // SharePoint Online answers a missing content type with 200 {"odata.null": true}, not 404 (spike 02).
      const ct = state.cts.find((c) => c.StringId === avail);
      return { body: ct || { 'odata.null': true } };
    }
    const byName = /availablecontenttypes\?\$filter=Name eq '([^']+)'/i.exec(req.url);
    if (req.method === 'GET' && byName) {
      return { body: state.cts.filter((c) => c.Name === byName[1]).map((c) => ({ StringId: c.StringId })) };
    }
    if (req.method === 'POST' && /\/_api\/contextinfo$/i.test(req.url)) {
      return { body: { FormDigestValue: 'digest', WebFullUrl: 'https://fabrikam.sharepoint.com/sites/cel' } };
    }
    if (req.method === 'POST' && web) {
      Object.assign(state.cts.find((c) => c.StringId === web)!, req.body);
      return { status: 204 };
    }
    return undefined;
  });

  // CSOM ProcessQuery: ContentTypes.Add, and FieldLinks.Add / GetById + SetProperty + Update.
  const csomBodies: string[] = [];
  const unescape = (v: string): string => v.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const ok = (...rest: unknown[]): Response => new Response(JSON.stringify([{ SchemaVersion: '15.0.0.0', ErrorInfo: null }, ...rest]), { status: 200 });
  const csomFetch = async (url: string, init: RequestInit): Promise<Response> => {
    expect(url).toBe('https://fabrikam.sharepoint.com/sites/cel/_vti_bin/client.svc/ProcessQuery');
    expect((init.headers as Record<string, string>)['X-RequestDigest']).toBe('digest');
    const xml = String(init.body);
    csomBodies.push(xml);

    if (xml.indexOf('Name="GetById"><Parameters><Parameter Type="String">') < 0) {
      const prop = (n: string): string => unescape((new RegExp(`<Property Name="${n}" Type="String">([^<]*)</Property>`).exec(xml) || [])[1] || '');
      // keepIds: false simulates what REST does (spike 02): a fresh ID under Item.
      const id = keepIds ? prop('Id') : `0x0100${'F'.repeat(32)}`;
      state.cts.push({ StringId: id, Name: prop('Name'), Group: prop('Group'), Description: prop('Description'), SchemaXml: '<ContentType />' });
      state.links[id] = (state.links[parentIdOf(id)!] || []).map((l) => ({ ...l })); // a new type starts with its parent's links
      return ok(10, { IsNull: false }, 11, { StringId: id });
    }

    if (failLinks) {
      // CSOM errors arrive with HTTP 200 and ErrorInfo.
      return new Response(JSON.stringify([{ SchemaVersion: '15.0.0.0', ErrorInfo: { ErrorMessage: 'Column not found', ErrorCode: -1, ErrorTypeName: 'System.ArgumentException' } }]), { status: 200 });
    }
    const ctIdOf = /Name="GetById"><Parameters><Parameter Type="String">([^<]+)</.exec(xml)![1];
    const links = state.links[ctIdOf];
    const byPath: { [path: string]: IFieldLinkInfoLike } = {};
    const fieldByPath: { [path: string]: string } = {};
    xml.replace(/<Method Id="(\d+)" ParentId="6" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">([^<]+)</g, (_m, p: string, n: string) => ((fieldByPath[p] = unescape(n)), ''));
    xml.replace(/<Method Id="(\d+)" ParentId="5" Name="Add"><Parameters><Parameter [^>]+><Property Name="Field" ObjectPathId="(\d+)"/g, (_m, p: string, f: string) => {
      const l = link(fieldByPath[f]);
      links.push(l);
      byPath[p] = l;
      return '';
    });
    xml.replace(/<Method Id="(\d+)" ParentId="5" Name="GetById"><Parameters><Parameter Type="Guid">\{([^}]+)\}/g, (_m, p: string, id: string) => {
      byPath[p] = links.find((l) => l.Id === id)!;
      return '';
    });
    xml.replace(/<SetProperty Id="\d+" ObjectPathId="(\d+)" Name="(Required|Hidden)"><Parameter Type="Boolean">(true|false)</g, (_m, p: string, n: string, v: string) => {
      byPath[p][n as 'Required' | 'Hidden'] = v === 'true';
      return '';
    });
    expect(xml).toContain('<Method Name="Update"');
    return ok();
  };
  return { sp, state, requests, provider: new ContentTypeProvider(csomFetch), csomBodies };
}

function ctx(): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: 'https://fabrikam.sharepoint.com/sites/cel', serverRelativeUrl: '/sites/cel', title: 'Cél' });
  return { targetSiteUrl: 'https://fabrikam.sharepoint.com/sites/cel', tokens, log: new Logger() };
}

async function extractAll(): Promise<IContentType[]> {
  const extractor = new ContentTypeExtractor();
  const { sp } = sourceSp();
  const writer = new JsonTemplateWriter(
    createEmptyTemplate({
      name: 'Teszt',
      createdBy: 'anna@contoso.com',
      createdAt: '2026-09-24T10:00:00Z',
      sourceSiteUrl: 'https://contoso.sharepoint.com/sites/forras',
      sourceTenant: 'contoso.onmicrosoft.com',
      sourceLcid: 1038,
      includesContent: false
    })
  );
  const refs = (await extractor.discover(sp)).map((d) => d.ref).reverse(); // order must not matter
  const tokens = TokenContext.forSite({ absoluteUrl: 'https://contoso.sharepoint.com/sites/forras', serverRelativeUrl: '/sites/forras', title: 'Forrás' });
  await extractor.extract(sp, refs, { includeContent: false, includeVersions: false, includeMembers: false, tokens, log: new Logger() }, writer);
  expect(validateTemplate(writer.manifest).errors).toEqual([]);
  return writer.manifest.contentTypes;
}

describe('content type model', () => {
  it('derives parent IDs for GUID-style and built-in children', () => {
    expect(parentIdOf(PROJEKT_ID)).toBe('0x01');
    expect(parentIdOf(KIEMELT_ID)).toBe(PROJEKT_ID);
    expect(parentIdOf('0x010101')).toBe('0x0101');
    expect(parentIdOf('0x0101')).toBe('0x01');
    expect(parentIdOf('0x01')).toBe('0x');
    expect(parentIdOf('0x')).toBeUndefined();
  });

  it('treats feature-deployed, hidden and known feature-less built-ins as built-in (spike 02)', () => {
    expect([itemCt, documentCt, hiddenCt, sitePageCt, projektCt, kiemeltCt, customPageCt].map(isCustomContentType)).toEqual([
      false, false, false, false, true, true, true
    ]);
    expect(isCustomContentType({ ...sitePageCt, StringId: sitePageCt.StringId.toLowerCase().replace('0X', '0x') })).toBe(false);
  });

  it('reports missing links and differing flags, but not extra target links', () => {
    const def: IContentType = { id: PROJEKT_ID, name: 'Projekt', group: 'CopyJet', description: 'Projekt adatlap', fieldRefs: [{ internalName: 'CJ_Status', required: true }, { internalName: 'CJ_Ugyfel' }] };
    expect(compareContentTypes(def, projektCt, [link('CJ_Status', true), link('CJ_Ugyfel'), link('Extra')])).toEqual([]);
    expect(compareContentTypes(def, { ...projektCt, Name: 'Project' }, [link('CJ_Status')])).toEqual(['name', 'fieldRefs', 'fieldRefFlags']);
  });
});

describe('ContentTypeExtractor', () => {
  it('discovers only custom types', async () => {
    const found = await new ContentTypeExtractor().discover(sourceSp().sp);
    expect(found.map((f) => f.ref.key)).toEqual([`contentType:${PROJEKT_ID}`, `contentType:${KIEMELT_ID}`]);
  });

  it('extracts parents first, with only own (or overridden) field links', async () => {
    const [projekt, kiemelt] = await extractAll();
    expect(projekt).toEqual({
      id: PROJEKT_ID,
      name: 'Projekt',
      group: 'CopyJet',
      description: 'Projekt adatlap',
      parentId: '0x01',
      fieldRefs: [{ internalName: 'CJ_Status', required: true }, { internalName: 'CJ_Ugyfel' }]
    });
    expect(kiemelt.parentId).toBe(PROJEKT_ID);
    expect(kiemelt.fieldRefs).toEqual([{ internalName: 'Title', required: true, hidden: true }, { internalName: 'CJ_Keret' }]);
  });

  it('depends on a custom parent and on the linked site columns', async () => {
    const [projekt, kiemelt] = await extractAll();
    const ex = new ContentTypeExtractor();
    expect(ex.dependencies(projekt)).toEqual([
      { kind: 'contentType', key: 'contentType:0x01' },
      { kind: 'siteField', key: 'field:CJ_Status' },
      { kind: 'siteField', key: 'field:CJ_Ugyfel' }
    ]);
    expect(ex.dependencies(kiemelt)[0]).toEqual({ kind: 'contentType', key: `contentType:${PROJEKT_ID}` });
  });
});

describe('ContentTypeProvider', () => {
  it('creates parent then child with the same IDs, adding only missing links; a rerun skips both', async () => {
    const [projekt, kiemelt] = await extractAll();
    const target = targetSite();
    const { provider } = target;
    const c = ctx();

    expect(await provider.apply(target.sp, projekt, 'skip', c)).toMatchObject({ outcome: 'created' });
    expect(target.csomBodies[0]).toContain(`<Property Name="Id" Type="String">${PROJEKT_ID}</Property>`);
    expect(await provider.apply(target.sp, kiemelt, 'skip', c)).toMatchObject({ outcome: 'created' });
    expect(target.state.cts.map((x) => x.StringId)).toEqual(['0x01', '0x0101', PROJEKT_ID, KIEMELT_ID]);
    // Inherited Title is not linked twice; its overridden Hidden flag is applied through CSOM.
    expect(target.state.links[KIEMELT_ID].map((l) => l.FieldInternalName)).toEqual(['ContentType', 'Title', 'CJ_Status', 'CJ_Ugyfel', 'CJ_Keret']);
    expect(target.state.links[KIEMELT_ID][1]).toMatchObject({ FieldInternalName: 'Title', Required: true, Hidden: true });
    expect(target.state.links[PROJEKT_ID].find((l) => l.FieldInternalName === 'CJ_Status')).toMatchObject({ Required: true });

    const posts = target.requests.filter((r) => r.method === 'POST').length;
    expect((await provider.diff(target.sp, projekt, c)).status).toBe('same');
    expect(await provider.apply(target.sp, projekt, 'update', c)).toMatchObject({ outcome: 'skipped' });
    expect(target.requests.filter((r) => r.method === 'POST').length).toBe(posts);
  });

  it('update mode renames and adds missing links, never removes target links', async () => {
    const [projekt] = await extractAll();
    const target = targetSite({
      cts: [itemCt, { ...projektCt, Name: 'Project', Description: 'Projekt adatlap' }],
      links: { [PROJEKT_ID]: [link('Title', true), link('CJ_Status', true), link('Sajat')] }
    });
    const { provider } = target;
    const c = ctx();

    expect(await provider.diff(target.sp, projekt, c)).toMatchObject({ status: 'different', changes: ['name', 'fieldRefs'] });
    expect(await provider.apply(target.sp, projekt, 'skip', c)).toMatchObject({ outcome: 'skipped' });
    expect(await provider.apply(target.sp, projekt, 'update', c)).toMatchObject({ outcome: 'updated' });
    expect(target.state.cts[1].Name).toBe('Projekt');
    expect(target.state.links[PROJEKT_ID].map((l) => l.FieldInternalName)).toEqual(['Title', 'CJ_Status', 'Sajat', 'CJ_Ugyfel']);
  });

  it('update mode aligns Required/Hidden on existing links', async () => {
    const [projekt] = await extractAll();
    const target = targetSite({ cts: [itemCt, projektCt], links: { [PROJEKT_ID]: [link('Title', true), link('CJ_Status'), link('CJ_Ugyfel', false, true)] } });
    const c = ctx();
    expect(await target.provider.diff(target.sp, projekt, c)).toMatchObject({ status: 'different', changes: ['fieldRefFlags'] });
    expect(await target.provider.apply(target.sp, projekt, 'update', c)).toMatchObject({ outcome: 'updated' });
    expect(target.state.links[PROJEKT_ID].map((l) => [l.FieldInternalName, l.Required, l.Hidden])).toEqual([
      ['Title', true, false],
      ['CJ_Status', true, false],
      ['CJ_Ugyfel', false, false]
    ]);
    expect(target.csomBodies.pop()).toContain('<Parameter Type="Guid">{id-CJ_Status}</Parameter>');
  });

  it('does not create a type whose name is taken by another ID', async () => {
    const [projekt] = await extractAll();
    const target = targetSite({ cts: [itemCt, { ...projektCt, StringId: `0x0100${'E'.repeat(32)}` }] });
    const { provider } = target;
    const c = ctx();
    expect(await provider.diff(target.sp, projekt, c)).toMatchObject({ status: 'unsupported', changes: ['nameConflict'] });
    expect(await provider.apply(target.sp, projekt, 'update', c)).toMatchObject({ outcome: 'skipped' });
    expect(c.log.entries.map((e) => e.code)).toContain('CT_UNSUPPORTED');
  });

  it('fails loudly when SharePoint does not keep the requested ID', async () => {
    const [projekt] = await extractAll();
    const target = targetSite({}, false);
    await expect(target.provider.apply(target.sp, projekt, 'skip', ctx())).rejects.toMatchObject({ code: 'CT_ID_NOT_KEPT' });
  });

  it('wraps a failing field link (e.g. site column not installed) in CT_FIELDLINK_FAILED', async () => {
    const [projekt] = await extractAll();
    const target = targetSite({}, true, true);
    await expect(target.provider.apply(target.sp, projekt, 'skip', ctx())).rejects.toMatchObject({ code: 'CT_FIELDLINK_FAILED' });
  });
});
