import { ListFieldExtractor } from '../../src/core/extractors/ListFieldExtractor';
import { ListFieldProvider } from '../../src/core/providers/ListFieldProvider';
import { parseFieldXml, type IFieldInfoLike } from '../../src/core/fields';
import { isCopiedListField, listFieldDefs, listFieldOrigin, loadSourceSite, toListDef, type IListFieldDef, type IListFieldInfoLike } from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import type { IInstallContext } from '../../src/core/model';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { validateTemplate } from '../../src/core/schema';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { SOURCE_WEB, lookupForras, sourceLists, tesztLista } from '../fixtures/lists';
import { created, siteColumn1, startDate, tesztListaFields, title, valaki, valassz } from '../fixtures/listFields';

const TARGET_WEB = '/sites/Cel';
const TARGET_LOOKUP_ID = '77777777-0000-4000-8000-000000000002';

function sourceSp(): ReturnType<typeof createMockSp> {
  return createMockSp((req) => {
    if (req.method !== 'GET') return undefined;
    if (/\/_api\/web\?\$select=/i.test(req.url)) return { body: SOURCE_WEB };
    if (/\/_api\/web\/lists\?\$select=/i.test(req.url)) return { body: sourceLists };
    const m = /\/getList\('([^']+)'\)\/fields\?/i.exec(req.url);
    if (m) return { body: m[1] === tesztLista.RootFolder.ServerRelativeUrl ? tesztListaFields : [title, created] };
    return undefined;
  }, SOURCE_WEB.Url);
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
  // The ListExtractor's part: list entries the columns are written into.
  writer.manifest.lists.push(toListDef(tesztLista, 'Teszt_lista', SOURCE_WEB.ServerRelativeUrl), toListDef(lookupForras, 'Teszt_lookup_forrs', SOURCE_WEB.ServerRelativeUrl));
  const extractor = new ListFieldExtractor();
  const refs = (await extractor.discover(sp)).map((d) => d.ref);
  await extractor.extract(sp, refs, { includeContent: false, includeVersions: false, includeMembers: false, tokens: site.tokens, log: new Logger() }, writer);
  expect(validateTemplate(writer.manifest).errors).toEqual([]);
  return writer;
}

interface ITargetField extends IFieldInfoLike {
  SourceID?: string;
}

/** Target web: lists with in-memory fields; site columns available for instances (spike 05 behaviour). */
function targetSite(existing: { [listUrl: string]: ITargetField[] } = {}): {
  sp: ReturnType<typeof createMockSp>['sp'];
  fields: { [listUrl: string]: ITargetField[] };
  requests: IMockRequest[];
} {
  const lists: { [url: string]: string } = {
    [`${TARGET_WEB}/Lists/Teszt lista`]: '77777777-0000-4000-8000-000000000001',
    [`${TARGET_WEB}/Lists/Teszt lookup forrs`]: TARGET_LOOKUP_ID
  };
  const siteColumns = [siteColumn1, startDate];
  const fields: { [listUrl: string]: ITargetField[] } = {};
  Object.keys(lists).forEach((u) => (fields[u] = (existing[u] || [{ ...title }]).map((x) => ({ ...x }))));

  const { sp, requests } = createMockSp((req) => {
    let m: RegExpExecArray | null;
    if ((m = /\/availablefields\?\$filter=Id eq guid'([^']+)'/i.exec(req.url))) {
      return { body: siteColumns.filter((c) => c.Id === m![1]).map((c) => ({ SchemaXml: c.SchemaXml })) };
    }
    if (req.method === 'POST' && (m = /\/getList\('([^']+)'\)\/fields\/createfieldasxml$/i.exec(req.url))) {
      const p = (req.body as { parameters: { SchemaXml: string; Options: number } }).parameters;
      expect(p.Options).toBe(12);
      const el = parseFieldXml(p.SchemaXml);
      const f: ITargetField = {
        Id: (el.getAttribute('ID') || '').replace(/[{}]/g, ''),
        InternalName: el.getAttribute('Name')!,
        Title: el.getAttribute('DisplayName')!,
        TypeAsString: el.getAttribute('Type')!,
        SchemaXml: p.SchemaXml,
        SourceID: el.getAttribute('SourceID') || `{${lists[m[1]]}}`
      };
      fields[m[1]].push(f);
      return { body: f };
    }
    if (req.method === 'POST' && (m = /\/getList\('([^']+)'\)\/fields\('([^']+)'\)$/i.exec(req.url))) {
      Object.assign(fields[m[1]].find((x) => x.Id === m![2])!, req.body);
      return { status: 204 };
    }
    if ((m = /\/getList\('([^']+)'\)\/fields\?\$filter=InternalName eq '([^']+)'/i.exec(req.url))) {
      return { body: (fields[m[1]] || []).filter((x) => x.InternalName === m![2]) };
    }
    if ((m = /\/getList\('([^']+)'\)\/fields\?\$filter=Id eq guid'([^']+)'/i.exec(req.url))) {
      return { body: (fields[m[1]] || []).filter((x) => x.Id === m![2]).map((x) => ({ Id: x.Id })) };
    }
    if ((m = /\/getList\('([^']+)'\)\?\$select=Id$/i.exec(req.url))) {
      return lists[m[1]] ? { body: { Id: lists[m[1]] } } : { status: 404, body: {} };
    }
    return undefined;
  });
  return { sp, fields, requests };
}

function ctx(): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: `https://fabrikam.sharepoint.com${TARGET_WEB}`, serverRelativeUrl: TARGET_WEB, title: 'Cél' })
    .set('listkey', 'Teszt_lookup_forrs', TARGET_LOOKUP_ID)
    .set('listurl', 'Teszt_lookup_forrs', 'Lists/Teszt lookup forrs')
    .set('listkey', 'Teszt_lista', '77777777-0000-4000-8000-000000000001')
    .set('listurl', 'Teszt_lista', 'Lists/Teszt lista');
  return { targetSiteUrl: `https://fabrikam.sharepoint.com${TARGET_WEB}`, tokens, log: new Logger() };
}

const defOf = (writer: JsonTemplateWriter, name: string): IListFieldDef => listFieldDefs(writer.manifest).filter((d) => d.field.internalName === name)[0];

describe('list field origin (spike 05)', () => {
  it('tells own columns, site column instances (custom and built-in) and base fields apart', () => {
    const origins = tesztListaFields.map((f) => `${f.InternalName}:${listFieldOrigin(f, tesztLista.Id)}`);
    expect(origins).toEqual([
      'Title:builtIn',
      'Created:builtIn',
      'Valaki:list',
      'V_x00e1_lassz:list',
      'ListaLookup:list',
      'SiteColumn1:siteColumn',
      'StartDate:siteColumn',
      'ListaSzoveg_0:list'
    ]);
    expect(tesztListaFields.filter((x: IListFieldInfoLike) => isCopiedListField(x, tesztLista.Id)).map((x) => x.InternalName)).toEqual([
      'Valaki',
      'V_x00e1_lassz',
      'ListaLookup',
      'SiteColumn1',
      'StartDate'
    ]);
  });
});

describe('ListFieldExtractor', () => {
  it('discovers copied columns under their list', async () => {
    const found = await new ListFieldExtractor().discover(sourceSp().sp);
    expect(found.map((x) => [x.ref.key, x.parentKey])).toEqual([
      ['listField:Teszt_lista/Valaki', 'list:Teszt_lista'],
      ['listField:Teszt_lista/V_x00e1_lassz', 'list:Teszt_lista'],
      ['listField:Teszt_lista/ListaLookup', 'list:Teszt_lista'],
      ['listField:Teszt_lista/SiteColumn1', 'list:Teszt_lista'],
      ['listField:Teszt_lista/StartDate', 'list:Teszt_lista']
    ]);
  });

  it('writes sanitized, tokenized columns into the list entry', async () => {
    const writer = await extract();
    const fields = writer.manifest.lists[0].fields!;
    expect(fields.map((x) => x.internalName)).toEqual(['Valaki', 'V_x00e1_lassz', 'ListaLookup', 'SiteColumn1', 'StartDate']);
    const lookup = fields[2];
    expect(lookup).toMatchObject({ type: 'Lookup', lookupList: '{{listkey:Teszt_lookup_forrs}}', lookupField: 'Title' });
    expect(lookup.schemaXml).not.toMatch(/SourceID|WebId/);
    expect(fields[0].schemaXml).not.toMatch(/ColName|RowOrdinal/);
    expect(fields[3].id).toBe(siteColumn1.Id);
    expect(writer.manifest.lists[1].fields).toBeUndefined();
  });

  it('depends on its list, the lookup target list and a site column of the same name', async () => {
    const writer = await extract();
    expect(new ListFieldExtractor().dependencies(defOf(writer, 'ListaLookup'))).toEqual([
      { kind: 'list', key: 'list:Teszt_lista' },
      { kind: 'list', key: 'list:Teszt_lookup_forrs' },
      { kind: 'siteField', key: 'field:ListaLookup' }
    ]);
  });
});

describe('ListFieldProvider', () => {
  const provider = new ListFieldProvider();
  const LIST = `${TARGET_WEB}/Lists/Teszt lista`;

  it('creates own columns from the template (lookup resolved) and site column instances from the site column', async () => {
    const writer = await extract();
    const target = targetSite();
    const c = ctx();
    for (const name of ['Valaki', 'V_x00e1_lassz', 'ListaLookup', 'SiteColumn1', 'StartDate']) {
      expect(await provider.apply(target.sp, defOf(writer, name), 'skip', c)).toMatchObject({ outcome: 'created' });
    }
    const byName = (n: string): ITargetField => target.fields[LIST].find((x) => x.InternalName === n)!;
    expect(byName('ListaLookup').SchemaXml).toContain(`List="{${TARGET_LOOKUP_ID}}"`);
    expect(byName('ListaLookup').SourceID).toBe('{77777777-0000-4000-8000-000000000001}'); // own column of the target list
    expect(byName('V_x00e1_lassz').Title).toBe('Válassz');
    // Instances keep the site column's ID and SourceID link.
    expect(byName('SiteColumn1')).toMatchObject({ Id: siteColumn1.Id, SourceID: '{8b98bd9e-6008-44fa-a2e2-784ff594eab3}' });
    expect(byName('StartDate')).toMatchObject({ Id: startDate.Id, SourceID: 'http://schemas.microsoft.com/sharepoint/v3' });
    expect(c.log.entries.filter((e) => e.message === 'Site column added to the list.')).toHaveLength(2);

    const posts = target.requests.filter((r) => r.method === 'POST').length;
    expect((await provider.diff(target.sp, defOf(writer, 'ListaLookup'), c)).status).toBe('same');
    expect(await provider.apply(target.sp, defOf(writer, 'Valaki'), 'update', c)).toMatchObject({ outcome: 'skipped' });
    expect(target.requests.filter((r) => r.method === 'POST').length).toBe(posts);
  });

  it('sets the template title when it differs from the SchemaXml DisplayName (multilingual site)', async () => {
    const writer = await extract();
    const target = targetSite();
    const c = ctx();
    const def = defOf(writer, 'Valaki');
    const localized = { ...def, field: { ...def.field, title: 'Felelős személy' } };
    expect(await provider.apply(target.sp, localized, 'skip', c)).toMatchObject({ outcome: 'created' });
    expect(target.fields[LIST].find((x) => x.InternalName === 'Valaki')!.Title).toBe('Felelős személy');
    expect((await provider.diff(target.sp, localized, c)).status).toBe('same');
  });

  it('update mode merges choices, skip mode leaves the column alone', async () => {
    const writer = await extract();
    const existing = { ...valassz, SchemaXml: valassz.SchemaXml.replace(/<CHOICES>.*<\/CHOICES>/, '<CHOICES><CHOICE>Egy</CHOICE><CHOICE>Három</CHOICE></CHOICES>') };
    const target = targetSite({ [LIST]: [title, existing] });
    const c = ctx();
    const def = defOf(writer, 'V_x00e1_lassz');
    expect(await provider.diff(target.sp, def, c)).toMatchObject({ status: 'different', changes: ['choices'] });
    expect(await provider.apply(target.sp, def, 'skip', c)).toMatchObject({ outcome: 'skipped' });
    expect(await provider.apply(target.sp, def, 'update', c)).toMatchObject({ outcome: 'updated' });
    expect(target.requests.filter((r) => r.method === 'POST').pop()!.body).toEqual({ Choices: ['Egy', 'Három', 'Kettő'] });
  });

  it('reports a missing list, an ID conflict and needs the lookup target list token', async () => {
    const writer = await extract();
    const c = ctx();
    const def = defOf(writer, 'ListaLookup');
    expect(await provider.diff(targetSite().sp, { ...def, listKey: 'Nincs', listUrl: 'Lists/Nincs' }, c)).toMatchObject({ status: 'unsupported', changes: ['listMissing'] });

    const conflict = targetSite({ [LIST]: [title, { ...valaki, InternalName: 'Masik' }] });
    expect(await provider.diff(conflict.sp, defOf(writer, 'Valaki'), c)).toMatchObject({ status: 'unsupported', changes: ['idConflict'] });

    const bare: IInstallContext = { ...c, tokens: TokenContext.forSite({ absoluteUrl: c.targetSiteUrl, serverRelativeUrl: TARGET_WEB, title: 'Cél' }) };
    await expect(provider.apply(targetSite().sp, def, 'skip', bare)).rejects.toMatchObject({ code: 'TOKEN_UNRESOLVED' });
  });

  it('installs into a renamed copy of the list through the {listurl} token', async () => {
    const writer = await extract();
    const target = targetSite();
    target.fields[`${TARGET_WEB}/Lists/Teszt lista_copy`] = [{ ...title }];
    const install = ctx();
    install.tokens.set('listurl', 'Teszt_lista', 'Lists/Teszt lista_copy');
    // The copy is unknown to the mock's list registry → reported as a missing list, proving the token is used.
    expect(await provider.diff(target.sp, defOf(writer, 'Valaki'), install)).toMatchObject({ status: 'unsupported', changes: ['listMissing'] });
    expect(target.requests.some((r) => r.url.indexOf("getList('/sites/Cel/Lists/Teszt lista_copy')") > 0)).toBe(true);
  });
});
