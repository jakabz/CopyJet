import type { IListFieldDef } from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import type { IField, IInstallContext } from '../../src/core/model';
import { ListFieldProvider } from '../../src/core/providers/ListFieldProvider';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp } from '../helpers/mockSp';

const WEB = 'https://fabrikam.sharepoint.com/sites/cel';
const LIST = '/sites/cel/Lists/Teszt lista';
// The template's term set (source tenant) and the same-named set in the target tenant's store.
const SOURCE_SET = 'caaf4dd7-a354-4085-b190-05c20a16d9e2';
const TARGET_STORE = '11111111-2222-4333-8444-555555555555';
const TARGET_SET = '66666666-7777-4888-8999-aaaaaaaaaaaa';

const field = (type: 'TaxonomyFieldType' | 'TaxonomyFieldTypeMulti', id?: string): IListFieldDef => ({
  listKey: 'Teszt_lista',
  listUrl: 'Lists/Teszt lista',
  field: {
    id,
    internalName: type === 'TaxonomyFieldType' ? 'Terulet' : 'Cimkek',
    type,
    title: type === 'TaxonomyFieldType' ? 'Terület' : 'Címkék',
    group: 'Egyéni oszlopok',
    schemaXml: '<Field Type="TaxonomyFieldType"><Customization /></Field>',
    termSet: { termSetId: SOURCE_SET, termStoreId: '755cfac8-49b6-4850-b75a-2d387c7676ac', path: 'CopyJet teszt/Terület', isOpen: false }
  } as IField
});

interface ITarget {
  sets: Array<{ id: string; group: string; name: string }>;
  siteColumns: string[];
  csom: string[];
}

function target(t: ITarget): { sp: ReturnType<typeof createMockSp>['sp']; fetch: (url: string, init: RequestInit) => Promise<Response> } {
  const { sp } = createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (req.method === 'POST' && /\/_api\/contextinfo$/i.test(req.url)) return { body: { FormDigestValue: 'digest', WebFullUrl: WEB } };
    if (req.method === 'GET' && /\/getList\('[^']+'\)\?\$select=Id$/i.test(req.url)) return { body: { Id: 'l1' } };
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/fields\?\$filter=(InternalName|Id) eq/i.test(req.url)) return { body: [] };
    if (req.method === 'GET' && (m = /\/availablefields\?\$filter=Id eq guid'([^']+)'/i.exec(req.url))) {
      return { body: t.siteColumns.indexOf(m[1]) >= 0 ? [{ Id: m[1] }] : [] };
    }
    if (/\/_api\/v2\.1\/termStore\?\$select=id$/i.test(req.url)) return { body: { id: TARGET_STORE } };
    if ((m = /\/termStore\/sets\/([0-9a-f-]+)\?\$expand=parentGroup$/i.exec(req.url))) {
      const set = t.sets.filter((s) => s.id === m![1])[0];
      return set ? { body: { id: set.id, localizedNames: [{ name: set.name, isDefault: true }], parentGroup: { name: set.group } } } : { status: 404, body: {} };
    }
    if (/\/termStore\/groups\?\$select=id,name$/i.test(req.url)) {
      return { body: { value: t.sets.map((s, i) => ({ id: `g${i}`, name: s.group })) } };
    }
    if ((m = /\/termStore\/groups\/g(\d+)\/sets\?/i.exec(req.url))) {
      const s = t.sets[Number(m[1])];
      return { body: { value: [{ id: s.id, localizedNames: [{ name: s.name }] }] } };
    }
    return undefined;
  }, WEB);
  const fetch = async (url: string, init: RequestInit): Promise<Response> => {
    expect(url).toBe(`${WEB}/_vti_bin/client.svc/ProcessQuery`);
    const body = String(init.body);
    t.csom.push(body);
    const name = (/Name=&quot;(\w+)&quot;/.exec(body) || [])[1] || 'Terulet';
    return new Response(JSON.stringify([{ SchemaVersion: '15.0.0.0' }, 20, { Id: '/Guid(bbbbbbbb-0000-4000-8000-000000000001)/', InternalName: name, Title: name }]), { status: 200 });
  };
  return { sp, fetch };
}

function ctx(): IInstallContext {
  return { targetSiteUrl: WEB, tokens: TokenContext.forSite({ absoluteUrl: WEB, serverRelativeUrl: '/sites/cel', title: 'Cél' }), log: new Logger() };
}

describe('Managed Metadata columns (spike 11 C2)', () => {
  it('creates a list column through CSOM, bound to the set found by its path in another tenant', async () => {
    const t: ITarget = { sets: [{ id: TARGET_SET, group: 'CopyJet teszt', name: 'Terület' }], siteColumns: [], csom: [] };
    const { sp, fetch } = target(t);
    const provider = new ListFieldProvider(fetch);
    const c = ctx();
    expect((await provider.diff(sp, field('TaxonomyFieldTypeMulti'), c)).status).toBe('new');
    expect(await provider.apply(sp, field('TaxonomyFieldTypeMulti'), 'skip', c)).toMatchObject({ outcome: 'created' });
    const body = t.csom[0];
    // A plain field (no Customization / List / TextField), then the target store and set, then Update.
    expect(body).toContain(`<Method Id="3" ParentId="2" Name="GetList"><Parameters><Parameter Type="String">${LIST}</Parameter>`);
    expect(body).toContain('Name="AddFieldAsXml"><Parameters><Parameter Type="String">&lt;Field Type=&quot;TaxonomyFieldTypeMulti&quot; Name=&quot;Cimkek&quot;');
    expect(body).toContain('Mult=&quot;TRUE&quot;');
    expect(body).not.toContain('Customization');
    expect(body).toContain(`Name="SspId"><Parameter Type="Guid">{${TARGET_STORE}}</Parameter>`);
    expect(body).toContain(`Name="TermSetId"><Parameter Type="Guid">{${TARGET_SET}}</Parameter>`);
    expect(body).toContain('<Parameter Type="Enum">12</Parameter>');
    expect(body).toContain('<Method Name="Update" Id="15" ObjectPathId="5" />');
    expect(c.log.entries.map((e) => e.message)).toEqual(['List column created (Managed Metadata, term set found by path).']);
  });

  it('adds an existing Managed Metadata site column to the list as it is', async () => {
    const siteColumn = 'a44669a5-4dd3-43c6-a515-0b10558cc317';
    const t: ITarget = { sets: [], siteColumns: [siteColumn], csom: [] };
    const { sp, fetch } = target(t);
    await new ListFieldProvider(fetch).apply(sp, field('TaxonomyFieldType', siteColumn), 'skip', ctx());
    expect(t.csom[0]).toContain(`<Property Id="6" ParentId="2" Name="AvailableFields" /><Method Id="7" ParentId="6" Name="GetById"><Parameters><Parameter Type="Guid">{${siteColumn}}</Parameter>`);
    expect(t.csom[0]).toContain('<Method Id="5" ParentId="4" Name="Add"><Parameters><Parameter ObjectPathId="7" /></Parameters></Method>');
  });

  it('skips a column whose term set the target does not have – nothing half-made', async () => {
    const t: ITarget = { sets: [{ id: TARGET_SET, group: 'Más csoport', name: 'Terület' }], siteColumns: [], csom: [] };
    const { sp, fetch } = target(t);
    const c = ctx();
    expect(await new ListFieldProvider(fetch).diff(sp, field('TaxonomyFieldType'), c)).toMatchObject({ status: 'unsupported', changes: ['termSetMissing'] });
    expect(await new ListFieldProvider(fetch).apply(sp, field('TaxonomyFieldType'), 'skip', c)).toMatchObject({ outcome: 'skipped' });
    expect(t.csom).toEqual([]);
  });
});
