import { SiteFieldExtractor } from '../../src/core/extractors/SiteFieldExtractor';
import { SiteFieldProvider } from '../../src/core/providers/SiteFieldProvider';
import { parseFieldXml, type IFieldInfoLike } from '../../src/core/fields';
import { Logger } from '../../src/core/logger';
import type { IField, IInstallContext } from '../../src/core/model';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { validateTemplate } from '../../src/core/schema';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { UGYFELEK_LIST_ID, allSourceFields, builtInField, statusField } from '../fixtures/fields';

const TARGET_UGYFELEK_ID = '11111111-2222-4333-8444-555555555555';

function sourceTokens(): TokenContext {
  return TokenContext.forSite({
    absoluteUrl: 'https://contoso.sharepoint.com/sites/forras',
    serverRelativeUrl: '/sites/forras',
    title: 'Forrás'
  }).set('listkey', 'Ugyfelek', UGYFELEK_LIST_ID);
}

function targetContext(): IInstallContext {
  const tokens = TokenContext.forSite({
    absoluteUrl: 'https://fabrikam.sharepoint.com/sites/cel',
    serverRelativeUrl: '/sites/cel',
    title: 'Cél'
  }).set('listkey', 'Ugyfelek', TARGET_UGYFELEK_ID);
  return { targetSiteUrl: 'https://fabrikam.sharepoint.com/sites/cel', tokens, log: new Logger() };
}

/** Source site: answers the custom-fields query with the fixtures. */
function sourceSp(): ReturnType<typeof createMockSp> {
  return createMockSp((req) => {
    if (req.method === 'GET' && req.url.indexOf('/_api/web/fields?') > 0) {
      expect(req.url).toContain('$filter=Hidden eq false');
      return { body: [...allSourceFields, builtInField] };
    }
    return undefined;
  }, 'https://contoso.sharepoint.com/sites/forras');
}

/** Target site with an in-memory field collection supporting the calls SiteFieldProvider makes. */
function targetSite(initial: IFieldInfoLike[] = []): { sp: ReturnType<typeof createMockSp>['sp']; fields: IFieldInfoLike[]; requests: IMockRequest[] } {
  const fields = initial.map((f) => ({ ...f }));
  let seq = 0;
  const { sp, requests } = createMockSp((req) => {
    const byName = /availablefields\?\$filter=InternalName eq '([^']+)'/.exec(req.url);
    if (req.method === 'GET' && byName) {
      return { body: fields.filter((f) => f.InternalName === byName[1]) };
    }
    const byId = /availablefields\?\$filter=Id eq guid'([^']+)'/.exec(req.url);
    if (req.method === 'GET' && byId) {
      return { body: fields.filter((f) => f.Id === byId[1]).map((f) => ({ Id: f.Id })) };
    }
    if (req.method === 'POST' && req.url.endsWith('/_api/web/fields/createfieldasxml')) {
      const p = (req.body as { parameters: { SchemaXml: string; Options: number } }).parameters;
      expect(p.Options).toBe(8);
      const el = parseFieldXml(p.SchemaXml);
      const id = (el.getAttribute('ID') || `00000000-0000-4000-8000-${String(++seq).padStart(12, '0')}`).replace(/[{}]/g, '');
      const created: IFieldInfoLike = {
        Id: id,
        InternalName: el.getAttribute('Name')!,
        Title: el.getAttribute('DisplayName')!,
        TypeAsString: el.getAttribute('Type')!,
        Group: el.getAttribute('Group') || '',
        Description: el.getAttribute('Description') || '',
        Required: el.getAttribute('Required') === 'TRUE',
        Hidden: false,
        SchemaXml: p.SchemaXml
      };
      fields.push(created);
      return { body: created };
    }
    const update = /availablefields\('([^']+)'\)$/.exec(req.url);
    if (req.method === 'POST' && update) {
      const f = fields.find((x) => x.Id === update[1])!;
      const props = req.body as { Title?: string; Choices?: string[] };
      if (props.Title) f.Title = props.Title;
      if (props.Choices) {
        f.SchemaXml = f.SchemaXml.replace(/<CHOICES>.*<\/CHOICES>/, `<CHOICES>${props.Choices.map((c) => `<CHOICE>${c}</CHOICE>`).join('')}</CHOICES>`);
      }
      return { status: 204 };
    }
    return undefined;
  });
  return { sp, fields, requests };
}

async function extractAll(): Promise<{ writer: JsonTemplateWriter; log: Logger }> {
  const extractor = new SiteFieldExtractor();
  const { sp } = sourceSp();
  const discovered = await extractor.discover(sp);
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
  const log = new Logger();
  await extractor.extract(
    sp,
    discovered.map((d) => d.ref),
    { includeContent: false, includeVersions: false, includeMembers: false, tokens: sourceTokens(), log },
    writer
  );
  return { writer, log };
}

describe('SiteFieldExtractor', () => {
  it('discovers only custom site columns (by SourceID), with their group', async () => {
    const found = await new SiteFieldExtractor().discover(sourceSp().sp);
    expect(found.map((f) => f.ref.key)).toEqual(['field:CJ_Status', 'field:CJ_Ugyfel', 'field:CJ_Terulet', 'field:CJ_Keret']);
    expect(found[0]).toMatchObject({ title: 'Státusz', group: 'CopyJet' });
  });

  it('extracts a schema-valid, sanitized and tokenized definition', async () => {
    const { writer, log } = await extractAll();
    const t = writer.manifest;
    expect(validateTemplate(t).errors).toEqual([]);
    expect(t.siteFields).toHaveLength(4);

    const status = t.siteFields[0];
    expect(status.id).toBe('8c3a1d52-3b4e-4f7a-9c21-5d6e7f8a9b01');
    expect(status.schemaXml).toContain('ID="{8c3a1d52-3b4e-4f7a-9c21-5d6e7f8a9b01}"');
    expect(status.schemaXml).not.toMatch(/SourceID|Version=|JSLink/);

    const lookup = t.siteFields[1];
    expect(lookup.lookupList).toBe('{{listkey:Ugyfelek}}');
    expect(lookup.schemaXml).toContain('List="{{listkey:Ugyfelek}}"');
    expect(lookup.schemaXml).not.toMatch(/WebId|ColName|RowOrdinal/);

    expect(log.entries.filter((e) => e.code === 'FIELD_XML_SANITIZED')).toHaveLength(4);
  });

  it('reports lookups to lists outside the template (once) and missing columns', async () => {
    const extractor = new SiteFieldExtractor();
    // Start from a manifest that already carries the warning: extracting again must not duplicate it.
    const writer = new JsonTemplateWriter(createEmptyTemplate({ ...(await extractAll()).writer.manifest.meta }));
    writer.manifest.meta.warnings = [{ code: 'LOOKUP_TARGET_UNKNOWN', message: 'x', artifact: 'field:CJ_Ugyfel' }];
    const log = new Logger();
    const tokens = TokenContext.forSite({ absoluteUrl: 'https://contoso.sharepoint.com/sites/forras', serverRelativeUrl: '/sites/forras', title: 'Forrás' });
    await extractor.extract(
      sourceSp().sp,
      [{ kind: 'siteField', key: 'field:CJ_Ugyfel' }, { kind: 'siteField', key: 'field:Gone' }],
      { includeContent: false, includeVersions: false, includeMembers: false, tokens, log },
      writer
    );
    expect(writer.manifest.meta.warnings).toEqual([expect.objectContaining({ code: 'LOOKUP_TARGET_UNKNOWN', artifact: 'field:CJ_Ugyfel' })]);
    expect(log.entries.map((e) => e.code)).toContain('FIELD_NOT_FOUND');
  });

  it('declares the lookup target list as a dependency', () => {
    const def = { internalName: 'X', type: 'Lookup', title: 'X', schemaXml: '<Field />', lookupList: '{listkey:Ugyfelek}', lookupField: 'Title' } as IField;
    expect(new SiteFieldExtractor().dependencies(def)).toEqual([{ kind: 'list', key: 'list:Ugyfelek' }]);
  });
});

describe('SiteFieldProvider', () => {
  const provider = new SiteFieldProvider();

  it('creates missing columns with resolved tokens and registers {fieldid}', async () => {
    const { writer } = await extractAll();
    const target = targetSite();
    const ctx = targetContext();
    const [status, lookup, taxonomy, calculated] = writer.manifest.siteFields;

    expect((await provider.diff(target.sp, status, ctx)).status).toBe('new');
    expect(await provider.apply(target.sp, status, 'skip', ctx)).toMatchObject({ outcome: 'created' });
    expect(ctx.tokens.get('fieldid', 'CJ_Status')).toBe('8c3a1d52-3b4e-4f7a-9c21-5d6e7f8a9b01');

    await provider.apply(target.sp, lookup, 'skip', ctx);
    const created = target.fields.find((f) => f.InternalName === 'CJ_Ugyfel')!;
    expect(created.SchemaXml).toContain(`List="{${TARGET_UGYFELEK_ID}}"`);

    expect(await provider.apply(target.sp, calculated, 'skip', ctx)).toMatchObject({ outcome: 'created' });

    // Taxonomy needs term mapping (phase 2): reported, not created.
    expect((await provider.diff(target.sp, taxonomy, ctx)).status).toBe('unsupported');
    expect(await provider.apply(target.sp, taxonomy, 'update', ctx)).toMatchObject({ outcome: 'skipped' });
    expect(target.fields.map((f) => f.InternalName)).toEqual(['CJ_Status', 'CJ_Ugyfel', 'CJ_Keret']);
  });

  it('is idempotent: a second run finds the same column and skips it', async () => {
    const { writer } = await extractAll();
    const target = targetSite();
    const ctx = targetContext();
    const status = writer.manifest.siteFields[0];
    await provider.apply(target.sp, status, 'update', ctx);
    const posts = (): number => target.requests.filter((r) => r.method === 'POST').length;
    const before = posts();

    expect((await provider.diff(target.sp, status, ctx)).status).toBe('same');
    expect(await provider.apply(target.sp, status, 'update', ctx)).toMatchObject({ outcome: 'skipped' });
    expect(posts()).toBe(before);
  });

  it('update mode merges choices and title without removing target choices', async () => {
    const { writer } = await extractAll();
    const existing: IFieldInfoLike = {
      ...statusField,
      Title: 'Status',
      SchemaXml: statusField.SchemaXml.replace(/<CHOICES>.*<\/CHOICES>/, '<CHOICES><CHOICE>Nyitott</CHOICE><CHOICE>Archivált</CHOICE></CHOICES>')
    };
    const target = targetSite([existing]);
    const ctx = targetContext();
    const status = writer.manifest.siteFields[0];

    const diff = await provider.diff(target.sp, status, ctx);
    expect(diff).toMatchObject({ status: 'different', changes: ['title', 'choices'] });

    expect(await provider.apply(target.sp, status, 'skip', ctx)).toMatchObject({ outcome: 'skipped' });
    expect(target.fields[0].Title).toBe('Status');

    expect(await provider.apply(target.sp, status, 'update', ctx)).toMatchObject({ outcome: 'updated' });
    const update = target.requests.filter((r) => r.method === 'POST').pop()!;
    expect(update.body).toEqual({ Title: 'Státusz', Choices: ['Nyitott', 'Archivált', 'Folyamatban', 'Lezárt'] });
  });

  it('does not create a column whose ID is taken by another column', async () => {
    const { writer } = await extractAll();
    const target = targetSite([{ ...statusField, InternalName: 'Other_Name' }]);
    const status = writer.manifest.siteFields[0];
    const ctx = targetContext();
    expect(await provider.diff(target.sp, status, ctx)).toMatchObject({ status: 'unsupported', changes: ['idConflict'] });
    expect(await provider.apply(target.sp, status, 'update', ctx)).toMatchObject({ outcome: 'skipped' });
    expect(target.fields).toHaveLength(1);
  });

  it('fails with TOKEN_UNRESOLVED when a lookup target list is not installed yet', async () => {
    const { writer } = await extractAll();
    const ctx = targetContext();
    const bare: IInstallContext = { ...ctx, tokens: TokenContext.forSite({ absoluteUrl: ctx.targetSiteUrl, serverRelativeUrl: '/sites/cel', title: 'Cél' }) };
    await expect(provider.apply(targetSite().sp, writer.manifest.siteFields[1], 'skip', bare)).rejects.toMatchObject({ code: 'TOKEN_UNRESOLVED' });
  });
});
