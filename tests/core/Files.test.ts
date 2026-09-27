import { FileExtractor } from '../../src/core/extractors/FileExtractor';
import { CHUNK_SIZE, listFilesDefs, type IListFilesDef } from '../../src/core/files';
import { loadSourceSite, toListDef } from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import { PrincipalMapper } from '../../src/core/mapping';
import type { IFilesMetaFile, IInstallContext, ITemplateReader } from '../../src/core/model';
import { ZipTemplateWriter, createEmptyTemplate, openTemplate } from '../../src/core/packager';
import { buildPlan } from '../../src/core/planner';
import { FileProvider } from '../../src/core/providers/FileProvider';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { SOURCE_WEB, documents, sourceLists } from '../fixtures/lists';

const LIB = documents.RootFolder.ServerRelativeUrl; // /sites/Forras/Shared Documents
const ANNA = 'i:0#.f|membership|anna@contoso.com';
const DOC_CT = `0x0101${'A'.repeat(32)}`;
const SOURCE_DOC_CT = `${DOC_CT}00${'B'.repeat(32)}`;
const TARGET_DOC_CT = `${DOC_CT}00${'C'.repeat(32)}`;

type Raw = { [k: string]: unknown };
const file = (ID: number, dir: string, name: string, length: number, extra: Raw = {}): Raw => ({
  ID,
  FSObjType: 0,
  FileRef: `${dir}/${name}`,
  FileDirRef: dir,
  FileLeafRef: name,
  ContentTypeId: SOURCE_DOC_CT,
  AuthorId: 7,
  EditorId: 7,
  Created: '2026-01-15T10:00:00Z',
  Modified: '2026-01-16T11:30:00Z',
  Title: name,
  File: { Length: String(length), CheckOutType: 2 },
  ...extra
});

const sourceRows: Raw[] = [
  // REST exposes _ExtendedDescription as OData__ExtendedDescription (1.4.0.0 selected the bare name: HTTP 400).
  file(1, LIB, 'Kép.png', 5, { OData__ExtendedDescription: 'Logó' }),
  { ID: 2, FSObjType: 1, FileRef: `${LIB}/2026`, FileDirRef: LIB, FileLeafRef: '2026' },
  file(3, `${LIB}/2026`, 'jelentés.docx', 8, { Modified: '2026-03-01T09:00:00Z' }),
  file(4, LIB, 'kivett.docx', 3, { File: { Length: '3', CheckOutType: 0 } }),
  file(5, LIB, 'nagy.bin', 300 * 1024 * 1024)
];

function sourceSp(): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  return createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (req.method === 'GET' && /\/_api\/web\?\$select=/i.test(req.url)) return { body: SOURCE_WEB };
    if (req.method === 'GET' && /\/_api\/web\/lists\?\$select=/i.test(req.url)) return { body: sourceLists };
    if (req.method === 'GET' && /\/_api\/web\/siteusers\?\$select=/i.test(req.url)) return { body: [{ Id: 7, LoginName: ANNA, Title: 'Kiss Anna', Email: 'anna@contoso.com' }] };
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/fields\?\$select=/i.test(req.url)) {
      return {
        body: [
          { InternalName: 'FileLeafRef', TypeAsString: 'File', Hidden: false, ReadOnlyField: false },
          { InternalName: 'Title', TypeAsString: 'Text', Hidden: false, ReadOnlyField: false },
          { InternalName: '_ExtendedDescription', TypeAsString: 'Note', Hidden: false, ReadOnlyField: false },
          { InternalName: 'MediaServiceImageTags', TypeAsString: 'TaxonomyFieldTypeMulti', Hidden: false, ReadOnlyField: false }
        ]
      };
    }
    if (req.method === 'GET' && /\/rootFolder\?\$select=ContentTypeOrder$/i.test(req.url)) return { body: { ContentTypeOrder: [{ StringValue: SOURCE_DOC_CT }] } };
    if (req.method === 'GET' && /\/contentTypes\?\$select=StringId,Parent\/StringId&\$expand=Parent$/i.test(req.url)) {
      return { body: [{ StringId: SOURCE_DOC_CT, Parent: { StringId: DOC_CT } }] };
    }
    if (req.method === 'GET' && /\/items\?\$select=[^&]*[,=]_ExtendedDescription/i.test(req.url)) {
      return { status: 400, body: { 'odata.error': { message: { value: 'A(z) _ExtendedDescription mező vagy tulajdonság nem létezik.' } } } };
    }
    if (req.method === 'GET' && (m = /\/getList\('[^']+'\)\/items\?.*\$filter=ID gt (\d+)/i.exec(req.url))) {
      return { body: sourceRows.filter((r) => (r.ID as number) > Number(m![1])) };
    }
    if (req.method === 'GET' && (m = /\/getFileByServerRelativePath\(decodedUrl='([^']+)'\)\/versions\?\$select=/i.exec(req.url))) {
      return m[1].indexOf('jelentés') >= 0
        ? {
            body: [
              { ID: 1024, VersionLabel: '2.0', Created: '2026-02-01T08:00:00Z', Size: 7, CheckInComment: 'második', CreatedBy: { Id: 7 } },
              { ID: 512, VersionLabel: '1.0', Created: '2026-01-15T10:00:00Z', Size: 6, CreatedBy: { Id: 9 } }
            ]
          }
        : { body: [] };
    }
    if (req.method === 'GET' && (m = /\/getFileByServerRelativePath\(decodedUrl='([^']+)'\)\/versions\((\d+)\)\/\$value$/i.exec(req.url))) {
      return { body: `v${m[2]} of ${m[1].split('/').pop()}` };
    }
    if (req.method === 'GET' && (m = /\/getFileByServerRelativePath\(decodedUrl='([^']+)'\)\/\$value$/i.exec(req.url))) return { body: `current ${m[1].split('/').pop()}` };
    return undefined;
  }, SOURCE_WEB.Url);
}

async function extractPackage(log = new Logger()): Promise<ITemplateReader> {
  const { sp } = sourceSp();
  const site = await loadSourceSite(sp);
  const writer = new ZipTemplateWriter(
    createEmptyTemplate({ name: 'Fájlok', createdBy: 'anna@contoso.com', createdAt: '2026-09-27T10:00:00Z', sourceSiteUrl: SOURCE_WEB.Url, sourceTenant: 'contoso', sourceLcid: 1038, includesContent: false })
  );
  writer.manifest.lists.push(toListDef(documents, 'Shared_Documents', SOURCE_WEB.ServerRelativeUrl));
  await new FileExtractor().extract(
    sp,
    [{ kind: 'files', key: 'files:Shared_Documents' }],
    { includeContent: true, includeVersions: true, includeMembers: false, versionsFor: ['Shared_Documents'], tokens: site.tokens, log },
    writer
  );
  return openTemplate(await writer.finalize());
}

describe('FileExtractor', () => {
  it('packs files with metadata and earlier versions; skips checked-out and too large files', async () => {
    const log = new Logger();
    const reader = await extractPackage(log);
    expect(log.entries.filter((e) => e.level === 'warn').map((e) => e.code)).toEqual(['FILE_CHECKED_OUT', 'FILE_TOO_LARGE']);
    expect(reader.manifest.lists[0].content).toEqual({ mode: 'files', sourceMode: 'embedded', includeVersions: true, fileCount: 2, sizeBytes: 5 + 8 + 7 + 6, source: 'files/Shared_Documents/' });
    const meta = await reader.getJson<IFilesMetaFile>('files/Shared_Documents/_meta.json');
    expect(meta.files).toEqual([
      {
        path: 'Kép.png',
        blob: 'files/Shared_Documents/Kép.png',
        sizeBytes: 5,
        sourceId: 1,
        contentType: DOC_CT,
        values: { Title: 'Kép.png', _ExtendedDescription: 'Logó' },
        system: { author: '{principal:anna}', editor: '{principal:anna}', created: '2026-01-15T10:00:00Z', modified: '2026-01-16T11:30:00Z' }
      },
      {
        path: '2026/jelentés.docx',
        blob: 'files/Shared_Documents/2026/jelentés.docx',
        sizeBytes: 8,
        sourceId: 3,
        contentType: DOC_CT,
        values: { Title: 'jelentés.docx' },
        system: { author: '{principal:anna}', editor: '{principal:anna}', created: '2026-01-15T10:00:00Z', modified: '2026-03-01T09:00:00Z' },
        // Oldest first; an unknown author (ID 9) keeps only the date.
        versions: [
          { label: '1.0', blob: 'files/Shared_Documents/_v/1.0/2026/jelentés.docx', system: { modified: '2026-01-15T10:00:00Z' } },
          { label: '2.0', blob: 'files/Shared_Documents/_v/2.0/2026/jelentés.docx', system: { editor: '{principal:anna}', modified: '2026-02-01T08:00:00Z' }, comment: 'második' }
        ]
      }
    ]);
    expect(await (await reader.getBlob('files/Shared_Documents/_v/1.0/2026/jelentés.docx')).text()).toBe('"v512 of jelentés.docx"');
  });
});

// ---------------------------------------------------------------------------------------------------------
// Target library: files by path with their uploads, chunked sessions, folders and item values.

const TARGET = { Url: 'https://fabrikam.sharepoint.com/sites/Cel', ServerRelativeUrl: '/sites/Cel', Title: 'Cél' };
const T_LIB = '/sites/Cel/Shared Documents';

interface ITargetFile {
  id: number;
  uploads: number[];
  values: { [name: string]: string };
}

interface ITargetLib {
  files: { [path: string]: ITargetFile };
  folders: string[];
  chunks: string[];
}

function budapest(iso: string): string {
  return new Date(Date.parse(iso) + 3600000).toISOString().slice(0, 19);
}

/** Bytes of a mock download body (the mock answers JSON: a quoted string). */
const bytes = (text: string): number => new TextEncoder().encode(JSON.stringify(text)).length;

function sizeOf(body: unknown): number {
  if (body instanceof Blob) return body.size;
  if (body instanceof ArrayBuffer) return body.byteLength;
  if (body && typeof (body as Uint8Array).byteLength === 'number') return (body as Uint8Array).byteLength;
  return typeof body === 'string' ? body.length : 0;
}

function targetSp(lib: ITargetLib): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  let seq = 100;
  const rel = (url: string): string => url.slice(T_LIB.length + 1);
  const mock = createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/items\?\$filter=FSObjType eq 0&\$select=Id&\$top=1$/i.test(req.url)) {
      return { body: Object.keys(lib.files).slice(0, 1).map((p) => ({ Id: lib.files[p].id })) };
    }
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/items\?\$select=ID,FSObjType,FileRef&\$filter=ID gt 0/i.test(req.url)) {
      return { body: Object.keys(lib.files).map((p) => ({ ID: lib.files[p].id, FSObjType: 0, FileRef: `${T_LIB}/${p}` })) };
    }
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/items\?\$select=ID,FSObjType,FileRef&\$filter=ID gt \d+/i.test(req.url)) return { body: [] };
    if (req.method === 'GET' && /\/getList\('[^']+'\)\/fields\?\$select=/i.test(req.url)) {
      return {
        body: [
          { InternalName: 'Title', TypeAsString: 'Text', Hidden: false, ReadOnlyField: false, SchemaXml: '<Field Name="Title" />' },
          { InternalName: '_ExtendedDescription', TypeAsString: 'Note', Hidden: false, ReadOnlyField: false, SchemaXml: '<Field Name="_ExtendedDescription" />' }
        ]
      };
    }
    if (req.method === 'GET' && /\/_api\/web\/regionalsettings\?\$select=/i.test(req.url)) return { body: { LocaleId: 1038, DecimalSeparator: ',', Time24: true } };
    if (req.method === 'POST' && (m = /\/regionalsettings\/timezone\/utctolocaltime\('([^']+)'\)/i.exec(req.url))) return { body: { value: budapest(m[1]) } };
    if (req.method === 'POST' && /\/_api\/web\/ensureuser$/i.test(req.url)) return { body: { Id: 5, LoginName: (req.body as { logonName: string }).logonName } };
    if (req.method === 'GET' && (m = /\/getFolderByServerRelativePath\(decodedUrl='([^']+)'\)\/folders\?\$select=Name$/i.exec(req.url))) {
      const r = m[1] === T_LIB ? '' : rel(m[1]);
      return { body: lib.folders.filter((f) => (r ? f.indexOf(`${r}/`) === 0 : f.indexOf('/') < 0)).map((f) => ({ Name: f.split('/').pop() })) };
    }
    if (req.method === 'POST' && (m = /\/_api\/web\/folders\/addUsingPath\(DecodedUrl='([^']+)'/i.exec(req.url))) {
      lib.folders.push(rel(m[1]));
      return { body: {} };
    }
    if (req.method === 'GET' && /\/rootFolder\?\$select=ContentTypeOrder$/i.test(req.url)) return { body: { ContentTypeOrder: [{ StringValue: TARGET_DOC_CT }] } };
    if (req.method === 'GET' && /\/contentTypes\?\$select=StringId,Parent\/StringId&\$expand=Parent$/i.test(req.url)) {
      return { body: [{ StringId: TARGET_DOC_CT, Parent: { StringId: DOC_CT } }] };
    }
    // Upload into a folder (spike 10 B1/B2): an existing path without Overwrite is HTTP 400 -2130575257.
    if (req.method === 'POST' && (m = /\/getFolderByServerRelativePath\(decodedUrl='([^']+)'\)\/files\/AddUsingPath\(decodedurl='([^']+)'(,Overwrite=true)?\)$/i.exec(req.url))) {
      const path = (m[1] === T_LIB ? '' : `${rel(m[1])}/`) + m[2];
      const existing = lib.files[path];
      if (existing && !m[3]) return { status: 400, body: { 'odata.error': { code: '-2130575257, Microsoft.SharePoint.SPException', message: { value: 'A fájl már létezik.' } } } };
      const f = existing || (lib.files[path] = { id: ++seq, uploads: [], values: {} });
      f.uploads.push(sizeOf(req.rawBody));
      return { body: { Name: m[2], ServerRelativeUrl: `${m[1]}/${m[2]}` } };
    }
    // Chunked upload session (spike 10 B4).
    if (req.method === 'POST' && (m = /\/getFileByServerRelativePath\(decodedUrl='([^']+)'\)\/(startUpload|continueUpload|finishUpload)\(uploadId=guid'[^']+'(,fileOffset=(\d+))?\)$/i.exec(req.url))) {
      lib.chunks.push(`${m[2]}@${m[4] || 0}`);
      const path = rel(m[1]);
      const offset = Number(m[4] || 0);
      const end = offset + sizeOf(req.rawBody);
      if (/finishUpload/i.test(m[2])) {
        lib.files[path].uploads.push(end);
        return { body: { Name: path } };
      }
      // start/continue answer the new offset.
      return { body: { value: end } };
    }
    if (req.method === 'GET' && (m = /\/getFileByServerRelativePath\(decodedUrl='([^']+)'\)\/listItemAllFields\?\$select=Id$/i.exec(req.url))) {
      return { body: { Id: lib.files[rel(m[1])].id } };
    }
    if (req.method === 'POST' && (m = /\/getList\('[^']+'\)\/items\((\d+)\)\/validateupdatelistitem$/i.exec(req.url))) {
      const f = Object.keys(lib.files).map((p) => lib.files[p]).filter((x) => x.id === Number(m![1]))[0];
      const body = req.body as { formValues: Array<{ FieldName: string; FieldValue: string }> };
      body.formValues.forEach((v) => (f.values[v.FieldName] = v.FieldValue));
      return { body: { value: body.formValues.map((v) => ({ ...v, HasException: false })) } };
    }
    return undefined;
  }, TARGET.Url);
  return mock;
}

function installContext(reader: ITemplateReader, sp: ReturnType<typeof createMockSp>['sp']): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: TARGET.Url, serverRelativeUrl: TARGET.ServerRelativeUrl, title: TARGET.Title });
  tokens.set('listkey', 'Shared_Documents', 'bbbbbbbb-0000-4000-8000-000000000001');
  return { targetSiteUrl: TARGET.Url, tokens, log: new Logger(), content: { reader, idMaps: {}, principals: new PrincipalMapper(sp, reader.manifest.principals) } };
}

describe('FileProvider', () => {
  it('plans files after the library and its content types', async () => {
    const reader = await extractPackage();
    const step = buildPlan(reader.manifest).steps.filter((s) => s.ref.key === 'files:Shared_Documents')[0];
    expect(step.dependsOn).toEqual(['list:Shared_Documents']);
  });

  it('uploads files and their versions oldest first, then the metadata; a rerun keeps what is there', async () => {
    const reader = await extractPackage();
    const lib: ITargetLib = { files: {}, folders: ['Forms'], chunks: [] };
    const target = targetSp(lib);
    const ctx = installContext(reader, target.sp);
    const def: IListFilesDef = listFilesDefs(reader.manifest)[0];
    const provider = new FileProvider();

    expect((await provider.diff(target.sp, def, ctx)).status).toBe('new');
    expect(await provider.apply(target.sp, def, 'skip', ctx)).toMatchObject({ outcome: 'created' });
    expect(lib.folders).toEqual(['Forms', '2026']);
    // Kép.png: one upload. jelentés.docx: 1.0, 2.0, then the current content.
    expect(lib.files['Kép.png'].uploads).toEqual([bytes('current Kép.png')]);
    expect(lib.files['2026/jelentés.docx'].uploads).toEqual([
      bytes('v512 of jelentés.docx'),
      bytes('v1024 of jelentés.docx'),
      bytes('current jelentés.docx')
    ]);
    const report = lib.files['2026/jelentés.docx'].values;
    expect(report).toEqual({
      Title: 'jelentés.docx',
      ContentTypeId: TARGET_DOC_CT,
      Author: `[{"Key":"${ANNA}"}]`,
      Editor: `[{"Key":"${ANNA}"}]`,
      Created: '2026. 01. 15. 11:00',
      Modified: '2026. 03. 01. 10:00'
    });
    // Written under its internal name (only reading needs the OData_ prefix).
    expect(lib.files['Kép.png'].values._ExtendedDescription).toBe('Logó');
    expect(ctx.content!.idMaps.Shared_Documents).toEqual({ 1: 101, 3: 102 });
    expect(ctx.log.counts.warn).toBe(0);

    // Rerun: both paths exist, nothing is uploaded again.
    const again = installContext(reader, target.sp);
    expect((await provider.diff(target.sp, def, again)).status).toBe('different');
    expect(await provider.apply(target.sp, def, 'skip', again)).toMatchObject({ outcome: 'skipped' });
    expect(again.log.entries.map((e) => e.code)).toEqual(['FILES_KEPT']);
    expect(lib.files['Kép.png'].uploads.length).toBe(1);
  });

  it('uploads files above 10 MB in chunks', async () => {
    const writer = new ZipTemplateWriter(createEmptyTemplate({ name: 'Nagy', createdBy: 'a', createdAt: '2026-09-27T10:00:00Z', sourceSiteUrl: SOURCE_WEB.Url, sourceTenant: 'c', sourceLcid: 1038, includesContent: true }));
    const libDef = toListDef(documents, 'Shared_Documents', SOURCE_WEB.ServerRelativeUrl);
    libDef.content = { mode: 'files', sourceMode: 'embedded', fileCount: 1, sizeBytes: CHUNK_SIZE + 5, source: 'files/Shared_Documents/' };
    writer.manifest.lists.push(libDef);
    writer.addJson('files/Shared_Documents/_meta.json', { listKey: 'Shared_Documents', files: [{ path: 'nagy.bin', blob: 'files/Shared_Documents/nagy.bin', sizeBytes: CHUNK_SIZE + 5 }] });
    writer.addBlob('files/Shared_Documents/nagy.bin', new Blob([new Uint8Array(CHUNK_SIZE + 5)]));
    const reader = await openTemplate(await writer.finalize());
    const lib: ITargetLib = { files: {}, folders: [], chunks: [] };
    const target = targetSp(lib);
    const ctx = installContext(reader, target.sp);
    await new FileProvider().apply(target.sp, listFilesDefs(reader.manifest)[0], 'skip', ctx);
    // An empty stub first, then startUpload with a full chunk and finishUpload with the rest.
    expect(lib.files['nagy.bin'].uploads).toEqual([0, CHUNK_SIZE + 5]);
    expect(lib.chunks).toEqual(['startUpload@0', `finishUpload@${CHUNK_SIZE}`]);
  });
});
