import { csomCreateContentType, csomUpdateFieldLinks, processQuery } from '../../src/core/http/csom';
import { CSOM_TYPE, escapeXml, fieldLinksBody } from '../../src/core/http/csomXml';
import { createMockSp } from '../helpers/mockSp';

function spWithDigest(): ReturnType<typeof createMockSp> {
  return createMockSp((req) =>
    req.method === 'POST' && /\/_api\/contextinfo$/i.test(req.url)
      ? { body: { FormDigestValue: 'digest-1', WebFullUrl: 'https://fabrikam.sharepoint.com/sites/cel/' } }
      : undefined
  );
}

const json = (body: unknown, status = 200): Response => new Response(JSON.stringify(body), { status });

describe('CSOM client', () => {
  it('escapes XML special characters', () => {
    expect(escapeXml(`a & <b> "c" 'd'`)).toBe('a &amp; &lt;b&gt; &quot;c&quot; &apos;d&apos;');
  });

  it('posts XML with the digest to the web ProcessQuery endpoint', async () => {
    const fetchMock = jest.fn().mockResolvedValue(json([{ SchemaVersion: '15.0.0.0', ErrorInfo: null }]));
    await processQuery(spWithDigest().sp, { actions: '<A />', objectPaths: '<B />' }, undefined, fetchMock);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe('https://fabrikam.sharepoint.com/sites/cel/_vti_bin/client.svc/ProcessQuery');
    expect(init.headers).toMatchObject({ 'Content-Type': 'text/xml', 'X-RequestDigest': 'digest-1' });
    expect(init.body).toContain('<Actions><A /></Actions><ObjectPaths><B /></ObjectPaths>');
  });

  it('raises CSOM errors that arrive with HTTP 200', async () => {
    const fetchMock = jest.fn().mockResolvedValue(
      json([{ SchemaVersion: '15.0.0.0', ErrorInfo: { ErrorMessage: 'A duplicate content type was found.', ErrorCode: -2146232832, ErrorTypeName: 'Microsoft.SharePoint.SPException' } }])
    );
    await expect(processQuery(spWithDigest().sp, { actions: '', objectPaths: '' }, undefined, fetchMock)).rejects.toMatchObject({
      code: 'CSOM_ERROR',
      message: 'A duplicate content type was found.'
    });
  });

  it('creates a content type with the requested ID and escaped properties', async () => {
    const fetchMock = jest.fn().mockResolvedValue(json([{ SchemaVersion: '15.0.0.0', ErrorInfo: null }, 10, { IsNull: false }, 11, { StringId: '0x0100AB' }]));
    const id = await csomCreateContentType(spWithDigest().sp, { id: '0x0100AB', name: 'Ügyfél & partner', group: 'CopyJet' }, undefined, fetchMock);
    expect(id).toBe('0x0100AB');
    const body = String((fetchMock.mock.calls[0] as [string, RequestInit])[1].body);
    expect(body).toContain('<Property Name="Id" Type="String">0x0100AB</Property>');
    expect(body).toContain('<Property Name="Name" Type="String">Ügyfél &amp; partner</Property>');
    expect(body).toContain('TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}"');
  });

  it('rejects a response without the new ID', async () => {
    const fetchMock = jest.fn().mockResolvedValue(json([{ SchemaVersion: '15.0.0.0', ErrorInfo: null }, 10, { IsNull: true }]));
    await expect(csomCreateContentType(spWithDigest().sp, { id: '0x01', name: 'x' }, undefined, fetchMock)).rejects.toMatchObject({ code: 'CSOM_UNEXPECTED_RESPONSE' });
  });

  it('builds a field link request: add by internal name, flags on new and existing links, Update(false)', () => {
    const body = fieldLinksBody('0x0100AB', [{ internalName: 'CJ_Status', required: true, hidden: false }], [{ id: 'fa564e0f-0c70-4ab9-b863-0177e6ddd247', required: false, hidden: true }]);
    expect(body.objectPaths).toContain('<Method Id="4" ParentId="3" Name="GetById"><Parameters><Parameter Type="String">0x0100AB</Parameter>');
    expect(body.objectPaths).toContain('<Method Id="100" ParentId="6" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">CJ_Status</Parameter>');
    expect(body.objectPaths).toContain(`<Method Id="101" ParentId="5" Name="Add"><Parameters><Parameter TypeId="${CSOM_TYPE.FieldLinkCreationInformation}"><Property Name="Field" ObjectPathId="100" />`);
    expect(body.objectPaths).toContain('<Method Id="102" ParentId="5" Name="GetById"><Parameters><Parameter Type="Guid">{fa564e0f-0c70-4ab9-b863-0177e6ddd247}</Parameter>');
    expect(body.actions).toContain('<SetProperty Id="13" ObjectPathId="101" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty>');
    expect(body.actions).toContain('ObjectPathId="102" Name="Hidden"><Parameter Type="Boolean">true</Parameter>');
    expect(body.actions).toMatch(/<Method Name="Update" Id="\d+" ObjectPathId="4"><Parameters><Parameter Type="Boolean">false<\/Parameter><\/Parameters><\/Method>$/);
  });

  it('skips the request when there is nothing to change', async () => {
    const fetchMock = jest.fn();
    await csomUpdateFieldLinks(spWithDigest().sp, '0x01', [], [], undefined, fetchMock);
    expect(fetchMock).not.toHaveBeenCalled();
  });
});
