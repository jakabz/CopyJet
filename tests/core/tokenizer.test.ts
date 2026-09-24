import {
  TokenContext,
  deepResolve,
  deepTokenize,
  getTokenDefinition,
  registerToken,
  resolve,
  tokenize
} from '../../src/core/tokenizer';
import { CopyJetError } from '../../src/core/errors';

const LIST_ID = '0f8b4c2e-1a2b-4c3d-8e9f-a0b1c2d3e4f5';
const FIELD_ID = '8C3A1D52-3B4E-4F7A-9C21-5D6E7F8A9B01';

function source(): TokenContext {
  return TokenContext.forSite({
    absoluteUrl: 'https://contoso.sharepoint.com/sites/projekt',
    serverRelativeUrl: '/sites/projekt',
    title: 'Projekt'
  })
    .set('listkey', 'Ugyfelek', LIST_ID)
    .set('fieldid', 'CJ_Status', `{${FIELD_ID}}`);
}

function target(): TokenContext {
  return TokenContext.forSite({
    absoluteUrl: 'https://fabrikam.sharepoint.com/sites/uj/',
    serverRelativeUrl: '/sites/uj',
    title: 'Új projekt'
  })
    .set('listkey', 'Ugyfelek', '11111111-2222-4333-8444-555555555555')
    .set('fieldid', 'CJ_Status', '99999999-8888-4777-8666-555555555555');
}

describe('tokenize', () => {
  it('prefers the absolute URL over the server-relative one (longest match)', () => {
    expect(tokenize('https://contoso.sharepoint.com/sites/projekt/Lists/Projektek', source())).toBe('{site}/Lists/Projektek');
    expect(tokenize('<a href="/sites/projekt/SitePages/Home.aspx">', source())).toBe('<a href="{siterelative}/SitePages/Home.aspx">');
  });

  it('respects URL boundaries', () => {
    const ctx = source();
    expect(tokenize('/sites/projekt2/x', ctx)).toBe('/sites/projekt2/x');
    expect(tokenize('https://other.com/sites/projekt', ctx)).toBe('https://other.com/sites/projekt');
    expect(tokenize('HTTPS://CONTOSO.sharepoint.com/sites/Projekt', ctx)).toBe('{site}');
  });

  it('replaces GUIDs in any case, keeping braces and URL-encoded braces', () => {
    const ctx = source();
    expect(tokenize(`List="{${LIST_ID.toUpperCase()}}"`, ctx)).toBe('List="{{listkey:Ugyfelek}}"');
    expect(tokenize(`List=%7B${LIST_ID}%7D`, ctx)).toBe('List=%7B{listkey:Ugyfelek}%7D');
    expect(tokenize(`<FieldRef ID="${FIELD_ID.toLowerCase()}" />`, ctx)).toBe('<FieldRef ID="{fieldid:CJ_Status}" />');
    expect(tokenize(`x${LIST_ID}0`, ctx)).toBe(`x${LIST_ID}0`);
  });

  it('never replaces the site title or a root site\'s "/"', () => {
    const root = TokenContext.forSite({ absoluteUrl: 'https://contoso.sharepoint.com', serverRelativeUrl: '/', title: 'Projekt' });
    expect(tokenize('Projekt / a/b', root)).toBe('Projekt / a/b');
  });
});

describe('resolve', () => {
  it('round-trips source → template → target', () => {
    const xml = `<Field ID="{${FIELD_ID}}" List="{${LIST_ID}}" Url="https://contoso.sharepoint.com/sites/projekt/Lists/Ugyfelek" />`;
    const tpl = tokenize(xml, source());
    expect(tpl).toBe('<Field ID="{{fieldid:CJ_Status}}" List="{{listkey:Ugyfelek}}" Url="{site}/Lists/Ugyfelek" />');
    expect(resolve(tpl, target())).toBe(
      '<Field ID="{99999999-8888-4777-8666-555555555555}" List="{11111111-2222-4333-8444-555555555555}" Url="https://fabrikam.sharepoint.com/sites/uj/Lists/Ugyfelek" />'
    );
  });

  it('resolves {sitename} and leaves GUID braces and JSON alone', () => {
    const ctx = target();
    expect(resolve('{sitename} Projektmenedzserek', ctx)).toBe('Új projekt Projektmenedzserek');
    expect(resolve('{deadbeef-0000-4000-8000-000000000000} {"a":1} {Title}', ctx)).toBe('{deadbeef-0000-4000-8000-000000000000} {"a":1} {Title}');
  });

  it.each([
    ['{nosuchtoken}', 'TOKEN_UNKNOWN'],
    ['{listkey:Missing}', 'TOKEN_UNRESOLVED'],
    ['{listkey}', 'TOKEN_INVALID'],
    ['{site:x}', 'TOKEN_INVALID']
  ])('%s throws %s', (value, code) => {
    expect(() => resolve(value, target())).toThrow(CopyJetError);
    expect(() => resolve(value, target())).toThrow(expect.objectContaining({ code }));
  });
});

describe('deep tokenize / resolve', () => {
  it('walks nested objects and arrays, leaving keys and non-strings alone', () => {
    const props = {
      title: 'Lista',
      listId: LIST_ID,
      nested: [{ url: 'https://contoso.sharepoint.com/sites/projekt/Shared Documents' }, 5, null, true]
    };
    const tpl = deepTokenize(props, source());
    expect(tpl).toEqual({
      title: 'Lista',
      listId: '{listkey:Ugyfelek}',
      nested: [{ url: '{site}/Shared Documents' }, 5, null, true]
    });
    expect(props.listId).toBe(LIST_ID);
    expect(deepResolve(tpl, target()).nested[0]).toEqual({ url: 'https://fabrikam.sharepoint.com/sites/uj/Shared Documents' });
  });
});

describe('token registry', () => {
  it('accepts new tokens with custom resolvers and rejects bad names', () => {
    registerToken({ name: 'testconst', hasArgument: true, kind: 'text', autoTokenize: false, resolve: (arg) => `const-${arg}` });
    expect(getTokenDefinition('testconst')).toBeDefined();
    expect(resolve('{testconst:a}', new TokenContext())).toBe('const-a');
    expect(() => registerToken({ name: 'Bad-Name', hasArgument: false, kind: 'text', autoTokenize: false })).toThrow();
  });

  it('validates names and arguments when setting values', () => {
    const ctx = new TokenContext();
    expect(() => ctx.set('nosuch', undefined, 'x')).toThrow(expect.objectContaining({ code: 'TOKEN_UNKNOWN' }));
    expect(() => ctx.set('listkey', undefined, 'x')).toThrow(expect.objectContaining({ code: 'TOKEN_INVALID' }));
    expect(() => ctx.set('listkey', 'a:b', 'x')).toThrow(expect.objectContaining({ code: 'TOKEN_INVALID' }));
  });
});
