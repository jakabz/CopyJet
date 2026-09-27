import {
  LocalTimeConverter,
  PrincipalCollector,
  principalKind,
  toItemField,
  toTargetValue,
  toTemplateValue,
  type IItemField,
  type IWebLocale
} from '../../src/core/items';
import type { FieldValue, IPrincipal } from '../../src/core/model';
import { TokenContext } from '../../src/core/tokenizer';

const EN: IWebLocale = { localeId: 1033, tag: 'en-US', decimalSeparator: '.', time24: false };
const HU: IWebLocale = { localeId: 1038, tag: 'hu-HU', decimalSeparator: ',', time24: true };

const field = (internalName: string, TypeAsString: string): IItemField => toItemField({ InternalName: internalName, TypeAsString })!;

/** SharePoint's utcToLocalTime for Budapest, computed with Intl (the real call goes to the target web). */
function budapest(iso: string): string {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Europe/Budapest',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23'
  }).formatToParts(new Date(iso));
  const p = (t: string): string => parts.filter((x) => x.type === t)[0].value;
  return `${p('year')}-${p('month')}-${p('day')}T${p('hour')}:${p('minute')}:${p('second')}`;
}

describe('toItemField', () => {
  it('copies writable content columns and skips system, hidden, read-only and unknown ones', () => {
    expect(toItemField({ InternalName: 'CJNum', TypeAsString: 'Currency' })!.kind).toBe('number');
    expect(toItemField({ InternalName: 'Valaki', TypeAsString: 'UserMulti' })!.kind).toBe('userMulti');
    expect(toItemField({ InternalName: 'Author', TypeAsString: 'User' })).toBeUndefined();
    expect(toItemField({ InternalName: 'ContentType', TypeAsString: 'Computed' })).toBeUndefined();
    expect(toItemField({ InternalName: 'X', TypeAsString: 'Text', Hidden: true })).toBeUndefined();
    expect(toItemField({ InternalName: 'X', TypeAsString: 'Text', ReadOnlyField: true })).toBeUndefined();
    expect(toItemField({ InternalName: 'Szamitott', TypeAsString: 'Calculated' })).toBeUndefined();
    expect(toItemField({ InternalName: 'Terulet', TypeAsString: 'TaxonomyFieldType' })).toBeUndefined();
  });
});

describe('toTemplateValue (REST values, spike 08 A)', () => {
  const tokens = TokenContext.forSite({ absoluteUrl: 'https://contoso.sharepoint.com/sites/Forras', serverRelativeUrl: '/sites/Forras', title: 'Forrás' });
  const ctx = { tokens, principal: (id: number) => (id === 7 ? '{principal:anna}' : undefined) };

  it.each<[string, string, unknown, FieldValue | undefined]>([
    ['Text', 'Title', 'ERP bevezetés', 'ERP bevezetés'],
    ['Text', 'Title', '', undefined],
    ['Number', 'CJNum', 11.5, 11.5],
    ['Boolean', 'CJBool', false, false],
    ['Choice', 'CJChoice', 'Kettő', 'Kettő'],
    ['MultiChoice', 'CJMulti', ['A', 'C'], ['A', 'C']],
    ['MultiChoice', 'CJMulti', { results: ['B'] }, ['B']],
    ['MultiChoice', 'CJMulti', [], undefined],
    ['DateTime', 'CJDate', '2026-03-02T08:15:00Z', '2026-03-02T08:15:00Z'],
    ['User', 'Valaki', 7, { principals: ['{principal:anna}'] }],
    ['UserMulti', 'Valakik', [7, 99], { principals: ['{principal:anna}'] }],
    ['User', 'Valaki', 99, undefined],
    ['Lookup', 'Ugyfel', 3, { lookup: [3] }],
    ['LookupMulti', 'Ugyfelek', [1, 2], { lookup: [1, 2] }],
    ['Lookup', 'Ugyfel', null, undefined],
    ['URL', 'CJUrl', { Url: 'https://example.com', Description: 'Példa' }, { url: 'https://example.com', description: 'Példa' }],
    ['URL', 'CJUrl', { Url: 'https://example.com', Description: 'https://example.com' }, { url: 'https://example.com' }],
    ['URL', 'CJUrl', { Url: 'https://contoso.sharepoint.com/sites/Forras/Lists/X', Description: 'X' }, { url: '{site}/Lists/X', description: 'X' }],
    ['Note', 'CJNote', '<a href="https://contoso.sharepoint.com/sites/Forras/SitePages/A.aspx">A</a>', '<a href="{site}/SitePages/A.aspx">A</a>']
  ])('%s %s', (type, name, raw, expected) => {
    expect(toTemplateValue(field(name, type), raw, ctx)).toEqual(expected);
  });
});

describe('toTargetValue (AddValidateUpdateItemUsingPath strings, spike 08 C–E)', () => {
  const tokens = TokenContext.forSite({ absoluteUrl: 'https://fabrikam.sharepoint.com/sites/Cel', serverRelativeUrl: '/sites/Cel', title: 'Cél' });
  tokens.set('principal', 'anna', 'i:0#.f|membership|anna@fabrikam.com');
  const times = { localTime: budapest };

  it.each<[string, FieldValue, IWebLocale, string | undefined]>([
    ['Number', 12.5, EN, '12.5'],
    ['Number', 12.5, HU, '12,5'],
    ['Number', -1234.75, HU, '-1234,75'],
    ['DateTime', '2026-03-02T08:15:00Z', EN, '3/2/2026 9:15 AM'],
    ['DateTime', '2026-07-16T11:30:00Z', HU, '2026. 07. 16. 13:30'],
    ['DateTime', '2026-12-31T23:05:00Z', EN, '1/1/2027 12:05 AM'],
    ['Boolean', true, EN, '1'],
    ['Boolean', false, EN, '0'],
    ['MultiChoice', ['A', 'C'], EN, ';#A;#C;#'],
    ['Choice', 'Egy', HU, 'Egy'],
    ['URL', { url: '{site}/a,b?x=1,2', description: 'Leírás, vesszővel' }, EN, 'https://fabrikam.sharepoint.com/sites/Cel/a,,b?x=1,,2, Leírás, vesszővel'],
    ['URL', { url: 'https://example.com' }, EN, 'https://example.com, https://example.com'],
    ['User', { principals: ['{principal:anna}'] }, EN, '[{"Key":"i:0#.f|membership|anna@fabrikam.com"}]'],
    ['UserMulti', { principals: ['{principal:anna}', '{principal:nincs}'] }, EN, '[{"Key":"i:0#.f|membership|anna@fabrikam.com"}]'],
    ['User', { principals: ['{principal:nincs}'] }, EN, undefined],
    ['Note', 'Lásd {site}/SitePages és {nem token}', EN, 'Lásd https://fabrikam.sharepoint.com/sites/Cel/SitePages és {nem token}'],
    ['Text', null, EN, undefined]
  ])('%s %j (%s)', (type, value, locale, expected) => {
    expect(toTargetValue(field('F', type), value, { locale: locale, tokens, ...times })).toBe(expected);
  });

  it('maps lookups through the ID map; multi lookups as "a;#;#b"', () => {
    const lookupIds = (_f: IItemField, ids: number[]): number[] => ids.map((id) => ({ 1: 11, 2: 12 } as { [k: number]: number })[id]).filter((x) => !!x);
    const ctx = { locale: EN, tokens, localTime: budapest, lookupIds };
    expect(toTargetValue(field('L', 'Lookup'), { lookup: [2] }, ctx)).toBe('12');
    expect(toTargetValue(field('L', 'LookupMulti'), { lookup: [1, 2, 3] }, ctx)).toBe('11;#;#12');
    expect(toTargetValue(field('L', 'Lookup'), { lookup: [3] }, ctx)).toBeUndefined();
    expect(toTargetValue(field('L', 'Lookup'), { lookup: [1] }, { locale: EN, tokens, localTime: budapest })).toBeUndefined();
  });
});

describe('LocalTimeConverter', () => {
  it('asks SharePoint once per UTC day and matches its answers, across daylight saving changes', async () => {
    const calls: string[] = [];
    const conv = new LocalTimeConverter(async (iso) => {
      calls.push(iso);
      return budapest(iso);
    });
    // 2026-03-29 01:00Z: Budapest moves from +1 to +2. 2026-10-25 01:00Z: back to +1.
    const values = [
      '2026-03-02T08:15:00Z',
      '2026-03-02T21:00:00Z',
      '2026-03-29T00:30:00Z',
      '2026-03-29T01:30:00Z',
      '2026-07-16T11:30:00Z',
      '2026-10-25T00:59:00Z',
      '2026-10-25T01:00:00Z',
      '2026-12-31T23:05:00Z'
    ];
    await conv.prepare(values);
    values.forEach((v) => expect(conv.local(v)).toBe(budapest(v)));
    // 5 distinct days → 10 day boundaries; the 2 transition days add 4 exact calls.
    expect(calls.length).toBe(14);
    await conv.prepare(['2026-03-02T10:00:00Z']);
    expect(calls.length).toBe(14);
    expect(conv.local('2026-03-02T10:00:00Z')).toBe('2026-03-02T11:00:00');
  });

  it('refuses values it was not prepared for', () => {
    expect(() => new LocalTimeConverter(async () => '').local('2026-01-01T00:00:00Z')).toThrow(/not prepared/);
  });
});

describe('principals', () => {
  it('knows the kinds CopyJet carries', () => {
    expect(principalKind('i:0#.f|membership|anna@contoso.com')).toBe('user');
    expect(principalKind('c:0t.c|tenant|3f1c…')).toBe('securityGroup');
    expect(principalKind('c:0o.c|federateddirectoryclaimprovider|a1b2')).toBe('m365Group');
    expect(principalKind('c:0-.f|rolemanager|spo-grid-all-users/abc')).toBe('everyone');
    expect(principalKind('SHAREPOINT\\system')).toBeUndefined();
    expect(principalKind('i:0#.f|membership|urn%3aspo%3aguest#x')).toBeUndefined();
  });

  it('collects each person once with a unique key and hands out tokens', () => {
    const principals: IPrincipal[] = [{ key: 'anna', kind: 'user', loginName: 'i:0#.f|membership|anna@other.com' }];
    const c = new PrincipalCollector(principals, [
      { Id: 7, LoginName: 'i:0#.f|membership|anna@contoso.com', Title: 'Kiss Anna', Email: 'anna@contoso.com' },
      { Id: 8, LoginName: 'i:0#.f|membership|peter@contoso.com', Title: 'Nagy Péter', Email: 'peter@contoso.com', UserPrincipalName: 'peter@contoso.com' },
      { Id: 9, LoginName: 'SHAREPOINT\\system', Title: 'Rendszerfiók' }
    ]);
    expect(c.token(7)).toBe('{principal:anna_2}');
    expect(c.token(7)).toBe('{principal:anna_2}');
    expect(c.token(8)).toBe('{principal:peter}');
    expect(c.token(9)).toBeUndefined();
    expect(c.token(42)).toBeUndefined();
    expect(principals.map((p) => p.key)).toEqual(['anna', 'anna_2', 'peter']);
    expect(principals[2]).toEqual({ key: 'peter', kind: 'user', loginName: 'i:0#.f|membership|peter@contoso.com', email: 'peter@contoso.com', upn: 'peter@contoso.com', displayName: 'Nagy Péter' });
  });
});
