import { Logger } from '../../src/core/logger';
import { PrincipalMapper, parseMappingCsv, replaceDomain } from '../../src/core/mapping';
import type { IPrincipal } from '../../src/core/model';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp } from '../helpers/mockSp';

describe('parseMappingCsv', () => {
  it('reads ";", "," and tab separated tables, with or without header, quoted values', () => {
    expect(parseMappingCsv('forrás;cél\nanna@contoso.com;anna@fabrikam.com\r\n\nEva.Szabo@contoso.com;e.szabo@fabrikam.com\n')).toEqual({
      rows: { 'anna@contoso.com': 'anna@fabrikam.com', 'eva.szabo@contoso.com': 'e.szabo@fabrikam.com' },
      invalid: []
    });
    expect(parseMappingCsv('﻿"anna@contoso.com","anna@fabrikam.com"').rows).toEqual({ 'anna@contoso.com': 'anna@fabrikam.com' });
    expect(parseMappingCsv('a@contoso.com\tb@fabrikam.com').rows).toEqual({ 'a@contoso.com': 'b@fabrikam.com' });
  });

  it('reports lines that are not source;target', () => {
    expect(parseMappingCsv('source;target\nanna@contoso.com\nkiss anna;anna@fabrikam.com\npeter@contoso.com;peter@fabrikam.com').invalid).toEqual([2, 3]);
  });
});

describe('replaceDomain', () => {
  const rules = [{ from: 'contoso.com', to: 'fabrikam.com' }];
  it('replaces the domain of e-mails and claims logins, case-insensitively', () => {
    expect(replaceDomain('Anna@Contoso.com', rules)).toBe('Anna@fabrikam.com');
    expect(replaceDomain('i:0#.f|membership|anna@contoso.com', [{ from: '@contoso.com', to: '@fabrikam.com' }])).toBe('i:0#.f|membership|anna@fabrikam.com');
    expect(replaceDomain('anna@sub.contoso.com', rules)).toBeUndefined();
    expect(replaceDomain('anna@other.com', rules)).toBeUndefined();
  });
});

describe('PrincipalMapper', () => {
  const principals: IPrincipal[] = [
    { key: 'anna', kind: 'user', loginName: 'i:0#.f|membership|anna@contoso.com', email: 'anna@contoso.com', displayName: 'Kiss Anna' },
    { key: 'peter', kind: 'user', loginName: 'i:0#.f|membership|peter.nagy@contoso.com', email: 'peter.nagy@contoso.com', displayName: 'Nagy Péter' },
    { key: 'eva', kind: 'user', loginName: 'i:0#.f|membership|eva.szabo@contoso.com', email: 'eva.szabo@contoso.com', displayName: 'Szabó Éva' },
    { key: 'gabor', kind: 'user', loginName: 'i:0#.f|membership|gabor.toth@contoso.com', email: 'gabor.toth@contoso.com', displayName: 'Tóth Gábor' }
  ];
  /** The target tenant (fabrikam) knows these accounts; ensureUser answers with their login. */
  const known = ['anna@fabrikam.com', 'peter.nagy@fabrikam.com', 'e.szabo@fabrikam.com', 'helpdesk@fabrikam.com'];
  const target = (): { sp: ReturnType<typeof createMockSp>['sp']; asked: string[] } => {
    const asked: string[] = [];
    const { sp } = createMockSp((req) => {
      if (req.method === 'POST' && /\/_api\/web\/ensureuser$/i.test(req.url)) {
        const value = (req.body as { logonName: string }).logonName;
        asked.push(value);
        const email = value.split('|').pop()!.toLowerCase();
        return known.indexOf(email) >= 0 ? { body: { Id: 1, LoginName: `i:0#.f|membership|${email}` } } : { status: 500, body: { 'odata.error': { message: { value: 'not found' } } } };
      }
      return undefined;
    });
    return { sp, asked };
  };

  it('maps by the mockup: domain replacement, CSV, nothing for the rest (warned once per log)', async () => {
    const { sp } = target();
    const mapper = new PrincipalMapper(sp, principals);
    mapper.setRules({ domains: [{ from: 'contoso.com', to: 'fabrikam.com' }], csv: { 'eva.szabo@contoso.com': 'e.szabo@fabrikam.com' } });
    const tokens = new TokenContext();
    const log = new Logger();
    const result = await mapper.map(['anna', 'peter', 'eva', 'gabor'], tokens, log);
    expect(result.map((m) => [m.key, m.strategy, m.login])).toEqual([
      ['anna', 'domain', 'i:0#.f|membership|anna@fabrikam.com'],
      ['peter', 'domain', 'i:0#.f|membership|peter.nagy@fabrikam.com'],
      ['eva', 'csv', 'i:0#.f|membership|e.szabo@fabrikam.com'],
      ['gabor', undefined, undefined]
    ]);
    expect(tokens.get('principal', 'eva')).toBe('i:0#.f|membership|e.szabo@fabrikam.com');
    await mapper.map(['gabor'], tokens, log);
    expect(log.entries.map((e) => e.code)).toEqual(['PRINCIPAL_NOT_FOUND']);
  });

  it('lets a manual choice win, uses the fallback user for the rest, and asks SharePoint once per value', async () => {
    const { sp, asked } = target();
    const mapper = new PrincipalMapper(sp, principals);
    mapper.setRules({ manual: { gabor: 'e.szabo@fabrikam.com' }, fallback: 'helpdesk@fabrikam.com' });
    const log = new Logger();
    const result = await mapper.map(['anna', 'gabor'], new TokenContext(), log);
    expect(result.map((m) => [m.key, m.strategy, m.login])).toEqual([
      ['anna', 'fallback', 'i:0#.f|membership|helpdesk@fabrikam.com'],
      ['gabor', 'manual', 'i:0#.f|membership|e.szabo@fabrikam.com']
    ]);
    expect(log.entries.map((e) => e.code)).toEqual(['PRINCIPAL_FALLBACK']);
    // New rules re-map, but values already asked are not sent again.
    const before = asked.length;
    mapper.setRules({ domains: [{ from: 'contoso.com', to: 'fabrikam.com' }], fallback: 'helpdesk@fabrikam.com' });
    expect((await mapper.map(['anna'], new TokenContext(), new Logger()))[0].strategy).toBe('domain');
    expect(asked.slice(before)).toEqual(['anna@fabrikam.com']);
  });
});
