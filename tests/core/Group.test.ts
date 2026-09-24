import { GroupExtractor } from '../../src/core/extractors/GroupExtractor';
import { GroupProvider } from '../../src/core/providers/GroupProvider';
import { createInstallContext } from '../../src/core/engine/targetContext';
import { compareGroups, roleKind, roleName, templateGroupTitle, type IGroupInfoLike } from '../../src/core/groups';
import { Logger } from '../../src/core/logger';
import type { IGroup, IInstallContext } from '../../src/core/model';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { validateTemplate } from '../../src/core/schema';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';

interface IRoleDef {
  Id: number;
  Name: string;
  RoleTypeKind: number;
}

interface IWebState {
  web: { Url: string; ServerRelativeUrl: string; Title: string };
  associated: { owner: number; member: number; visitor: number };
  groups: IGroupInfoLike[];
  /** Web role assignments: principal ID → role definition IDs. */
  assignments: { [principalId: number]: number[] };
  roleDefs: IRoleDef[];
}

const HU_ROLES: IRoleDef[] = [
  { Id: 1073741829, Name: 'Teljes hozzáférés', RoleTypeKind: 5 },
  { Id: 1073741828, Name: 'Tervezés', RoleTypeKind: 4 },
  { Id: 1073741830, Name: 'Szerkesztés', RoleTypeKind: 6 },
  { Id: 1073741827, Name: 'Munkatárs', RoleTypeKind: 3 },
  { Id: 1073741826, Name: 'Olvasó', RoleTypeKind: 2 },
  { Id: 1073741825, Name: 'Korlátozott hozzáférés', RoleTypeKind: 1 }
];

/** In-memory web covering the group/security REST calls CopyJet makes, plus the CSOM owner call (spike 07). */
function fakeWeb(state: IWebState): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[]; fetch: (url: string, init: RequestInit) => Promise<Response> } {
  let seq = 100;
  const byId = (id: number): IGroupInfoLike | undefined => state.groups.find((g) => g.Id === id);
  const bindings = (principal: number): Array<{ Name: string; RoleTypeKind: number }> =>
    (state.assignments[principal] || []).map((id) => state.roleDefs.find((d) => d.Id === id)!).map((d) => ({ Name: d.Name, RoleTypeKind: d.RoleTypeKind }));

  const { sp, requests } = createMockSp((req) => {
    let m: RegExpExecArray | null;
    if (req.method === 'POST' && /\/_api\/contextinfo$/i.test(req.url)) return { body: { FormDigestValue: 'digest', WebFullUrl: state.web.Url } };
    if (req.method === 'GET' && /\/_api\/web\?\$select=/i.test(req.url)) return { body: state.web };
    if ((m = /\/associated(Owner|Member|Visitor)Group\?/i.exec(req.url))) {
      return { body: { Id: state.associated[m[1].toLowerCase() as 'owner' | 'member' | 'visitor'] } };
    }
    if (req.method === 'GET' && /\/_api\/web\/roleAssignments\?/i.test(req.url)) {
      return {
        body: Object.keys(state.assignments).map((p) => ({
          PrincipalId: Number(p),
          Member: { PrincipalType: byId(Number(p)) ? 8 : 1 },
          RoleDefinitionBindings: bindings(Number(p))
        }))
      };
    }
    if ((m = /\/roleassignments\/getByPrincipalId\((\d+)\)/i.exec(req.url))) {
      return state.assignments[Number(m[1])] ? { body: { RoleDefinitionBindings: bindings(Number(m[1])) } } : { status: 404, body: {} };
    }
    if (req.method === 'POST' && (m = /\/roleassignments\/addroleassignment\(principalid=(\d+),\s*roledefid=(\d+)\)/i.exec(req.url))) {
      (state.assignments[Number(m[1])] = state.assignments[Number(m[1])] || []).push(Number(m[2]));
      return { body: {} };
    }
    if ((m = /\/roledefinitions\/getByType\((\d+)\)/i.exec(req.url))) {
      const d = state.roleDefs.find((x) => x.RoleTypeKind === Number(m![1]));
      return d ? { body: { Id: d.Id } } : { status: 404, body: {} };
    }
    if ((m = /\/roledefinitions\/getbyname\('([^']+)'\)/i.exec(req.url))) {
      const d = state.roleDefs.find((x) => x.Name === m![1]);
      return d ? { body: { Id: d.Id } } : { status: 404, body: {} };
    }
    if (req.method === 'GET' && (m = /\/siteGroups\?\$filter=Title eq '([^']+)'/i.exec(req.url))) {
      return { body: state.groups.filter((g) => g.Title === m![1]) };
    }
    if (req.method === 'GET' && /\/siteGroups\?/i.test(req.url)) return { body: state.groups };
    if ((m = /\/siteGroups(?:\/getById)?\((\d+)\)/i.exec(req.url))) {
      const g = byId(Number(m[1]))!;
      if (req.method === 'POST') {
        Object.assign(g, req.body);
        return { status: 204 };
      }
      return { body: g };
    }
    if (req.method === 'POST' && /\/siteGroups$/i.test(req.url)) {
      const g: IGroupInfoLike = { Id: ++seq, Owner: { Id: 6, PrincipalType: 1 }, ...(req.body as Partial<IGroupInfoLike> & { Title: string }) };
      state.groups.push(g);
      return { body: g };
    }
    return undefined;
  }, state.web.Url);

  // CSOM: SiteGroups.GetById(group).Owner = SiteGroups.GetById(owner)
  const csomFetch = async (_url: string, init: RequestInit): Promise<Response> => {
    const ids = Array.from(String(init.body).matchAll(/<Parameter Type="Int32">(\d+)<\/Parameter>/g)).map((x) => Number(x[1]));
    byId(ids[0])!.Owner = { Id: ids[1], PrincipalType: 8 };
    return new Response(JSON.stringify([{ SchemaVersion: '15.0.0.0', ErrorInfo: null }]), { status: 200 });
  };
  return { sp, requests, fetch: csomFetch };
}

function sourceState(): IWebState {
  const owners = { Id: 3, PrincipalType: 8 };
  return {
    web: { Url: 'https://contoso.sharepoint.com/sites/Forras', ServerRelativeUrl: '/sites/Forras', Title: 'Forrás' },
    associated: { owner: 3, member: 5, visitor: 4 },
    groups: [
      { Id: 5, Title: 'Forrás Members', AllowMembersEditMembership: true, Owner: owners },
      { Id: 3, Title: 'Forrás Owners', Owner: owners },
      { Id: 10, Title: 'Forrás Projektmenedzserek', Description: 'PM-ek', OnlyAllowMembersViewMembership: true, Owner: { Id: 6, PrincipalType: 1 } },
      { Id: 4, Title: 'Forrás Visitors', Owner: owners },
      { Id: 20, Title: 'SharingLinks.abc.OrganizationView', Owner: owners } // no web role: system group
    ],
    assignments: { 3: [1073741829], 5: [1073741830], 4: [1073741826], 10: [1073741827, 1073741924] },
    roleDefs: HU_ROLES.concat({ Id: 1073741924, Name: 'CopyJet olvasó', RoleTypeKind: 0 })
  };
}

function targetState(extra: Partial<IWebState> = {}): IWebState {
  const owners = { Id: 3, PrincipalType: 8 };
  return {
    web: { Url: 'https://fabrikam.sharepoint.com/sites/Cel', ServerRelativeUrl: '/sites/Cel', Title: 'Cél' },
    associated: { owner: 3, member: 5, visitor: 4 },
    groups: [
      { Id: 3, Title: 'Cél Owners', Owner: owners },
      { Id: 4, Title: 'Cél Visitors', Owner: owners },
      { Id: 5, Title: 'Cél Members', Owner: owners }
    ],
    assignments: { 3: [1073741829], 5: [1073741830], 4: [1073741826] },
    roleDefs: HU_ROLES.slice(),
    ...extra
  };
}

async function extract(): Promise<{ groups: IGroup[]; log: Logger }> {
  const { sp } = fakeWeb(sourceState());
  const extractor = new GroupExtractor();
  const writer = new JsonTemplateWriter(
    createEmptyTemplate({
      name: 'Teszt',
      createdBy: 'anna@contoso.com',
      createdAt: '2026-09-24T10:00:00Z',
      sourceSiteUrl: 'https://contoso.sharepoint.com/sites/Forras',
      sourceTenant: 'contoso.onmicrosoft.com',
      sourceLcid: 1038,
      includesContent: false
    })
  );
  const log = new Logger();
  const tokens = TokenContext.forSite({ absoluteUrl: 'https://contoso.sharepoint.com/sites/Forras', serverRelativeUrl: '/sites/Forras', title: 'Forrás' });
  await extractor.extract(sp, (await extractor.discover(sp)).map((d) => d.ref), { includeContent: false, includeVersions: false, includeMembers: false, tokens, log }, writer);
  expect(validateTemplate(writer.manifest).errors).toEqual([]);
  return { groups: writer.manifest.groups, log };
}

async function installCtx(sp: ReturnType<typeof createMockSp>['sp']): Promise<IInstallContext> {
  return createInstallContext(sp, new Logger());
}

describe('group model', () => {
  it('names permission levels language-independently', () => {
    expect(roleName({ Name: 'Munkatárs', RoleTypeKind: 3 })).toBe('Contribute');
    expect(roleName({ Name: 'Korlátozott hozzáférés', RoleTypeKind: 1 })).toBeUndefined();
    expect(roleName({ Name: 'CopyJet olvasó', RoleTypeKind: 0 })).toBe('CopyJet olvasó');
    expect(roleKind('full control')).toBe(5);
    expect(roleKind('CopyJet olvasó')).toBeUndefined();
  });

  it('turns a site-title prefix into {sitename}', () => {
    expect(templateGroupTitle('Forrás Projektmenedzserek', 'Forrás')).toBe('{sitename} Projektmenedzserek');
    expect(templateGroupTitle('Forrásvédelem', 'Forrás')).toBe('Forrásvédelem');
  });

  it('reports missing roles but not extra ones', () => {
    const def: IGroup = { key: 'A', title: 'A', roles: ['Contribute'], onlyAllowMembersViewMembership: true };
    expect(compareGroups(def, { Id: 1, Title: 'A', OnlyAllowMembersViewMembership: true }, ['Contribute', 'Read'], undefined)).toEqual([]);
    expect(compareGroups(def, { Id: 1, Title: 'A', Owner: { Id: 6, PrincipalType: 1 } }, ['Read'], 3)).toEqual(['onlyAllowMembersViewMembership', 'roles', 'owner']);
  });
});

describe('install context', () => {
  it('registers the site and the associated group tokens', async () => {
    const c = await installCtx(fakeWeb(targetState()).sp);
    expect(c.tokens.get('sitename')).toBe('Cél');
    expect([c.tokens.get('associatedownergroup'), c.tokens.get('associatedmembergroup'), c.tokens.get('associatedvisitorgroup')]).toEqual(['3', '5', '4']);
  });
});

describe('GroupExtractor', () => {
  it('discovers the site’s own groups only (not associated, not role-less system groups)', async () => {
    const found = await new GroupExtractor().discover(fakeWeb(sourceState()).sp);
    expect(found.map((f) => [f.ref.key, f.title])).toEqual([['group:Projektmenedzserek', 'Forrás Projektmenedzserek']]);
  });

  it('extracts a portable definition and warns about the user owner and the custom level', async () => {
    const { groups, log } = await extract();
    expect(groups).toEqual([
      {
        key: 'Projektmenedzserek',
        title: '{sitename} Projektmenedzserek',
        description: 'PM-ek',
        owner: '{associatedownergroup}',
        roles: ['Contribute', 'CopyJet olvasó'],
        includeMembers: false,
        members: [],
        onlyAllowMembersViewMembership: true,
        allowMembersEditMembership: false,
        allowRequestToJoinLeave: false
      }
    ]);
    expect(log.entries.map((e) => e.code)).toEqual(['GROUP_CUSTOM_ROLE', 'GROUP_OWNER_USER']);
  });
});

describe('GroupProvider', () => {
  it('creates the group with the target title, the Owners group as owner and the built-in level by kind', async () => {
    const { groups } = await extract();
    const state = targetState();
    const web = fakeWeb(state);
    const c = await installCtx(web.sp);
    const provider = new GroupProvider(web.fetch);

    expect(await provider.apply(web.sp, groups[0], 'skip', c)).toMatchObject({ outcome: 'created' });
    const g = state.groups.find((x) => x.Title === 'Cél Projektmenedzserek')!;
    expect(g).toMatchObject({ Description: 'PM-ek', OnlyAllowMembersViewMembership: true, Owner: { Id: 3, PrincipalType: 8 } });
    expect(state.assignments[g.Id]).toEqual([1073741827]); // Munkatárs = Contribute
    expect(c.tokens.get('groupkey', 'Projektmenedzserek')).toBe(String(g.Id));
    expect(c.log.entries.map((e) => e.code)).toContain('ROLE_NOT_FOUND'); // custom level missing on target
    expect(c.log.entries.map((e) => e.code)).not.toContain('GROUP_OWNER_NOT_SET');
  });

  it('is idempotent when the custom level exists on the target', async () => {
    const { groups } = await extract();
    const state = targetState({ roleDefs: HU_ROLES.concat({ Id: 1073741999, Name: 'CopyJet olvasó', RoleTypeKind: 0 }) });
    const web = fakeWeb(state);
    const c = await installCtx(web.sp);
    const provider = new GroupProvider(web.fetch);
    await provider.apply(web.sp, groups[0], 'skip', c);
    const g = state.groups.find((x) => x.Title === 'Cél Projektmenedzserek')!;
    expect(state.assignments[g.Id]).toEqual([1073741827, 1073741999]);

    const posts = web.requests.filter((r) => r.method === 'POST').length;
    expect((await provider.diff(web.sp, groups[0], c)).status).toBe('same');
    expect(await provider.apply(web.sp, groups[0], 'update', c)).toMatchObject({ outcome: 'skipped' });
    expect(web.requests.filter((r) => r.method === 'POST').length).toBe(posts);
  });

  it('update mode fixes settings, owner and missing roles without removing roles', async () => {
    const { groups } = await extract();
    const state = targetState();
    state.groups.push({ Id: 50, Title: 'Cél Projektmenedzserek', Description: 'régi', Owner: { Id: 6, PrincipalType: 1 } });
    state.assignments[50] = [1073741826]; // Olvasó
    const web = fakeWeb(state);
    const c = await installCtx(web.sp);
    const provider = new GroupProvider(web.fetch);

    expect(await provider.diff(web.sp, groups[0], c)).toMatchObject({ status: 'different', changes: ['description', 'onlyAllowMembersViewMembership', 'roles', 'owner'] });
    expect(await provider.apply(web.sp, groups[0], 'skip', c)).toMatchObject({ outcome: 'skipped' });
    expect(state.groups.find((x) => x.Id === 50)!.Description).toBe('régi');
    expect(await provider.apply(web.sp, groups[0], 'update', c)).toMatchObject({ outcome: 'updated' });
    expect(state.groups.find((x) => x.Id === 50)).toMatchObject({ Description: 'PM-ek', OnlyAllowMembersViewMembership: true, Owner: { Id: 3 } });
    expect(state.assignments[50]).toEqual([1073741826, 1073741827]);
  });

  it('rename mode creates a "(copy)" group and points the key token at it', async () => {
    const { groups } = await extract();
    const state = targetState();
    state.groups.push({ Id: 50, Title: 'Cél Projektmenedzserek', Owner: { Id: 6, PrincipalType: 1 } });
    const web = fakeWeb(state);
    const c = await installCtx(web.sp);
    expect(await new GroupProvider(web.fetch).apply(web.sp, groups[0], 'rename', c)).toMatchObject({ outcome: 'created' });
    const copy = state.groups.find((x) => x.Title === 'Cél Projektmenedzserek (copy)')!;
    expect(c.tokens.get('groupkey', 'Projektmenedzserek')).toBe(String(copy.Id));
  });
});
