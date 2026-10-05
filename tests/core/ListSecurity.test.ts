import { ListSecurityExtractor } from '../../src/core/extractors/ListSecurityExtractor';
import { toListDef } from '../../src/core/lists';
import { Logger } from '../../src/core/logger';
import type { IGroup, IInstallContext, ISecurity } from '../../src/core/model';
import { JsonTemplateWriter, createEmptyTemplate } from '../../src/core/packager';
import { buildPlan } from '../../src/core/planner';
import { ListSecurityProvider } from '../../src/core/providers/ListSecurityProvider';
import type { IListSecurityDef, IRoleAssignmentLike } from '../../src/core/security';
import { TokenContext } from '../../src/core/tokenizer';
import { createMockSp, type IMockRequest } from '../helpers/mockSp';
import { SOURCE_WEB, sourceLists, tesztLista } from '../fixtures/lists';

const ADELE = 'i:0#.f|membership|adelev@contoso.com';
const ADMINS = 'c:0t.c|tenant|22222222-0000-4000-8000-000000000001';
const PORTAL = 'c:0o.c|federateddirectoryclaimprovider|99999999-0000-4000-8000-000000000001';
const TEST_LIST = tesztLista.RootFolder.ServerRelativeUrl;

const binding = (Name: string, RoleTypeKind: number) => ({ Name, RoleTypeKind });
const member = (Id: number, Title: string, PrincipalType: number, LoginName = Title) => ({ Id, Title, PrincipalType, LoginName });

/** The "Teszt lista" assignments of spike 15 A. */
const sourceAssignments: IRoleAssignmentLike[] = [
  { Member: member(3, 'Forrás Owners', 8), RoleDefinitionBindings: [binding('Teljes hozzáférés', 5)] },
  { Member: member(5, 'Forrás Members', 8), RoleDefinitionBindings: [binding('Szerkesztés', 6)] },
  { Member: member(10, 'Forrás Projektmenedzserek', 8), RoleDefinitionBindings: [binding('Szerkesztés', 6), binding('Munkatárs', 3), binding('CopyJet olvasó', 0)] },
  { Member: member(11, 'Adele Vance', 1, ADELE), RoleDefinitionBindings: [binding('Olvasó', 2)] },
  { Member: member(13, 'KerVezAdmins', 4, ADMINS), RoleDefinitionBindings: [binding('Korlátozott nézet', 8)] },
  { Member: member(14, 'KerVezPortal Members', 4, PORTAL), RoleDefinitionBindings: [binding('Korlátozott hozzáférés', 1)] },
  { Member: member(20, 'Csak listán', 8), RoleDefinitionBindings: [binding('Olvasó', 2)] }
];

function sourceSp(): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  return createMockSp((req) => {
    const u = req.url;
    if (/\/_api\/web\?\$select=Url,ServerRelativeUrl,Title$/.test(u)) return { body: SOURCE_WEB };
    if (/\/_api\/web\?\$select=Title$/.test(u)) return { body: { Title: SOURCE_WEB.Title } };
    if (/\/_api\/web\/lists\?\$select=Id,HasUniqueRoleAssignments&\$filter=Hidden eq false$/.test(u)) {
      return { body: sourceLists.map((l) => ({ Id: l.Id, HasUniqueRoleAssignments: l.Id === tesztLista.Id })) };
    }
    if (/\/_api\/web\/lists\?\$select=/.test(u)) return { body: sourceLists };
    if (/associatedOwnerGroup/i.test(u)) return { body: { Id: 3 } };
    if (/associatedMemberGroup/i.test(u)) return { body: { Id: 5 } };
    if (/associatedVisitorGroup/i.test(u)) return { body: { Id: 4 } };
    if (/\/_api\/web\/siteGroups\?/i.test(u)) {
      return { body: [3, 4, 5, 10, 20].map((Id) => ({ Id, Title: Id === 10 ? 'Forrás Projektmenedzserek' : Id === 20 ? 'Csak listán' : `Group ${Id}`, Owner: { Id: 3, PrincipalType: 8 } })) };
    }
    if (/\/_api\/web\/roleAssignments\?/i.test(u)) {
      // Group 20 has no web role: CopyJet does not copy it.
      return { body: [3, 4, 5, 10].map((PrincipalId) => ({ PrincipalId, Member: { PrincipalType: 8 }, RoleDefinitionBindings: [binding('Szerkesztés', 6)] })) };
    }
    if (/\/_api\/web\/siteUsers\?/i.test(u)) {
      return {
        body: [
          { Id: 11, LoginName: ADELE, Title: 'Adele Vance', Email: 'adelev@contoso.com', PrincipalType: 1 },
          { Id: 13, LoginName: ADMINS, Title: 'KerVezAdmins', PrincipalType: 4 },
          { Id: 14, LoginName: PORTAL, Title: 'KerVezPortal Members', PrincipalType: 4 }
        ]
      };
    }
    if (new RegExp(`getList\\('${TEST_LIST}'\\)/roleAssignments\\?`, 'i').test(u)) return { body: sourceAssignments };
    return undefined;
  }, SOURCE_WEB.Url);
}

const PM: IGroup = { key: 'Projektmenedzserek', title: '{sitename} Projektmenedzserek', owner: '{associatedownergroup}', roles: ['Edit'], includeMembers: false, members: [] };

async function extract(log = new Logger()): Promise<JsonTemplateWriter> {
  const writer = new JsonTemplateWriter(
    createEmptyTemplate({ name: 'Sec', createdBy: 'a', createdAt: '2026-10-05T10:00:00Z', sourceSiteUrl: SOURCE_WEB.Url, sourceTenant: 'contoso.sharepoint.com', sourceLcid: 1038, includesContent: false })
  );
  writer.manifest.lists.push(toListDef(tesztLista, 'Teszt_lista', SOURCE_WEB.ServerRelativeUrl));
  await new ListSecurityExtractor().extract(
    sourceSp().sp,
    [{ kind: 'listSecurity', key: 'listSecurity:Teszt_lista' }],
    { includeContent: false, includeVersions: false, includeMembers: false, tokens: new TokenContext(), log },
    writer
  );
  return writer;
}

describe('ListSecurityExtractor', () => {
  it('discovers only lists with unique permissions, under their list, counting real assignments', async () => {
    const found = await new ListSecurityExtractor().discover(sourceSp().sp);
    expect(found.map((a) => [a.ref.key, a.parentKey, a.itemCount])).toEqual([['listSecurity:Teszt_lista', 'list:Teszt_lista', 6]]);
  });

  it('carries the assignments as tokens and template role names, without Limited Access', async () => {
    const log = new Logger();
    const writer = await extract(log);
    expect(writer.manifest.lists[0].security).toEqual({
      breakInheritance: true,
      copyRoleAssignments: false,
      roleAssignments: [
        { principal: '{associatedmembergroup}', roles: ['Edit'] },
        { principal: '{associatedownergroup}', roles: ['Full Control'] },
        { principal: '{groupkey:Projektmenedzserek}', roles: ['Edit', 'Contribute', 'CopyJet olvasó'] },
        { principal: '{principal:KerVezAdmins}', roles: ['Restricted View'] },
        { principal: '{principal:adelev}', roles: ['Read'] }
      ]
    });
    expect(writer.manifest.principals.map((p) => [p.key, p.kind])).toEqual([
      ['adelev', 'user'],
      ['KerVezAdmins', 'securityGroup']
    ]);
    expect(log.entries.filter((e) => e.level === 'warn').map((e) => [e.code, e.detail])).toEqual([['LIST_SECURITY_GROUP_OUTSIDE_TEMPLATE', 'Csak listán']]);
  });

  it('puts the step after the list and the group it names', async () => {
    const writer = await extract();
    writer.manifest.groups.push(PM);
    const plan = buildPlan(writer.manifest);
    const step = plan.steps.filter((s) => s.ref.key === 'listSecurity:Teszt_lista')[0];
    expect(step.dependsOn.sort()).toEqual(['group:Projektmenedzserek', 'list:Teszt_lista']);
    expect(step.lock).toBe('list:Teszt_lista');
  });
});

// ---------------------------------------------------------------- provider

const TARGET = 'https://fabrikam.sharepoint.com/sites/Cel';
const TARGET_LIST = '/sites/Cel/Lists/Teszt lista';
const ADELE_T = 'i:0#.f|membership|adelev@fabrikam.com';

const security: ISecurity = {
  breakInheritance: true,
  copyRoleAssignments: false,
  roleAssignments: [
    { principal: '{associatedownergroup}', roles: ['Full Control'] },
    { principal: '{groupkey:Projektmenedzserek}', roles: ['Edit', 'CopyJet olvasó'] },
    { principal: '{principal:adelev}', roles: ['Read'] },
    { principal: '{principal:nobody}', roles: ['Read'] }
  ]
};
const def: IListSecurityDef = { listKey: 'Teszt_lista', listUrl: 'Lists/Teszt lista', security };

interface ITargetState {
  unique: boolean;
  /** principal ID → role definition IDs */
  assignments: { [id: number]: number[] };
  siteAdmin: boolean;
}

const ROLE_IDS: { [kind: number]: number } = { 2: 1073741826, 5: 1073741829, 6: 1073741830 };
const LOGIN_BY_ID: { [id: number]: string } = { 3: 'Cél Owners', 6: 'i:0#.f|membership|installer@fabrikam.com', 12: 'Cél Projektmenedzserek', 31: ADELE_T };

function targetSp(state: ITargetState): { sp: ReturnType<typeof createMockSp>['sp']; requests: IMockRequest[] } {
  return createMockSp((req) => {
    const u = req.url;
    let m: RegExpExecArray | null;
    const L = `getList('${TARGET_LIST}')`;
    if (u.indexOf(L) >= 0) {
      const rest = u.slice(u.indexOf(L) + L.length);
      if (rest === '?$select=HasUniqueRoleAssignments') return { body: { HasUniqueRoleAssignments: state.unique } };
      if (/^\/roleAssignments\?/i.test(rest)) {
        const names: { [id: number]: [string, number] } = { 1073741826: ['Olvasás', 2], 1073741829: ['Teljes hozzáférés', 5], 1073741830: ['Szerkesztés', 6], 1073741900: ['CopyJet olvasó', 0] };
        return {
          body: Object.keys(state.assignments).map((id) => ({
            Member: { Id: Number(id), Title: LOGIN_BY_ID[Number(id)], LoginName: LOGIN_BY_ID[Number(id)], PrincipalType: Number(id) === 31 ? 1 : 8 },
            RoleDefinitionBindings: state.assignments[Number(id)].map((r) => ({ Name: names[r][0], RoleTypeKind: names[r][1] }))
          }))
        };
      }
      if ((m = /^\/breakRoleInheritance\(copyRoleAssignments=false, ?clearSubscopes=false\)$/i.exec(rest))) {
        state.unique = true;
        state.assignments = { 6: [ROLE_IDS[5]] };
        return { status: 204 };
      }
      if ((m = /^\/roleAssignments\/addRoleAssignment\(principalId=(\d+), ?roleDefId=(\d+)\)$/i.exec(rest))) {
        const list = (state.assignments[Number(m[1])] = state.assignments[Number(m[1])] || []);
        if (list.indexOf(Number(m[2])) < 0) list.push(Number(m[2]));
        return { status: 204 };
      }
      if ((m = /^\/roleAssignments\/removeRoleAssignment\(principalId=(\d+), ?roleDefId=(\d+)\)$/i.exec(rest))) {
        state.assignments[Number(m[1])] = (state.assignments[Number(m[1])] || []).filter((r) => r !== Number(m![2]));
        if (!state.assignments[Number(m[1])].length) delete state.assignments[Number(m[1])];
        return { status: 204 };
      }
      return undefined;
    }
    if ((m = /\/roleDefinitions\/getByType\((\d+)\)/i.exec(u))) return ROLE_IDS[Number(m[1])] ? { body: { Id: ROLE_IDS[Number(m[1])] } } : { status: 404, body: {} };
    if (/\/roleDefinitions\?\$select=Id,Name$/i.test(u)) return { body: [{ Id: 1073741829, Name: 'Teljes hozzáférés' }, { Id: 1073741830, Name: 'Szerkesztés' }] };
    if (/\/_api\/web\/ensureuser$/i.test(u)) {
      const login = (req.body as { logonName: string }).logonName;
      return login === ADELE_T ? { body: { Id: 31, LoginName: ADELE_T } } : { status: 500, body: {} };
    }
    if (/\/_api\/web\/currentUser\?\$select=Id,IsSiteAdmin$/i.test(u)) return { body: { Id: 6, IsSiteAdmin: state.siteAdmin } };
    return undefined;
  }, TARGET);
}

function installContext(): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: TARGET, serverRelativeUrl: '/sites/Cel', title: 'Cél' });
  tokens.set('listurl', 'Teszt_lista', 'Lists/Teszt lista');
  tokens.set('associatedownergroup', undefined, '3');
  tokens.set('groupkey', 'Projektmenedzserek', '12');
  tokens.set('principal', 'adelev', ADELE_T);
  return { targetSiteUrl: TARGET, tokens, log: new Logger() };
}

describe('ListSecurityProvider', () => {
  it('breaks the inheritance of an inheriting list, adds the assignments and removes the installer as site admin', async () => {
    const state: ITargetState = { unique: false, assignments: {}, siteAdmin: true };
    const { sp, requests } = targetSp(state);
    const ctx = installContext();
    expect((await new ListSecurityProvider().diff(sp, def, ctx)).changes).toEqual(['inheritsPermissions', 'assignments:5']);
    expect(await new ListSecurityProvider().apply(sp, def, 'update', ctx)).toMatchObject({ outcome: 'updated' });
    expect(state.assignments).toEqual({ 3: [ROLE_IDS[5]], 12: [ROLE_IDS[6]], 31: [ROLE_IDS[2]] });
    expect(requests.some((r) => /clearSubscopes=false/i.test(r.url))).toBe(true);
    expect(ctx.log.entries.map((e) => e.code).filter((c) => !!c)).toEqual(['ROLE_NOT_FOUND', 'LIST_SECURITY_PRINCIPAL_MISSING', 'LIST_SECURITY_INSTALLER_REMOVED']);

    // A rerun finds what is there: only the custom level and the unknown user are still missing.
    const again = await new ListSecurityProvider().diff(sp, def, installContext());
    expect(again).toMatchObject({ status: 'different', changes: ['assignments:2'] });
  });

  it('keeps the installer when not a site collection admin', async () => {
    const state: ITargetState = { unique: false, assignments: {}, siteAdmin: false };
    const { sp } = targetSp(state);
    const ctx = installContext();
    await new ListSecurityProvider().apply(sp, def, 'update', ctx);
    expect(state.assignments[6]).toEqual([ROLE_IDS[5]]);
    expect(ctx.log.entries.map((e) => e.code)).toContain('LIST_SECURITY_INSTALLER_KEPT');
  });

  it('only adds to a list that is already unique, keeping what is there', async () => {
    const state: ITargetState = { unique: true, assignments: { 6: [ROLE_IDS[5]], 40: [ROLE_IDS[2]] }, siteAdmin: true };
    LOGIN_BY_ID[40] = 'Valaki más';
    const { sp, requests } = targetSp(state);
    await new ListSecurityProvider().apply(sp, def, 'update', installContext());
    expect(requests.some((r) => /breakRoleInheritance|removeRoleAssignment/i.test(r.url))).toBe(false);
    expect(state.assignments[40]).toEqual([ROLE_IDS[2]]);
    expect(state.assignments[6]).toEqual([ROLE_IDS[5]]);
  });

  it('maps the people through the mapping step, and gives nothing to the fallback user', async () => {
    const state: ITargetState = { unique: false, assignments: {}, siteAdmin: true };
    const { sp } = targetSp(state);
    const ctx = installContext();
    const tokens = TokenContext.forSite({ absoluteUrl: TARGET, serverRelativeUrl: '/sites/Cel', title: 'Cél' });
    tokens.set('listurl', 'Teszt_lista', 'Lists/Teszt lista');
    ctx.tokens = tokens; // no principal tokens yet: as in a real install, where only the mapper sets them
    const asked: string[][] = [];
    ctx.content = {
      reader: {} as never,
      idMaps: {},
      principals: {
        map: async (keys: string[], t: TokenContext) => {
          asked.push(keys);
          t.set('principal', 'adelev', ADELE_T);
          t.set('principal', 'nobody', 'i:0#.f|membership|fallback@fabrikam.com');
          return [
            { key: 'adelev', login: ADELE_T, strategy: 'email' },
            { key: 'nobody', login: 'i:0#.f|membership|fallback@fabrikam.com', strategy: 'fallback' }
          ];
        }
      } as never
    };
    const onlyPeople: IListSecurityDef = { ...def, security: { ...security, roleAssignments: security.roleAssignments!.filter((a) => a.principal.indexOf('{principal:') === 0) } };
    await new ListSecurityProvider().apply(sp, onlyPeople, 'update', ctx);
    expect(asked[0]).toEqual(['adelev', 'nobody']);
    expect(state.assignments).toEqual({ 31: [ROLE_IDS[2]] });
    expect(ctx.log.entries.filter((e) => e.code === 'LIST_SECURITY_PRINCIPAL_MISSING').map((e) => e.detail)).toEqual(['{principal:nobody}']);
  });

  it("leaves an existing list alone in 'skip' mode", async () => {
    const state: ITargetState = { unique: false, assignments: {}, siteAdmin: true };
    const { sp, requests } = targetSp(state);
    const ctx = installContext();
    expect(await new ListSecurityProvider().apply(sp, def, 'skip', ctx)).toMatchObject({ outcome: 'skipped' });
    expect(requests.filter((r) => r.method === 'POST')).toEqual([]);
    expect(ctx.log.entries.map((e) => e.code)).toEqual(['LIST_SECURITY_SKIPPED']);
  });
});
