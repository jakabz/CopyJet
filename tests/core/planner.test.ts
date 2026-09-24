import { buildPlan } from '../../src/core/planner';
import type { ICopyJetTemplate } from '../../src/core/model';

const PROJEKT = '0x0100A1B2C3D4E5F60718293A4B5C6D7E8F90';
const KIEMELT = `${PROJEKT}00B1C2D3E4F5061728394A5B6C7D8E9F01`;

function template(): ICopyJetTemplate {
  return {
    schemaVersion: '1.0',
    meta: {
      name: 'Teszt',
      createdBy: 'anna@contoso.com',
      createdAt: '2026-09-24T10:00:00Z',
      sourceSiteUrl: 'https://contoso.sharepoint.com/sites/forras',
      sourceTenant: 'contoso.onmicrosoft.com',
      sourceLcid: 1038,
      includesContent: false,
      generator: { name: 'CopyJet', version: '0.0.1' }
    },
    groups: [
      { key: 'PM', title: '{sitename} PM', owner: '{associatedownergroup}' },
      { key: 'Tamogatas', title: 'Támogatás', owner: '{groupkey:PM}' }
    ],
    siteFields: [
      { internalName: 'CJ_Status', type: 'Text', title: 'Státusz', schemaXml: '<Field />' },
      { internalName: 'CJ_Ugyfel', type: 'Lookup', title: 'Ügyfél', lookupList: '{{listkey:Ugyfelek}}', lookupField: 'Title', schemaXml: '<Field />' }
    ],
    contentTypes: [
      { id: KIEMELT, name: 'Kiemelt', parentId: PROJEKT, fieldRefs: [] },
      { id: PROJEKT, name: 'Projekt', parentId: '0x01', fieldRefs: [{ internalName: 'CJ_Status' }, { internalName: 'Title' }] }
    ],
    lists: [
      { key: 'Ugyfelek', url: 'Lists/Ugyfelek', title: 'Ügyfelek', template: 100, content: { mode: 'none' } },
      {
        key: 'Projektek',
        url: 'Lists/Projektek',
        title: 'Projektek',
        template: 100,
        contentTypes: [PROJEKT, '0x01'],
        content: { mode: 'none' },
        fields: [
          { internalName: 'Ugyfel', type: 'Lookup', title: 'Ügyfél', lookupList: '{{listkey:Ugyfelek}}', lookupField: 'Title', schemaXml: '<Field />' },
          { internalName: 'CJ_Status', type: 'Text', title: 'Státusz', schemaXml: '<Field />' }
        ],
        views: [{ title: 'Minden elem', default: true, fields: ['LinkTitle', 'Ugyfel'] }]
      }
    ],
    pages: [],
    principals: []
  };
}

const keys = (steps: Array<{ ref: { key: string } }>): string[] => steps.map((s) => s.ref.key);

describe('buildPlan', () => {
  it('orders steps into dependency levels, kind order within a level', () => {
    const plan = buildPlan(template());
    expect(plan.levels.map(keys)).toEqual([
      ['group:PM', 'field:CJ_Status', 'list:Ugyfelek'],
      ['group:Tamogatas', 'field:CJ_Ugyfel', `contentType:${PROJEKT}`],
      [`contentType:${KIEMELT}`, 'list:Projektek'],
      ['listField:Projektek/CJ_Status', 'listField:Projektek/Ugyfel'],
      ['view:Projektek/Minden elem']
    ]);
    expect(plan.excluded).toEqual([]);
    expect(plan.steps).toHaveLength(11); // 2 groups, 2 site columns, 2 content types, 2 lists, 2 list columns, 1 view
    // Built-in dependencies (Title, 0x01, LinkTitle) are not edges.
    expect(plan.steps.find((s) => s.ref.key === `contentType:${PROJEKT}`)!.dependsOn).toEqual(['field:CJ_Status']);
    expect(plan.steps.find((s) => s.ref.key === 'listField:Projektek/Ugyfel')!.dependsOn).toEqual(['list:Projektek', 'list:Ugyfelek']);
  });

  it('carries the provider input: list columns and views with their list', () => {
    const plan = buildPlan(template());
    expect(plan.steps.find((s) => s.ref.kind === 'view')!.def).toMatchObject({ listKey: 'Projektek', listUrl: 'Lists/Projektek', view: { title: 'Minden elem' } });
  });

  it('switches off a disabled step and everything depending on it', () => {
    const plan = buildPlan(template(), { disabled: ['list:Ugyfelek', 'unknown:key'] });
    expect(plan.excluded).toEqual([
      { ref: { kind: 'list', key: 'list:Ugyfelek' }, reason: 'disabled' },
      { ref: { kind: 'siteField', key: 'field:CJ_Ugyfel' }, reason: 'dependencyDisabled', cause: 'list:Ugyfelek' },
      { ref: { kind: 'listField', key: 'listField:Projektek/Ugyfel' }, reason: 'dependencyDisabled', cause: 'list:Ugyfelek' },
      { ref: { kind: 'view', key: 'view:Projektek/Minden elem' }, reason: 'dependencyDisabled', cause: 'list:Ugyfelek' }
    ]);
    expect(keys(plan.steps)).not.toContain('list:Ugyfelek');
    expect(keys(plan.steps)).toContain('list:Projektek');
  });

  it('excludes steps caught in a dependency cycle instead of guessing an order', () => {
    const t = template();
    // Projekt links a lookup column pointing at Projektek, which uses Projekt: list → CT → column → list.
    t.siteFields.push({ internalName: 'CJ_Szulo', type: 'Lookup', title: 'Szülő', lookupList: '{{listkey:Projektek}}', lookupField: 'Title', schemaXml: '<Field />' });
    t.contentTypes[1].fieldRefs.push({ internalName: 'CJ_Szulo' });
    const plan = buildPlan(t);
    expect(plan.excluded.filter((e) => e.reason === 'cycle').map((e) => e.ref.key).sort()).toEqual(
      [`contentType:${KIEMELT}`, `contentType:${PROJEKT}`, 'field:CJ_Szulo', 'list:Projektek', 'listField:Projektek/CJ_Status', 'listField:Projektek/Ugyfel', 'view:Projektek/Minden elem'].sort()
    );
    expect(keys(plan.steps)).toEqual(['group:PM', 'field:CJ_Status', 'list:Ugyfelek', 'group:Tamogatas', 'field:CJ_Ugyfel']);
  });
});
