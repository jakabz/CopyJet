import { missingDependencies } from '../../src/core/engine';
import type { ICopyJetTemplate, IDiscoveredArtifact } from '../../src/core/model';

const found = (kind: IDiscoveredArtifact['ref']['kind'], key: string, unsupported?: string): IDiscoveredArtifact => ({ ref: { kind, key }, title: key, unsupported });

describe('missingDependencies', () => {
  it('offers discovered, supported artifacts the selection needs but lacks', () => {
    const template = {
      groups: [],
      siteFields: [{ internalName: 'CJ_Ugyfel', type: 'Lookup', title: 'Ügyfél', lookupList: '{{listkey:Ugyfelek}}', lookupField: 'Title', schemaXml: '<Field />' }],
      contentTypes: [],
      lists: [
        {
          key: 'Projektek',
          url: 'Lists/Projektek',
          title: 'Projektek',
          template: 100,
          content: { mode: 'none' },
          fields: [{ internalName: 'Kapcsolat', type: 'Lookup', title: 'Kapcsolat', lookupList: '{{listkey:Kapcsolatok}}', lookupField: 'Title', schemaXml: '<Field />' }],
          views: [{ title: 'Minden elem', fields: ['LinkTitle', 'Kapcsolat'] }]
        }
      ]
    } as unknown as ICopyJetTemplate;
    const discovered = [
      found('list', 'list:Ugyfelek'),
      found('list', 'list:Kapcsolatok', 'LIST_TEMPLATE_UNSUPPORTED'),
      found('list', 'list:Projektek'),
      found('listField', 'listField:Projektek/Kapcsolat'),
      found('list', 'list:Egyeb')
    ];
    // Ugyfelek is needed (site column lookup); Kapcsolatok is unsupported; the view's column is in the template;
    // LinkTitle was never discovered (built-in).
    expect(missingDependencies(template, discovered).map((a) => a.ref.key)).toEqual(['list:Ugyfelek']);
  });
});
