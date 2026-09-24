import { compareFields, isCustomField, sanitizeFieldXml, toFieldDef, parseFieldXml } from '../../src/core/fields';
import { areaTaxonomyField, budgetCalculatedField, builtInField, clientLookupField, statusField } from '../fixtures/fields';

describe('isCustomField', () => {
  it('tells user-created columns (web GUID SourceID) from built-in ones (schema URI SourceID)', () => {
    expect(isCustomField(statusField.SchemaXml)).toBe(true);
    expect(isCustomField('<Field Name="A" SourceID="7d3e1f20-5a4b-4c6d-9e8f-001122334455" />')).toBe(true);
    expect(isCustomField(builtInField.SchemaXml)).toBe(false);
    expect(isCustomField('<Field Name="A" SourceID="http://schemas.microsoft.com/sharepoint/v3" />')).toBe(false);
    expect(isCustomField('<Field Name="A" />')).toBe(false);
  });
});

describe('sanitizeFieldXml', () => {
  it('drops source identifiers, versioning and script hooks but keeps the definition', () => {
    const r = sanitizeFieldXml(statusField.SchemaXml);
    expect(r.removedAttributes.sort()).toEqual(['JSLink', 'SourceID', 'Version']);
    expect(r.removedElements).toEqual([]);
    const el = parseFieldXml(r.xml);
    expect(el.getAttribute('ID')).toBe('{8c3a1d52-3b4e-4f7a-9c21-5d6e7f8a9b01}');
    expect(el.getAttribute('Name')).toBe('CJ_Status');
    expect(el.getAttribute('DisplayName')).toBe('Státusz');
    expect(el.getElementsByTagName('CHOICE')).toHaveLength(3);
  });

  it('removes WebId, ColName and RowOrdinal from lookups, keeping List/ShowField', () => {
    const r = sanitizeFieldXml(clientLookupField.SchemaXml);
    expect(r.removedAttributes.sort()).toEqual(['ColName', 'RowOrdinal', 'SourceID', 'WebId']);
    expect(r.xml).toContain('List="{0f8b4c2e-1a2b-4c3d-8e9f-a0b1c2d3e4f5}"');
    expect(r.xml).toContain('ShowField="Title"');
  });

  it('removes unknown child elements', () => {
    const r = sanitizeFieldXml('<Field Name="A" Type="Text"><Script>alert(1)</Script><Default>x</Default></Field>');
    expect(r.removedElements).toEqual(['Script']);
    expect(r.xml).toBe('<Field Name="A" Type="Text"><Default>x</Default></Field>');
  });

  it('rejects malformed XML and non-Field roots', () => {
    expect(() => sanitizeFieldXml('<Field Name="A"')).toThrow(expect.objectContaining({ code: 'FIELD_XML_INVALID' }));
    expect(() => sanitizeFieldXml('<View />')).toThrow(expect.objectContaining({ code: 'FIELD_XML_INVALID' }));
  });
});

describe('toFieldDef', () => {
  it('reads choices, default and description', () => {
    const def = toFieldDef(statusField, statusField.SchemaXml);
    expect(def).toMatchObject({
      id: '8c3a1d52-3b4e-4f7a-9c21-5d6e7f8a9b01',
      internalName: 'CJ_Status',
      type: 'Choice',
      title: 'Státusz',
      group: 'CopyJet',
      description: 'Projekt állapota',
      choices: ['Nyitott', 'Folyamatban', 'Lezárt'],
      defaultValue: 'Nyitott'
    });
    expect(def.required).toBeUndefined();
  });

  it('reads lookup target and taxonomy term set', () => {
    expect(toFieldDef(clientLookupField, clientLookupField.SchemaXml)).toMatchObject({
      lookupList: '{0f8b4c2e-1a2b-4c3d-8e9f-a0b1c2d3e4f5}',
      lookupField: 'Title'
    });
    expect(toFieldDef(areaTaxonomyField, areaTaxonomyField.SchemaXml).termSet).toEqual({
      termStoreId: 'c1d2e3f4-a5b6-4c7d-8e9f-0a1b2c3d4e5f',
      termSetId: 'a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d',
      path: '',
      isOpen: false
    });
  });

  it('reads calculated formula', () => {
    expect(toFieldDef(budgetCalculatedField, budgetCalculatedField.SchemaXml).formula).toBe('=[Keret]*1.27');
  });

  it('maps unknown types to Other', () => {
    expect(toFieldDef({ ...statusField, TypeAsString: 'Geolocation' }, '<Field Name="G" Type="Geolocation" />').type).toBe('Other');
  });
});

describe('compareFields', () => {
  it('reports only real differences', () => {
    const a = toFieldDef(statusField, statusField.SchemaXml);
    expect(compareFields(a, { ...a, schemaXml: '<Field />', id: 'other' })).toEqual([]);
    expect(compareFields(a, { ...a, title: 'Status', choices: ['Nyitott'], required: false })).toEqual(['title', 'choices']);
  });
});
