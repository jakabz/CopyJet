// Fields of the Forrás "Teszt lista" as REST returns them (spike 05 A).
import type { IListFieldInfoLike } from '../../src/core/lists';
import { lookupForras, tesztLista } from './lists';

const WEB_ID = '8b98bd9e-6008-44fa-a2e2-784ff594eab3';
const LIST = `{${tesztLista.Id}}`;

const f = (InternalName: string, TypeAsString: string, schema: string, extra: Partial<IListFieldInfoLike> = {}): IListFieldInfoLike => ({
  Id: extra.Id || `44444444-0000-4000-8000-${InternalName.length.toString().padStart(12, '0')}`,
  InternalName,
  Title: InternalName,
  TypeAsString,
  Hidden: false,
  FromBaseType: false,
  CanBeDeleted: true,
  ...extra,
  SchemaXml: schema
});

export const title = f('Title', 'Text', '<Field ID="{fa564e0f-0c70-4ab9-b863-0177e6ddd247}" Name="Title" SourceID="http://schemas.microsoft.com/sharepoint/v3" Type="Text" DisplayName="Cím" Required="TRUE" />', {
  Id: 'fa564e0f-0c70-4ab9-b863-0177e6ddd247',
  FromBaseType: true,
  CanBeDeleted: false
});
export const created = f('Created', 'DateTime', '<Field ID="{8c06beca-0777-48f7-91c7-6da68bc07b69}" Name="Created" SourceID="http://schemas.microsoft.com/sharepoint/v3" Type="DateTime" DisplayName="Létrehozva" />', {
  Id: '8c06beca-0777-48f7-91c7-6da68bc07b69',
  FromBaseType: true,
  CanBeDeleted: false
});
export const valaki = f('Valaki', 'User', `<Field ID="{55555555-0000-4000-8000-000000000001}" Name="Valaki" StaticName="Valaki" SourceID="${LIST}" Type="User" DisplayName="Valaki" UserSelectionMode="PeopleOnly" ColName="int1" RowOrdinal="0" />`, {
  Id: '55555555-0000-4000-8000-000000000001'
});
export const valassz = f(
  'V_x00e1_lassz',
  'Choice',
  `<Field ID="{55555555-0000-4000-8000-000000000002}" Name="V_x00e1_lassz" StaticName="V_x00e1_lassz" SourceID="${LIST}" Type="Choice" DisplayName="Válassz" Format="Dropdown"><CHOICES><CHOICE>Egy</CHOICE><CHOICE>Kettő</CHOICE></CHOICES></Field>`,
  { Id: '55555555-0000-4000-8000-000000000002', Title: 'Válassz' }
);
export const listaLookup = f(
  'ListaLookup',
  'Lookup',
  `<Field ID="{55555555-0000-4000-8000-000000000003}" Name="ListaLookup" StaticName="ListaLookup" SourceID="${LIST}" Type="Lookup" DisplayName="ListaLookup" List="{${lookupForras.Id}}" WebId="${WEB_ID}" ShowField="Title" />`,
  { Id: '55555555-0000-4000-8000-000000000003' }
);
export const siteColumn1 = f('SiteColumn1', 'Text', `<Field ID="{66666666-0000-4000-8000-000000000001}" Name="SiteColumn1" StaticName="SiteColumn1" SourceID="{${WEB_ID}}" Type="Text" DisplayName="SiteColumn1" Group="Egyéni oszlopok" />`, {
  Id: '66666666-0000-4000-8000-000000000001'
});
// Built-in site column added to the list: schema URI SourceID, but not from the base type.
export const startDate = f('StartDate', 'DateTime', '<Field ID="{64cd368d-2f95-4bfc-a1f9-8d4324ecb007}" Name="StartDate" SourceID="http://schemas.microsoft.com/sharepoint/v3" Type="DateTime" DisplayName="Kezdő dátum" />', {
  Id: '64cd368d-2f95-4bfc-a1f9-8d4324ecb007'
});
export const hiddenNote = f('ListaSzoveg_0', 'Note', `<Field ID="{55555555-0000-4000-8000-000000000009}" Name="ListaSzoveg_0" SourceID="${LIST}" Type="Note" Hidden="TRUE" />`, {
  Id: '55555555-0000-4000-8000-000000000009',
  Hidden: true
});

export const tesztListaFields: IListFieldInfoLike[] = [title, created, valaki, valassz, listaLookup, siteColumn1, startDate, hiddenNote];
