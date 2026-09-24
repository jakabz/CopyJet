// REST responses for /_api/web/fields, shaped like SharePoint Online returns them (odata=nometadata).
import type { IFieldInfoLike } from '../../src/core/fields';

export const SOURCE_WEB_ID = '7d3e1f20-5a4b-4c6d-9e8f-001122334455';
export const UGYFELEK_LIST_ID = '0f8b4c2e-1a2b-4c3d-8e9f-a0b1c2d3e4f5';

export const statusField: IFieldInfoLike = {
  Id: '8c3a1d52-3b4e-4f7a-9c21-5d6e7f8a9b01',
  InternalName: 'CJ_Status',
  Title: 'Státusz',
  TypeAsString: 'Choice',
  Group: 'CopyJet',
  Description: 'Projekt állapota',
  Required: false,
  Hidden: false,
  SchemaXml:
    `<Field Type="Choice" DisplayName="Státusz" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Dropdown" ` +
    `FillInChoice="FALSE" Group="CopyJet" Description="Projekt állapota" ID="{8c3a1d52-3b4e-4f7a-9c21-5d6e7f8a9b01}" ` +
    `SourceID="{${SOURCE_WEB_ID}}" StaticName="CJ_Status" Name="CJ_Status" Version="3" JSLink="~site/SiteAssets/x.js">` +
    `<Default>Nyitott</Default><CHOICES><CHOICE>Nyitott</CHOICE><CHOICE>Folyamatban</CHOICE><CHOICE>Lezárt</CHOICE></CHOICES></Field>`
};

export const clientLookupField: IFieldInfoLike = {
  Id: '3a4b5c6d-7e8f-4a0b-9c1d-2e3f4a5b6c7d',
  InternalName: 'CJ_Ugyfel',
  Title: 'Ügyfél',
  TypeAsString: 'Lookup',
  Group: 'CopyJet',
  Required: false,
  Hidden: false,
  SchemaXml:
    `<Field Type="Lookup" DisplayName="Ügyfél" Required="FALSE" EnforceUniqueValues="FALSE" List="{${UGYFELEK_LIST_ID}}" ` +
    `WebId="${SOURCE_WEB_ID}" ShowField="Title" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ` +
    `Group="CopyJet" ID="{3a4b5c6d-7e8f-4a0b-9c1d-2e3f4a5b6c7d}" SourceID="{${SOURCE_WEB_ID}}" StaticName="CJ_Ugyfel" ` +
    `Name="CJ_Ugyfel" ColName="int1" RowOrdinal="0" />`
};

export const areaTaxonomyField: IFieldInfoLike = {
  Id: '1f2e3d4c-5b6a-4978-8a1b-2c3d4e5f6a7b',
  InternalName: 'CJ_Terulet',
  Title: 'Terület',
  TypeAsString: 'TaxonomyFieldType',
  Group: 'CopyJet',
  Required: false,
  Hidden: false,
  SchemaXml:
    `<Field Type="TaxonomyFieldType" DisplayName="Terület" List="{aaaaaaaa-1111-4222-8333-444444444444}" ` +
    `WebId="${SOURCE_WEB_ID}" ShowField="Term1038" Required="FALSE" EnforceUniqueValues="FALSE" Group="CopyJet" ` +
    `ID="{1f2e3d4c-5b6a-4978-8a1b-2c3d4e5f6a7b}" SourceID="{${SOURCE_WEB_ID}}" StaticName="CJ_Terulet" Name="CJ_Terulet">` +
    `<Default></Default><Customization><ArrayOfProperty>` +
    `<Property><Name>SspId</Name><Value xmlns:q1="http://www.w3.org/2001/XMLSchema" p4:type="q1:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">c1d2e3f4-a5b6-4c7d-8e9f-0a1b2c3d4e5f</Value></Property>` +
    `<Property><Name>TermSetId</Name><Value xmlns:q2="http://www.w3.org/2001/XMLSchema" p4:type="q2:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d</Value></Property>` +
    `<Property><Name>AnchorId</Name><Value xmlns:q3="http://www.w3.org/2001/XMLSchema" p4:type="q3:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">00000000-0000-0000-0000-000000000000</Value></Property>` +
    `<Property><Name>Open</Name><Value xmlns:q5="http://www.w3.org/2001/XMLSchema" p4:type="q5:boolean" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">false</Value></Property>` +
    `</ArrayOfProperty></Customization></Field>`
};

export const budgetCalculatedField: IFieldInfoLike = {
  Id: '9e8d7c6b-5a49-4382-a716-f5e4d3c2b1a0',
  InternalName: 'CJ_Keret',
  Title: 'Keret (bruttó)',
  TypeAsString: 'Calculated',
  Group: 'CopyJet',
  Required: false,
  Hidden: false,
  SchemaXml:
    `<Field Type="Calculated" DisplayName="Keret (bruttó)" ResultType="Currency" ReadOnly="TRUE" LCID="1038" Decimals="0" ` +
    `Group="CopyJet" ID="{9e8d7c6b-5a49-4382-a716-f5e4d3c2b1a0}" SourceID="{${SOURCE_WEB_ID}}" StaticName="CJ_Keret" ` +
    `Name="CJ_Keret" Version="1"><Formula>=[Keret]*1.27</Formula><FieldRefs><FieldRef Name="Keret" /></FieldRefs></Field>`
};

/** Built-in column: CanBeDeleted and visible like a custom one, but its SourceID is a schema URI. */
export const builtInField: IFieldInfoLike = {
  Id: '3c0e9e00-8fcc-479f-9d8d-3447cda34c5b',
  InternalName: 'OtherAddressCountry',
  Title: 'Egyéb cím – ország',
  TypeAsString: 'Text',
  Group: 'Alapszintű Névjegy és Naptár oszlopok',
  Required: false,
  Hidden: false,
  SchemaXml:
    '<Field ID="{3c0e9e00-8fcc-479f-9d8d-3447cda34c5b}" Name="OtherAddressCountry" StaticName="OtherAddressCountry" ' +
    'SourceID="http://schemas.microsoft.com/sharepoint/v3/fields" Group="Alapszintű Névjegy és Naptár oszlopok" ' +
    'DisplayName="Egyéb cím – ország" Type="Text" Sealed="TRUE" AllowDeletion="TRUE" />'
};

export const allSourceFields = [statusField, clientLookupField, areaTaxonomyField, budgetCalculatedField];
