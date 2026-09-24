// REST responses for /_api/web/contenttypes and .../fieldlinks (odata=nometadata).
import type { IContentTypeInfoLike, IFieldLinkInfoLike } from '../../src/core/contentTypes';

export const PROJEKT_ID = '0x0100A1B2C3D4E5F60718293A4B5C6D7E8F90';
export const KIEMELT_ID = `${PROJEKT_ID}00B1C2D3E4F5061728394A5B6C7D8E9F01`;

const link = (name: string, required = false, hidden = false): IFieldLinkInfoLike => ({ Id: `id-${name}`, FieldInternalName: name, Name: name, Required: required, Hidden: hidden });

export const itemCt: IContentTypeInfoLike = {
  StringId: '0x01',
  Name: 'Elem',
  Group: 'Lista-tartalomtípusok',
  Description: 'Új elem létrehozása.',
  SchemaXml: '<ContentType ID="0x01" Name="Elem" Group="Lista-tartalomtípusok" FeatureId="{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}" Version="0"><FieldRefs /></ContentType>'
};

export const documentCt: IContentTypeInfoLike = {
  StringId: '0x0101',
  Name: 'Dokumentum',
  Group: 'Dokumentum-tartalomtípusok',
  SchemaXml: '<ContentType ID="0x0101" Name="Dokumentum" FeatureId="{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}"><FieldRefs /></ContentType>'
};

export const hiddenCt: IContentTypeInfoLike = {
  StringId: '0x0100D7E8F90A1B2C3D4E5F60718293A4B5C6',
  Name: 'Rejtett rendszer',
  Group: '_Hidden',
  Hidden: true,
  SchemaXml: '<ContentType ID="0x0100D7E8F90A1B2C3D4E5F60718293A4B5C6" Name="Rejtett rendszer"><FieldRefs /></ContentType>'
};

export const projektCt: IContentTypeInfoLike = {
  StringId: PROJEKT_ID,
  Name: 'Projekt',
  Group: 'CopyJet',
  Description: 'Projekt adatlap',
  SchemaXml: `<ContentType ID="${PROJEKT_ID}" Name="Projekt" Group="CopyJet" Description="Projekt adatlap" Version="2"><FieldRefs /></ContentType>`
};

export const kiemeltCt: IContentTypeInfoLike = {
  StringId: KIEMELT_ID,
  Name: 'Kiemelt projekt',
  Group: 'CopyJet',
  SchemaXml: `<ContentType ID="${KIEMELT_ID}" Name="Kiemelt projekt" Group="CopyJet"><FieldRefs /></ContentType>`
};

/** Built-in but provisioned without FeatureId (spike 02). */
export const sitePageCt: IContentTypeInfoLike = {
  StringId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
  Name: 'Webhelylap',
  Group: 'Dokumentumtartalom-típusok',
  SchemaXml: '<ContentType ID="0x0101009D1CB255DA76424F860D91F20E6C4118" Name="Webhelylap" Group="Dokumentumtartalom-típusok" Version="0"><FieldRefs /></ContentType>'
};

/** A user-created page type: child of Site Page, no FeatureId – custom. */
export const customPageCt: IContentTypeInfoLike = {
  StringId: '0x0101009D1CB255DA76424F860D91F20E6C411800AB12CD34EF56AB78CD90EF12AB34CD56',
  Name: 'Projektlap',
  Group: 'CopyJet',
  SchemaXml: '<ContentType ID="0x0101009D1CB255DA76424F860D91F20E6C411800AB12CD34EF56AB78CD90EF12AB34CD56" Name="Projektlap" Version="0"><FieldRefs /></ContentType>'
};

export const sourceContentTypes = [itemCt, documentCt, hiddenCt, sitePageCt, projektCt, kiemeltCt];

export const sourceLinks: { [id: string]: IFieldLinkInfoLike[] } = {
  '0x01': [link('ContentType'), link('Title', true)],
  [PROJEKT_ID]: [link('ContentType'), link('Title', true), link('CJ_Status', true), link('CJ_Ugyfel')],
  // Child overrides Title (hidden) and adds CJ_Keret.
  [KIEMELT_ID]: [link('ContentType'), link('Title', true, true), link('CJ_Status', true), link('CJ_Ugyfel'), link('CJ_Keret')]
};

export { link };
