/**
 * CSOM request bodies (client.svc/ProcessQuery) as pure functions, so tests and spike scripts see exactly
 * what the provider sends. Type IDs were read from SharePoint Online's sp.js and verified (spike 02).
 */

export const CSOM_TYPE = {
  RequestContext: '{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}',
  ContentTypeCreationInformation: '{168f3091-4554-4f14-8866-b20d48e45b54}',
  FieldLinkCreationInformation: '{63fb2c92-8f65-4bbb-a658-b6cd294403f4}'
};

export interface ICsomBody {
  actions: string;
  objectPaths: string;
}

export function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

export function wrapRequest(body: ICsomBody): string {
  return (
    '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CopyJet">' +
    `<Actions>${body.actions}</Actions><ObjectPaths>${body.objectPaths}</ObjectPaths></Request>`
  );
}

const stringProp = (name: string, value: string): string => `<Property Name="${name}" Type="String">${escapeXml(value)}</Property>`;
const boolParam = (value: boolean): string => `<Parameter Type="Boolean">${value ? 'true' : 'false'}</Parameter>`;
const ROOT = `<StaticProperty Id="1" TypeId="${CSOM_TYPE.RequestContext}" Name="Current" /><Property Id="2" ParentId="1" Name="Web" />`;

/** Query id whose response element carries the new content type (see csomCreateContentType). */
export const CREATE_CT_QUERY_ID = 11;

/** web.ContentTypes.Add({ Id, Name, Group, Description }) – the only way to keep the requested ID. */
export function createContentTypeBody(ct: { id: string; name: string; group?: string; description?: string }): ICsomBody {
  return {
    actions:
      '<ObjectPath Id="10" ObjectPathId="4" />' +
      `<Query Id="${CREATE_CT_QUERY_ID}" ObjectPathId="4"><Query SelectAllProperties="false"><Properties><Property Name="StringId" ScalarProperty="true" /></Properties></Query></Query>`,
    objectPaths:
      ROOT +
      '<Property Id="3" ParentId="2" Name="ContentTypes" />' +
      `<Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="${CSOM_TYPE.ContentTypeCreationInformation}">` +
      stringProp('Description', ct.description || '') +
      stringProp('Group', ct.group || '') +
      stringProp('Id', ct.id) +
      stringProp('Name', ct.name) +
      '<Property Name="ParentContentType" Type="Null" />' +
      '</Parameter></Parameters></Method>'
  };
}

export interface IFieldLinkAdd {
  internalName: string;
  required: boolean;
  hidden: boolean;
}

export interface IFieldLinkFlags {
  /** FieldLink Id (the field's GUID). */
  id: string;
  required: boolean;
  hidden: boolean;
}

/**
 * On one content type: adds field links (web.AvailableFields.GetByInternalNameOrTitle) and sets
 * Required/Hidden on new and existing links, then ContentType.Update(false) – child types are not pushed.
 * REST can do neither: POST .../fieldlinks reports existing columns as missing, and FieldLink rejects MERGE.
 */
export function fieldLinksBody(contentTypeId: string, add: IFieldLinkAdd[], flags: IFieldLinkFlags[]): ICsomBody {
  let id = 10;
  let actions = '';
  let objectPaths =
    ROOT +
    '<Property Id="3" ParentId="2" Name="ContentTypes" />' +
    `<Method Id="4" ParentId="3" Name="GetById"><Parameters><Parameter Type="String">${escapeXml(contentTypeId)}</Parameter></Parameters></Method>` +
    '<Property Id="5" ParentId="4" Name="FieldLinks" />' +
    '<Property Id="6" ParentId="2" Name="AvailableFields" />';
  actions += `<ObjectPath Id="${id++}" ObjectPathId="4" />`;

  const setFlags = (linkPath: number, required: boolean, hidden: boolean): void => {
    actions += `<SetProperty Id="${id++}" ObjectPathId="${linkPath}" Name="Required">${boolParam(required)}</SetProperty>`;
    actions += `<SetProperty Id="${id++}" ObjectPathId="${linkPath}" Name="Hidden">${boolParam(hidden)}</SetProperty>`;
  };

  let path = 100;
  add.forEach((link) => {
    const fieldPath = path++;
    const linkPath = path++;
    objectPaths +=
      `<Method Id="${fieldPath}" ParentId="6" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">${escapeXml(link.internalName)}</Parameter></Parameters></Method>` +
      `<Method Id="${linkPath}" ParentId="5" Name="Add"><Parameters><Parameter TypeId="${CSOM_TYPE.FieldLinkCreationInformation}">` +
      `<Property Name="Field" ObjectPathId="${fieldPath}" /></Parameter></Parameters></Method>`;
    actions += `<ObjectPath Id="${id++}" ObjectPathId="${fieldPath}" /><ObjectPath Id="${id++}" ObjectPathId="${linkPath}" />`;
    setFlags(linkPath, link.required, link.hidden);
  });
  flags.forEach((f) => {
    const linkPath = path++;
    objectPaths += `<Method Id="${linkPath}" ParentId="5" Name="GetById"><Parameters><Parameter Type="Guid">{${escapeXml(f.id.replace(/[{}]/g, ''))}}</Parameter></Parameters></Method>`;
    actions += `<ObjectPath Id="${id++}" ObjectPathId="${linkPath}" />`;
    setFlags(linkPath, f.required, f.hidden);
  });
  actions += `<Method Name="Update" Id="${id++}" ObjectPathId="4"><Parameters>${boolParam(false)}</Parameters></Method>`;
  return { actions, objectPaths };
}

/**
 * web.SiteGroups.GetById(groupId).Owner = web.SiteGroups.GetById(ownerId); Update(). REST's
 * SetUserAsOwner accepts only users – with a group ID it returns 200 and changes nothing (spike 07).
 */
export function setGroupOwnerBody(groupId: number, ownerGroupId: number): ICsomBody {
  return {
    actions:
      '<ObjectPath Id="10" ObjectPathId="4" /><ObjectPath Id="11" ObjectPathId="5" />' +
      '<SetProperty Id="12" ObjectPathId="4" Name="Owner"><Parameter ObjectPathId="5" /></SetProperty>' +
      '<Method Name="Update" Id="13" ObjectPathId="4" />',
    objectPaths:
      ROOT +
      '<Property Id="3" ParentId="2" Name="SiteGroups" />' +
      `<Method Id="4" ParentId="3" Name="GetById"><Parameters><Parameter Type="Int32">${Math.floor(groupId)}</Parameter></Parameters></Method>` +
      `<Method Id="5" ParentId="3" Name="GetById"><Parameters><Parameter Type="Int32">${Math.floor(ownerGroupId)}</Parameter></Parameters></Method>`
  };
}

/** Query id whose response element carries the created or added field (Id, InternalName, Title). */
export const FIELD_QUERY_ID = 20;

const fieldQuery = (objectPathId: number): string =>
  `<Query Id="${FIELD_QUERY_ID}" ObjectPathId="${objectPathId}"><Query SelectAllProperties="false"><Properties>` +
  '<Property Name="Id" ScalarProperty="true" /><Property Name="InternalName" ScalarProperty="true" /><Property Name="Title" ScalarProperty="true" />' +
  '</Properties></Query></Query>';

/** Web.GetList(url).Fields, or Web.Fields without a list: object path 4. */
const fieldsPath = (listUrl: string | undefined): string =>
  listUrl
    ? `<Method Id="3" ParentId="2" Name="GetList"><Parameters><Parameter Type="String">${escapeXml(listUrl)}</Parameter></Parameters></Method><Property Id="4" ParentId="3" Name="Fields" />`
    : '<Property Id="4" ParentId="2" Name="Fields" />';

export interface ITaxonomyFieldSpec {
  /** A plain field element (Type, ID, Name, DisplayName, ShowField, Mult …) without Customization. */
  schemaXml: string;
  termStoreId: string;
  termSetId: string;
  anchorId?: string;
  open?: boolean;
}

const guidParam = (id: string): string => `<Parameter Type="Guid">{${escapeXml(id.replace(/[{}]/g, ''))}}</Parameter>`;

/**
 * A Managed Metadata column (spike 11 C2): Fields.AddFieldAsXml(plain field), then SspId / TermSetId /
 * AnchorId / Open on the same object and Update() – one request. SharePoint adds the hidden note field and
 * the TaxonomyHiddenList binding itself. A SchemaXml with Customization fails with HTTP 500 and leaves the
 * list unusable, so this is the only way CopyJet creates such columns.
 */
export function createTaxonomyFieldBody(listUrl: string | undefined, spec: ITaxonomyFieldSpec, options: number): ICsomBody {
  return {
    actions:
      '<ObjectPath Id="10" ObjectPathId="5" />' +
      `<SetProperty Id="11" ObjectPathId="5" Name="SspId">${guidParam(spec.termStoreId)}</SetProperty>` +
      `<SetProperty Id="12" ObjectPathId="5" Name="TermSetId">${guidParam(spec.termSetId)}</SetProperty>` +
      `<SetProperty Id="13" ObjectPathId="5" Name="AnchorId">${guidParam(spec.anchorId || '00000000-0000-0000-0000-000000000000')}</SetProperty>` +
      `<SetProperty Id="14" ObjectPathId="5" Name="Open">${boolParam(!!spec.open)}</SetProperty>` +
      '<Method Name="Update" Id="15" ObjectPathId="5" />' +
      fieldQuery(5),
    objectPaths:
      ROOT +
      fieldsPath(listUrl) +
      `<Method Id="5" ParentId="4" Name="AddFieldAsXml"><Parameters><Parameter Type="String">${escapeXml(spec.schemaXml)}</Parameter>${boolParam(false)}<Parameter Type="Enum">${Math.floor(options)}</Parameter></Parameters></Method>`
  };
}

/** List.Fields.Add(Web.AvailableFields.GetById(id)): puts an existing site column on a list as it is. */
export function addExistingFieldBody(listUrl: string, fieldId: string): ICsomBody {
  return {
    actions: '<ObjectPath Id="10" ObjectPathId="7" /><ObjectPath Id="11" ObjectPathId="5" />' + fieldQuery(5),
    objectPaths:
      ROOT +
      fieldsPath(listUrl) +
      '<Property Id="6" ParentId="2" Name="AvailableFields" />' +
      `<Method Id="7" ParentId="6" Name="GetById"><Parameters>${guidParam(fieldId)}</Parameters></Method>` +
      '<Method Id="5" ParentId="4" Name="Add"><Parameters><Parameter ObjectPathId="7" /></Parameters></Method>'
  };
}
