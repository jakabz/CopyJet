import { CopyJetError } from '../errors';

/**
 * Attributes of the root <Field> element that may pass into a template and onto the target site.
 * Everything else is dropped: source-site identifiers (SourceID, WebId, ColName, RowOrdinal), schema
 * versioning (Version) and script hooks (JSLink). See docs/spikes/01-field-schemaxml.md.
 */
const ALLOWED_ATTRIBUTES: string[] = [
  // identity and common
  'ID', 'Name', 'StaticName', 'DisplayName', 'Type', 'Group', 'Description', 'Required', 'Hidden', 'ReadOnly',
  'Sealed', 'Indexed', 'EnforceUniqueValues', 'Viewable', 'Filterable', 'Sortable',
  'ShowInNewForm', 'ShowInEditForm', 'ShowInDisplayForm', 'ShowInViewForms', 'ShowInFileDlg', 'ShowInListSettings',
  'CustomFormatter', 'ClientSideComponentId', 'ClientSideComponentProperties', 'Format',
  'AuthoringInfo', 'Dir', 'IMEMode', 'ShowInVersionHistory',
  // text / note
  'MaxLength', 'NumLines', 'RichText', 'RichTextMode', 'AppendOnly', 'UnlimitedLengthInDocumentLibrary',
  'IsolateStyles', 'RestrictedMode',
  // number / currency / calculated
  'Min', 'Max', 'Decimals', 'Percentage', 'LCID', 'CommaSeparator', 'ResultType',
  // choice
  'FillInChoice',
  // date
  'FriendlyDisplayFormat', 'CalType', 'StorageTZ', 'DisplayFormat',
  // lookup / user / taxonomy
  'List', 'ShowField', 'Mult', 'RelationshipDeleteBehavior', 'PrependId', 'FieldRef',
  'UserSelectionMode', 'UserSelectionScope', 'Presence'
];

/** Child elements of <Field> that carry field definition data. */
const ALLOWED_ELEMENTS: string[] = ['Default', 'DefaultFormula', 'CHOICES', 'MAPPINGS', 'Formula', 'FieldRefs', 'Validation', 'Customization'];

export interface ISanitizeResult {
  xml: string;
  removedAttributes: string[];
  removedElements: string[];
}

export function parseFieldXml(xml: string): Element {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const root = doc.documentElement;
  if (!root || root.nodeName !== 'Field' || doc.getElementsByTagName('parsererror').length > 0) {
    throw new CopyJetError('FIELD_XML_INVALID', 'SchemaXml is not a well-formed <Field> element.', { xml });
  }
  return root;
}

export function serializeFieldXml(root: Element): string {
  return new XMLSerializer().serializeToString(root);
}

/** Keeps only allowlisted attributes and child elements of a <Field> SchemaXml. */
export function sanitizeFieldXml(xml: string): ISanitizeResult {
  const root = parseFieldXml(xml);
  const removedAttributes: string[] = [];
  const removedElements: string[] = [];

  Array.prototype.slice.call(root.attributes).forEach((attr: Attr) => {
    if (ALLOWED_ATTRIBUTES.indexOf(attr.name) < 0) {
      removedAttributes.push(attr.name);
      root.removeAttribute(attr.name);
    }
  });
  Array.prototype.slice.call(root.childNodes).forEach((node: Node) => {
    if (node.nodeType === 1 && ALLOWED_ELEMENTS.indexOf(node.nodeName) < 0) {
      removedElements.push(node.nodeName);
      root.removeChild(node);
    }
  });

  return { xml: serializeFieldXml(root), removedAttributes, removedElements };
}

/**
 * True for columns created by users or solutions. Built-in columns carry a schema URI as SourceID
 * (http://schemas.microsoft.com/sharepoint/v3...), user-created ones the GUID of the web that created them.
 * CanBeDeleted is true for most built-in columns too, so it cannot be used (spike 01).
 */
export function isCustomField(schemaXml: string): boolean {
  const m = /\sSourceID="\{?([0-9a-fA-F-]{36})\}?"/.exec(schemaXml);
  return !!m && /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(m[1]);
}

/** Returns a copy of `xml` with the root attribute set (or removed when value is undefined). */
export function setFieldXmlAttribute(xml: string, name: string, value: string | undefined): string {
  const root = parseFieldXml(xml);
  if (value === undefined) {
    root.removeAttribute(name);
  } else {
    root.setAttribute(name, value);
  }
  return serializeFieldXml(root);
}
