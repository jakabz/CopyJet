import type { FieldType, IField, ITermSetRef } from '../model';
import { parseFieldXml } from './schemaXml';

const FIELD_TYPES: FieldType[] = [
  'Text', 'Note', 'Number', 'Currency', 'Integer', 'Boolean', 'Choice', 'MultiChoice', 'DateTime', 'URL',
  'Lookup', 'LookupMulti', 'User', 'UserMulti', 'TaxonomyFieldType', 'TaxonomyFieldTypeMulti', 'Calculated',
  'Location', 'Image', 'Thumbnail', 'Computed'
];

const EMPTY_GUID = '00000000-0000-0000-0000-000000000000';

/** Field properties as returned by the REST API (subset CopyJet selects). */
export interface IFieldInfoLike {
  Id: string;
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Group?: string;
  Description?: string;
  Required?: boolean;
  Hidden?: boolean;
  SchemaXml: string;
}

export function toFieldType(typeAsString: string): FieldType {
  return FIELD_TYPES.indexOf(typeAsString as FieldType) >= 0 ? (typeAsString as FieldType) : 'Other';
}

function boolAttr(el: Element, name: string): boolean | undefined {
  const v = el.getAttribute(name);
  return v === null ? undefined : v.toUpperCase() === 'TRUE';
}

function childText(el: Element, name: string): string | undefined {
  for (let i = 0; i < el.childNodes.length; i++) {
    const n = el.childNodes[i];
    if (n.nodeName === name) {
      return n.textContent || '';
    }
  }
  return undefined;
}

function choicesOf(el: Element): string[] | undefined {
  const containers = el.getElementsByTagName('CHOICES');
  if (containers.length === 0) {
    return undefined;
  }
  const out: string[] = [];
  const items = containers[0].getElementsByTagName('CHOICE');
  for (let i = 0; i < items.length; i++) {
    out.push(items[i].textContent || '');
  }
  return out;
}

/** Reads SspId / TermSetId / AnchorId / Open from a taxonomy field's <Customization>. */
function termSetOf(el: Element): ITermSetRef | undefined {
  const props: { [name: string]: string } = {};
  const nodes = el.getElementsByTagName('Property');
  for (let i = 0; i < nodes.length; i++) {
    const name = nodes[i].getElementsByTagName('Name')[0];
    const value = nodes[i].getElementsByTagName('Value')[0];
    if (name && value) {
      props[name.textContent || ''] = (value.textContent || '').trim();
    }
  }
  if (!props.TermSetId || props.TermSetId === EMPTY_GUID) {
    return undefined;
  }
  const ref: ITermSetRef = { termSetId: props.TermSetId, path: '' };
  if (props.SspId && props.SspId !== EMPTY_GUID) ref.termStoreId = props.SspId;
  if (props.AnchorId && props.AnchorId !== EMPTY_GUID) ref.anchorId = props.AnchorId;
  if (props.Open) ref.isOpen = props.Open.toLowerCase() === 'true';
  return ref;
}

/**
 * Builds a template field definition from REST field info. `schemaXml` is the (already sanitized and
 * tokenized) XML to store; the descriptive properties are read from it so both stay consistent.
 */
export function toFieldDef(info: IFieldInfoLike, schemaXml: string): IField {
  const el = parseFieldXml(schemaXml);
  const type = toFieldType(info.TypeAsString);
  const def: IField = {
    internalName: info.InternalName,
    type,
    title: info.Title,
    schemaXml
  };
  if (info.Id) def.id = info.Id.replace(/^\{|\}$/g, '').toLowerCase();
  if (info.Group) def.group = info.Group;
  if (info.Description) def.description = info.Description;
  if (info.Required) def.required = true;
  if (info.Hidden) def.hidden = true;
  if (boolAttr(el, 'Indexed')) def.indexed = true;
  if (boolAttr(el, 'EnforceUniqueValues')) def.enforceUniqueValues = true;

  const choices = choicesOf(el);
  if (choices || type === 'Choice' || type === 'MultiChoice') def.choices = choices || [];
  const defaultValue = childText(el, 'Default');
  if (defaultValue !== undefined) def.defaultValue = defaultValue;
  const formula = childText(el, 'Formula');
  if (formula !== undefined || type === 'Calculated') def.formula = formula || '';
  if (type === 'Lookup' || type === 'LookupMulti') {
    def.lookupList = el.getAttribute('List') || '';
    def.lookupField = el.getAttribute('ShowField') || 'Title';
  }
  if (type === 'TaxonomyFieldType' || type === 'TaxonomyFieldTypeMulti') {
    def.termSet = termSetOf(el) || { termSetId: EMPTY_GUID, path: '' };
  }
  const formatter = el.getAttribute('CustomFormatter');
  if (formatter) def.customFormatter = formatter;
  return def;
}

/** Properties compared by diff(); order is the order changes are reported in. */
const COMPARED: Array<keyof IField> = [
  'type', 'title', 'group', 'description', 'required', 'hidden', 'indexed', 'enforceUniqueValues',
  'choices', 'defaultValue', 'formula', 'lookupField', 'customFormatter'
];

function norm(value: unknown): string {
  if (value === undefined || value === null || value === false || value === '') return '';
  return JSON.stringify(value);
}

/** Names of the properties that differ between the template definition and the target field. */
export function compareFields(template: IField, target: IField): string[] {
  return COMPARED.filter((k) => norm(template[k]) !== norm(target[k]));
}
