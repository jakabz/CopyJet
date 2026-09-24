import type { IField } from '../model';
import { resolve, type TokenContext } from '../tokenizer';
import { toFieldDef, type IFieldInfoLike } from './fieldModel';

/** compareFields() changes that can be applied to an existing field by MERGE. */
export const UPDATABLE_FIELD_CHANGES = ['title', 'group', 'description', 'required', 'choices', 'defaultValue', 'customFormatter'];

/** The definition with tokens in comparable properties resolved against the target. */
export function resolveFieldDef(def: IField, tokens: TokenContext): IField {
  const r = (v: string | undefined): string | undefined => (v === undefined ? v : resolve(v, tokens));
  return { ...def, title: r(def.title)!, defaultValue: def.defaultValue === null ? null : r(def.defaultValue), formula: r(def.formula) };
}

function union(a: string[] | undefined, b: string[] | undefined): string[] {
  const out = (a || []).slice();
  (b || []).forEach((c) => {
    if (out.indexOf(c) < 0) out.push(c);
  });
  return out;
}

export interface IFieldUpdate {
  /** REST properties to MERGE; empty when nothing updatable changed. */
  props: { [prop: string]: unknown };
  /** OData type for Field.update(). */
  fieldType: string;
  /** Changes that cannot be applied to an existing field (type, hidden, indexed …). */
  notUpdatable: string[];
}

/**
 * MERGE properties that bring `target` in line with the (resolved) template field for the given changes.
 * Choices are merged, never removed: existing items may use them.
 */
export function fieldUpdate(resolved: IField, target: IFieldInfoLike, changes: string[]): IFieldUpdate {
  const props: { [prop: string]: unknown } = {};
  if (changes.indexOf('title') >= 0) props.Title = resolved.title;
  if (changes.indexOf('group') >= 0) props.Group = resolved.group || '';
  if (changes.indexOf('description') >= 0) props.Description = resolved.description || '';
  if (changes.indexOf('required') >= 0) props.Required = !!resolved.required;
  if (changes.indexOf('defaultValue') >= 0) props.DefaultValue = resolved.defaultValue === undefined ? null : resolved.defaultValue;
  if (changes.indexOf('customFormatter') >= 0) props.CustomFormatter = resolved.customFormatter || '';
  if (changes.indexOf('choices') >= 0) props.Choices = union(toFieldDef(target, target.SchemaXml).choices, resolved.choices);
  return {
    props,
    fieldType: `SP.Field${resolved.type === 'Choice' || resolved.type === 'MultiChoice' ? resolved.type : ''}`,
    notUpdatable: changes.filter((c) => UPDATABLE_FIELD_CHANGES.indexOf(c) < 0)
  };
}
