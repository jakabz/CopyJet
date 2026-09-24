import type { IArtifactRef, IField } from '../model';
import type { Logger } from '../logger/Logger';
import { tokenize, type TokenContext } from '../tokenizer';
import { toFieldDef, type IFieldInfoLike } from './fieldModel';
import { sanitizeFieldXml, setFieldXmlAttribute } from './schemaXml';

/**
 * Template definition of a source field: SchemaXml sanitized (allowlist) and tokenized, but the field's own
 * ID kept as-is – content types and site column instances on lists are matched by it.
 */
export function templateFieldFrom(info: IFieldInfoLike, tokens: TokenContext, log: Logger, ref: IArtifactRef): IField {
  const sanitized = sanitizeFieldXml(info.SchemaXml);
  if (sanitized.removedAttributes.length || sanitized.removedElements.length) {
    log.info('Removed source-specific parts of SchemaXml.', {
      artifact: ref,
      code: 'FIELD_XML_SANITIZED',
      detail: { attributes: sanitized.removedAttributes, elements: sanitized.removedElements }
    });
  }
  const originalId = /\bID="([^"]*)"/.exec(sanitized.xml);
  let xml = tokenize(sanitized.xml, tokens);
  if (originalId) {
    xml = setFieldXmlAttribute(xml, 'ID', originalId[1]);
  }
  return toFieldDef(info, xml);
}
