import Ajv2020, { type ErrorObject, type ValidateFunction } from 'ajv/dist/2020';
import addFormats from 'ajv-formats';
import addErrors from 'ajv-errors';
import { copyJetSchemaV1 } from './generated/schemaV1';

/** Which document of a template package to validate. */
export type SchemaTarget = 'manifest' | 'itemsFile' | 'filesMetaFile' | 'pageFile';

export interface IValidationIssue {
  /** JSON Pointer of the offending value, '' for the document root. */
  path: string;
  keyword: string;
  message: string;
  params: Record<string, unknown>;
}

export interface IValidationResult {
  valid: boolean;
  errors: IValidationIssue[];
}

const SCHEMA_ID = 'https://copyjet.local/schema/copyjet.v1.schema.json';

const REFS: Record<SchemaTarget, string> = {
  manifest: SCHEMA_ID,
  itemsFile: `${SCHEMA_ID}#/$defs/itemsFile`,
  filesMetaFile: `${SCHEMA_ID}#/$defs/filesMetaFile`,
  pageFile: `${SCHEMA_ID}#/$defs/pageFile`
};

let ajv: Ajv2020 | undefined;
const compiled: Partial<Record<SchemaTarget, ValidateFunction>> = {};

function getValidator(target: SchemaTarget): ValidateFunction {
  const cached = compiled[target];
  if (cached) {
    return cached;
  }
  if (!ajv) {
    // allErrors is required by ajv-errors. The schema's conditionals ("if": {properties} / "then": {required})
    // are valid 2020-12 but trip strictTypes/strictRequired, so only those two strict checks are relaxed.
    ajv = new Ajv2020({ allErrors: true, strict: true, strictTypes: false, strictRequired: false });
    addFormats(ajv);
    addErrors(ajv);
    ajv.addSchema(copyJetSchemaV1);
  }
  const fn = ajv.getSchema(REFS[target]);
  if (!fn) {
    throw new Error(`Schema reference not found: ${REFS[target]}`);
  }
  compiled[target] = fn;
  return fn;
}

function toIssue(e: ErrorObject): IValidationIssue {
  return {
    path: e.instancePath,
    keyword: e.keyword,
    message: e.message ?? e.keyword,
    params: e.params as Record<string, unknown>
  };
}

/**
 * Validates a document against the v1 schema. Failed `if` branches are reported by AJV as an extra
 * error next to the real cause; they carry no information, so they are dropped.
 */
export function validate(target: SchemaTarget, json: unknown): IValidationResult {
  const fn = getValidator(target);
  const valid = fn(json) as boolean;
  const errors = valid ? [] : (fn.errors ?? []).filter((e) => e.keyword !== 'if').map(toIssue);
  return { valid, errors };
}

export function validateTemplate(json: unknown): IValidationResult {
  return validate('manifest', json);
}

/** One line per issue, for logs and developer diagnostics. */
export function formatIssues(issues: IValidationIssue[]): string {
  return issues.map((i) => `${i.path || '/'}: ${i.message}`).join('\n');
}
