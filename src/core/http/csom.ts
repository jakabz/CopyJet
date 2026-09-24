import type { SPFI } from '@pnp/sp';
import { CopyJetError } from '../errors';
import { CREATE_CT_QUERY_ID, createContentTypeBody, fieldLinksBody, wrapRequest, type ICsomBody, type IFieldLinkAdd, type IFieldLinkFlags } from './csomXml';
import { rawPost, type FetchLike } from './raw';

export type { FetchLike } from './raw';

/**
 * Minimal CSOM (client.svc/ProcessQuery) client for what REST cannot do (spike 02): creating a content type
 * with a given ID, and adding field links / setting their flags.
 */

interface ICsomErrorInfo {
  ErrorMessage: string;
  ErrorCode: number;
  ErrorTypeName: string;
}

/**
 * Posts a CSOM request and returns the response array. A CSOM-level failure arrives as HTTP 200 with
 * ErrorInfo in the first element; it is raised as CopyJetError CSOM_ERROR.
 */
export async function processQuery(sp: SPFI, body: ICsomBody, signal?: AbortSignal, fetchImpl?: FetchLike): Promise<unknown[]> {
  let response: Response;
  try {
    response = await rawPost(
      sp,
      { path: '_vti_bin/client.svc/ProcessQuery', headers: { 'Content-Type': 'text/xml', Accept: 'application/json' }, body: wrapRequest(body) },
      signal,
      fetchImpl
    );
  } catch (e) {
    if (e instanceof CopyJetError && e.code === 'RAW_HTTP_ERROR') {
      throw new CopyJetError('CSOM_HTTP_ERROR', e.message, e.detail);
    }
    throw e;
  }
  const result = (await response.json()) as unknown[];
  const head = (result && result[0]) as { ErrorInfo?: ICsomErrorInfo } | undefined;
  if (head && head.ErrorInfo) {
    throw new CopyJetError('CSOM_ERROR', head.ErrorInfo.ErrorMessage, head.ErrorInfo);
  }
  return result;
}

export interface ICsomContentTypeInfo {
  id: string;
  name: string;
  group?: string;
  description?: string;
}

/**
 * Creates a content type on the web with exactly the given ID (the parent is implied by the ID).
 * Returns the StringId SharePoint reports for the new type.
 */
export async function csomCreateContentType(sp: SPFI, ct: ICsomContentTypeInfo, signal?: AbortSignal, fetchImpl?: FetchLike): Promise<string> {
  const result = await processQuery(sp, createContentTypeBody(ct), signal, fetchImpl);
  // Response: [header, <action id>, <value>, ...]; the query's value follows its id.
  const i = result.indexOf(CREATE_CT_QUERY_ID);
  const created = i >= 0 ? (result[i + 1] as { StringId?: string }) : undefined;
  if (!created || !created.StringId) {
    throw new CopyJetError('CSOM_UNEXPECTED_RESPONSE', 'ProcessQuery did not return the new content type ID.', result);
  }
  return created.StringId;
}

/** Adds field links and sets Required/Hidden on new and existing links of one content type, in one request. */
export async function csomUpdateFieldLinks(
  sp: SPFI,
  contentTypeId: string,
  add: IFieldLinkAdd[],
  flags: IFieldLinkFlags[],
  signal?: AbortSignal,
  fetchImpl?: FetchLike
): Promise<void> {
  if (add.length === 0 && flags.length === 0) {
    return;
  }
  await processQuery(sp, fieldLinksBody(contentTypeId, add, flags), signal, fetchImpl);
}
