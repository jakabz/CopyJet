import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/context-info';
import { CopyJetError, throwIfAborted } from '../errors';
import { fetchWithRetry } from './retry';

/**
 * POST requests PnPjs cannot send as tested: its default headers override Content-Type with
 * application/json;charset=utf-8, but CSOM needs text/xml and some REST calls were only verified with
 * odata=verbose (spikes 02, 04). The request digest still comes from PnPjs and the retry policy is shared.
 */

export type FetchLike = (url: string, init: RequestInit) => Promise<Response>;

export interface IRawPost {
  /** Path after the web URL, e.g. "_api/web/getList('…')/RootFolder". */
  path: string;
  headers: { [name: string]: string };
  body: string;
}

/** Sends the request against the web of `sp` and returns the (ok) response; HTTP errors become RAW_HTTP_ERROR. */
export async function rawPost(sp: SPFI, request: IRawPost, signal?: AbortSignal, fetchImpl: FetchLike = fetch): Promise<Response> {
  throwIfAborted(signal);
  const info = await sp.web.getContextInfo();
  const url = `${info.WebFullUrl.replace(/\/+$/, '')}/${request.path.replace(/^\/+/, '')}`;
  const response = await fetchWithRetry(
    () =>
      fetchImpl(url, {
        method: 'POST',
        credentials: 'same-origin',
        headers: { ...request.headers, 'X-RequestDigest': String(info.FormDigestValue) },
        body: request.body,
        signal
      }),
    {},
    signal
  );
  if (!response.ok) {
    let detail = '';
    try {
      detail = (await response.text()).slice(0, 500);
    } catch {
      // body unavailable
    }
    throw new CopyJetError('RAW_HTTP_ERROR', `POST ${request.path} failed with HTTP ${response.status}.`, { status: response.status, detail });
  }
  return response;
}

/** Quotes a server-relative URL for use inside getList('…') / getFolderByServerRelativePath(decodedurl='…'). */
export function odataPathArg(serverRelativeUrl: string): string {
  return encodeURIComponent(serverRelativeUrl.replace(/'/g, "''"));
}

/**
 * Sets a list's content type order (first = default) on its root folder: REST MERGE with odata=verbose,
 * the form verified in spike 04.
 */
export async function setContentTypeOrder(sp: SPFI, listServerRelativeUrl: string, listContentTypeIds: string[], signal?: AbortSignal, fetchImpl?: FetchLike): Promise<void> {
  await rawPost(
    sp,
    {
      path: `_api/web/getList('${odataPathArg(listServerRelativeUrl)}')/RootFolder`,
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify({
        __metadata: { type: 'SP.Folder' },
        UniqueContentTypeOrder: {
          __metadata: { type: 'Collection(SP.ContentTypeId)' },
          results: listContentTypeIds.map((id) => ({ __metadata: { type: 'SP.ContentTypeId' }, StringValue: id }))
        }
      })
    },
    signal,
    fetchImpl
  );
}
