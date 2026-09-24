import { spfi, type SPFI } from '@pnp/sp';
import { DefaultHeaders, DefaultInit } from '@pnp/sp/behaviors/defaults';
import { DefaultParse, type Queryable } from '@pnp/queryable';

export interface IMockRequest {
  method: string;
  /** Decoded URL, e.g. https://x/sites/a/_api/web/fields?$filter=InternalName eq 'A' */
  url: string;
  body?: unknown;
}

export interface IMockResponse {
  status?: number;
  body?: unknown;
}

/** Returns a response for a request, or undefined to answer 404 (unexpected call). */
export type MockHandler = (req: IMockRequest) => IMockResponse | undefined;

/**
 * A real PnPjs SPFI whose send step is answered by `handler` instead of the network, so tests exercise
 * the exact REST calls the core makes. Every request is recorded in `requests`.
 */
export function createMockSp(handler: MockHandler, siteUrl: string = 'https://fabrikam.sharepoint.com/sites/cel'): { sp: SPFI; requests: IMockRequest[] } {
  const requests: IMockRequest[] = [];
  const mockSend = (instance: Queryable): Queryable => {
    instance.on.send.clear();
    instance.on.send(async (url: URL, init: RequestInit) => {
      const req: IMockRequest = {
        method: init.method || 'GET',
        // Query strings encode spaces as '+'; a literal plus arrives as %2B.
        url: decodeURIComponent(url.toString().replace(/\+/g, ' ')),
        body: typeof init.body === 'string' && init.body ? JSON.parse(init.body) : undefined
      };
      requests.push(req);
      const res = handler(req) || { status: 404, body: { 'odata.error': { message: { value: `Unmocked ${req.method} ${req.url}` } } } };
      const status = res.status || 200;
      return new Response(status === 204 ? null : JSON.stringify(res.body === undefined ? {} : res.body), {
        status,
        headers: { 'Content-Type': 'application/json' }
      });
    });
    return instance;
  };
  const sp = spfi(siteUrl).using(DefaultHeaders(), DefaultInit(), DefaultParse(), mockSend);
  return { sp, requests };
}
