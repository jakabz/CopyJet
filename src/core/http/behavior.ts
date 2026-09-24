import type { TimelinePipe } from '@pnp/core';
import { InjectHeaders, type IQueryableInternal, type Queryable } from '@pnp/queryable';
import { COPYJET_VERSION } from '../version';
import { fetchWithRetry, type IRetryOptions } from './retry';

/** Traffic decoration recommended by Microsoft for throttling diagnostics. */
export const CLIENT_TAG = `NONISV|CopyJet|${COPYJET_VERSION}`;

/**
 * CopyJet's PnPjs behavior. Replaces the send step installed by SPFx() (BrowserFetchWithRetry: 3 tries,
 * 200 ms doubling) with fetchWithRetry (Retry-After aware, 5 tries, 1-2-4-8-16 s), and tags every request.
 * Apply after SPFx(context).
 */
export function CopyJetBehavior(retry: Partial<IRetryOptions> = {}): TimelinePipe<Queryable> {
  return (instance: Queryable) => {
    // Browsers ignore a script-set User-Agent; X-ClientService-ClientTag is what SharePoint logs instead.
    instance.using(InjectHeaders({ 'User-Agent': CLIENT_TAG, 'X-ClientService-ClientTag': CLIENT_TAG }));
    instance.on.send.clear();
    instance.on.send(function (this: IQueryableInternal, url: URL, init: RequestInit): Promise<Response> {
      const u = url.toString();
      return fetchWithRetry(
        () => fetch(u, init),
        {
          ...retry,
          onRetry: (attempt, waitMs, reason) => {
            this.log(`CopyJet retry #${attempt} in ${waitMs} ms (${reason}): ${init.method} ${u}`, 1);
            if (retry.onRetry) {
              retry.onRetry(attempt, waitMs, reason);
            }
          }
        },
        init.signal || undefined
      );
    });
    return instance;
  };
}
