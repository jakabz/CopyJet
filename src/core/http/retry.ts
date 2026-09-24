import { AbortError, throwIfAborted } from '../errors';

export interface IRetryOptions {
  /** Total attempts including the first one. */
  maxAttempts: number;
  /** Backoff base in ms when the response has no Retry-After: base, 2*base, 4*base ... */
  baseDelayMs: number;
  /** Upper bound for a single wait, also applied to Retry-After. */
  maxDelayMs: number;
  retryStatuses: number[];
  /** Injected for tests. */
  sleep: (ms: number, signal?: AbortSignal) => Promise<void>;
  onRetry?: (attempt: number, waitMs: number, reason: string) => void;
}

export const DEFAULT_RETRY_OPTIONS: IRetryOptions = {
  maxAttempts: 5,
  baseDelayMs: 1000,
  maxDelayMs: 120000,
  retryStatuses: [429, 503, 504],
  sleep: abortableSleep
};

export function abortableSleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    if (signal && signal.aborted) {
      reject(new AbortError());
      return;
    }
    let timer: ReturnType<typeof setTimeout> | undefined = undefined;
    const onAbort = (): void => {
      clearTimeout(timer);
      reject(new AbortError());
    };
    timer = setTimeout(() => {
      if (signal) {
        signal.removeEventListener('abort', onAbort);
      }
      resolve();
    }, ms);
    if (signal) {
      signal.addEventListener('abort', onAbort);
    }
  });
}

/** Retry-After is either delta-seconds or an HTTP date. Returns ms, or undefined if absent/unparsable. */
// eslint-disable-next-line @rushstack/no-new-null -- takes Headers.get() output, which is null when absent
export function parseRetryAfter(value: string | null, now: number = Date.now()): number | undefined {
  if (!value) {
    return undefined;
  }
  const trimmed = value.trim();
  if (/^\d+$/.test(trimmed)) {
    return parseInt(trimmed, 10) * 1000;
  }
  const date = Date.parse(trimmed);
  return isNaN(date) ? undefined : Math.max(0, date - now);
}

function isAbort(err: unknown): boolean {
  return err instanceof AbortError || (typeof err === 'object' && err !== null && (err as Error).name === 'AbortError');
}

/**
 * Calls `doFetch` until it returns a non-retryable response or attempts run out.
 * Throttling (429/503/504) waits Retry-After when given, otherwise exponential backoff; network errors
 * back off the same way. The last retryable response is returned (not thrown) so the caller's parser
 * produces the usual HTTP error.
 */
export async function fetchWithRetry(
  doFetch: () => Promise<Response>,
  options: Partial<IRetryOptions> = {},
  signal?: AbortSignal
): Promise<Response> {
  const o: IRetryOptions = { ...DEFAULT_RETRY_OPTIONS, ...options };
  let attempt = 0;
  for (;;) {
    throwIfAborted(signal);
    attempt++;
    let wait: number;
    let reason: string;
    try {
      const response = await doFetch();
      if (o.retryStatuses.indexOf(response.status) < 0 || attempt >= o.maxAttempts) {
        return response;
      }
      const retryAfter = parseRetryAfter(response.headers.get('Retry-After'));
      wait = retryAfter !== undefined ? retryAfter : o.baseDelayMs * Math.pow(2, attempt - 1);
      reason = `HTTP ${response.status}`;
    } catch (err) {
      if (isAbort(err) || attempt >= o.maxAttempts) {
        throw err;
      }
      wait = o.baseDelayMs * Math.pow(2, attempt - 1);
      reason = err instanceof Error ? err.message : String(err);
    }
    wait = Math.min(wait, o.maxDelayMs);
    if (o.onRetry) {
      o.onRetry(attempt, wait, reason);
    }
    await o.sleep(wait, signal);
  }
}
