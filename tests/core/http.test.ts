import { spfi, type SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import { DefaultHeaders, DefaultInit } from '@pnp/sp/behaviors/defaults';
import { DefaultParse } from '@pnp/queryable';
import {
  CLIENT_TAG,
  CopyJetBehavior,
  fetchWithRetry,
  limitConcurrency,
  parseRetryAfter,
  runBatched,
  abortableSleep
} from '../../src/core/http';
import { AbortError } from '../../src/core/errors';

const res = (status: number, headers: Record<string, string> = {}, body = '{}'): Response =>
  new Response(status === 204 ? null : body, { status, headers: { 'Content-Type': 'application/json', ...headers } });

describe('parseRetryAfter', () => {
  it('reads delta-seconds and HTTP dates', () => {
    expect(parseRetryAfter('7')).toBe(7000);
    const now = Date.parse('2026-09-24T10:00:00Z');
    expect(parseRetryAfter('Thu, 24 Sep 2026 10:00:30 GMT', now)).toBe(30000);
    expect(parseRetryAfter(null)).toBeUndefined();
    expect(parseRetryAfter('soon')).toBeUndefined();
  });
});

describe('fetchWithRetry', () => {
  const sleeps: number[] = [];
  const sleep = (ms: number): Promise<void> => {
    sleeps.push(ms);
    return Promise.resolve();
  };
  beforeEach(() => (sleeps.length = 0));

  it('honours Retry-After on 429 and returns the successful response', async () => {
    const doFetch = jest.fn().mockResolvedValueOnce(res(429, { 'Retry-After': '3' })).mockResolvedValueOnce(res(200));
    const r = await fetchWithRetry(doFetch, { sleep });
    expect(r.status).toBe(200);
    expect(doFetch).toHaveBeenCalledTimes(2);
    expect(sleeps).toEqual([3000]);
  });

  it('backs off 1-2-4-8 s and gives up after 5 attempts, returning the last 503', async () => {
    const doFetch = jest.fn().mockImplementation(() => Promise.resolve(res(503)));
    const r = await fetchWithRetry(doFetch, { sleep });
    expect(r.status).toBe(503);
    expect(doFetch).toHaveBeenCalledTimes(5);
    expect(sleeps).toEqual([1000, 2000, 4000, 8000]);
  });

  it('does not retry other errors such as 404', async () => {
    const doFetch = jest.fn().mockResolvedValue(res(404));
    expect((await fetchWithRetry(doFetch, { sleep })).status).toBe(404);
    expect(doFetch).toHaveBeenCalledTimes(1);
  });

  it('retries network errors and rethrows the last one', async () => {
    const doFetch = jest.fn().mockRejectedValue(new TypeError('Failed to fetch'));
    await expect(fetchWithRetry(doFetch, { sleep, maxAttempts: 3 })).rejects.toThrow('Failed to fetch');
    expect(doFetch).toHaveBeenCalledTimes(3);
  });

  it('stops when aborted', async () => {
    const ac = new AbortController();
    const doFetch = jest.fn().mockImplementation(() => {
      ac.abort();
      return Promise.resolve(res(429));
    });
    await expect(fetchWithRetry(doFetch, {}, ac.signal)).rejects.toBeInstanceOf(AbortError);
    expect(doFetch).toHaveBeenCalledTimes(1);
  });

  it('abortableSleep rejects on abort', async () => {
    const ac = new AbortController();
    const p = abortableSleep(60000, ac.signal);
    ac.abort();
    await expect(p).rejects.toBeInstanceOf(AbortError);
  });
});

describe('CopyJetBehavior', () => {
  const realFetch = global.fetch;
  afterEach(() => (global.fetch = realFetch));

  it('replaces the PnPjs send step: retries 429, tags requests', async () => {
    const fetchMock = jest
      .fn()
      .mockResolvedValueOnce(res(429, { 'Retry-After': '0' }))
      .mockResolvedValueOnce(res(200, {}, JSON.stringify({ Title: 'Projekt' })));
    global.fetch = fetchMock as unknown as typeof fetch;

    const sp = spfi('https://contoso.sharepoint.com/sites/projekt').using(
      DefaultHeaders(),
      DefaultInit(),
      DefaultParse(),
      CopyJetBehavior({ sleep: () => Promise.resolve() })
    );
    const web = await sp.web.select('Title')();

    expect(web).toEqual({ Title: 'Projekt' });
    expect(fetchMock).toHaveBeenCalledTimes(2);
    const init = fetchMock.mock.calls[0][1] as RequestInit;
    expect((init.headers as Record<string, string>)['X-ClientService-ClientTag']).toBe(CLIENT_TAG);
  });
});

describe('limitConcurrency', () => {
  it('keeps at most N tasks in flight and preserves order', async () => {
    let inFlight = 0;
    let peak = 0;
    const tasks = Array.from({ length: 10 }, (_, i) => async () => {
      inFlight++;
      peak = Math.max(peak, inFlight);
      await new Promise((r) => setTimeout(r, 5));
      inFlight--;
      if (i === 3) throw new Error('boom');
      return i;
    });
    const results = await limitConcurrency(tasks, 4);
    expect(peak).toBe(4);
    expect(results.map((r) => (r.ok ? r.value : 'err'))).toEqual([0, 1, 2, 'err', 4, 5, 6, 7, 8, 9]);
  });

  it('handles an empty task list', async () => {
    expect(await limitConcurrency([], 4)).toEqual([]);
  });

  it('does not start new tasks after abort', async () => {
    const ac = new AbortController();
    const started: number[] = [];
    const tasks = Array.from({ length: 5 }, (_, i) => async () => {
      started.push(i);
      if (i === 0) ac.abort();
      return i;
    });
    await expect(limitConcurrency(tasks, 1, ac.signal)).rejects.toBeInstanceOf(AbortError);
    expect(started).toEqual([0]);
  });
});

describe('runBatched', () => {
  /** Minimal stand-in for SPFI.batched(): ops queue on the batch and settle when execute() runs. */
  function fakeSp(): { sp: SPFI; batchSizes: number[] } {
    const batchSizes: number[] = [];
    const sp = {
      batched: () => {
        const queue: Array<() => void> = [];
        const batchedSp = {
          enqueue: <T>(fn: () => T): Promise<T> =>
            new Promise<T>((resolve, reject) =>
              queue.push(() => {
                try {
                  resolve(fn());
                } catch (e) {
                  reject(e);
                }
              })
            )
        };
        const execute = async (): Promise<void> => {
          batchSizes.push(queue.length);
          queue.forEach((q) => q());
        };
        return [batchedSp, execute];
      }
    } as unknown as SPFI;
    return { sp, batchSizes };
  }

  type Enq = { enqueue: <T>(fn: () => T) => Promise<T> };

  it('splits into batches of at most 100 and settles each op', async () => {
    const { sp, batchSizes } = fakeSp();
    const ops = Array.from({ length: 250 }, (_, i) => (b: SPFI) =>
      (b as unknown as Enq).enqueue(() => {
        if (i === 120) throw new Error('item 120');
        return i * 2;
      })
    );
    const results = await runBatched(sp, ops);
    expect(batchSizes).toEqual([100, 100, 50]);
    expect(results).toHaveLength(250);
    expect(results[5]).toEqual({ ok: true, value: 10 });
    expect(results[120].ok).toBe(false);
  });

  it('caps a larger requested size at 100', async () => {
    const { sp, batchSizes } = fakeSp();
    const ops = Array.from({ length: 150 }, () => (b: SPFI) => (b as unknown as Enq).enqueue(() => 1));
    await runBatched(sp, ops, 500);
    expect(batchSizes).toEqual([100, 50]);
  });
});
