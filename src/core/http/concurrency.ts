import { throwIfAborted } from '../errors';

export type Settled<T> = { ok: true; value: T } | { ok: false; error: unknown };

/**
 * Runs tasks with at most `limit` in flight; results keep the input order. A failing task does not stop
 * the others. On abort, tasks not yet started are not started and the promise rejects with AbortError.
 */
export function limitConcurrency<T>(
  tasks: Array<() => Promise<T>>,
  limit: number = 4,
  signal?: AbortSignal
): Promise<Array<Settled<T>>> {
  const results: Array<Settled<T>> = new Array(tasks.length);
  let next = 0;

  const worker = async (): Promise<void> => {
    while (next < tasks.length) {
      throwIfAborted(signal);
      const i = next++;
      try {
        results[i] = { ok: true, value: await tasks[i]() };
      } catch (error) {
        results[i] = { ok: false, error };
      }
    }
  };

  const workers: Array<Promise<void>> = [];
  for (let w = 0; w < Math.max(1, Math.min(limit, tasks.length)); w++) {
    workers.push(worker());
  }
  return Promise.all(workers).then(() => {
    throwIfAborted(signal);
    return results;
  });
}
