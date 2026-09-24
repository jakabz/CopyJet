import type { SPFI } from '@pnp/sp';
import '@pnp/sp/batching';
import { throwIfAborted } from '../errors';
import type { Settled } from './concurrency';

/** SharePoint accepts at most 100 requests in one $batch. */
export const MAX_BATCH_SIZE = 100;

/** One operation: chain calls on the batched root it receives, and return the resulting promise. */
export type BatchOp<T> = (batchedSp: SPFI) => Promise<T>;

/**
 * Splits `ops` into $batch requests of at most `size`, executed one after another.
 * Results keep the input order; one failing operation does not fail the others.
 */
export async function runBatched<T>(
  sp: SPFI,
  ops: Array<BatchOp<T>>,
  size: number = MAX_BATCH_SIZE,
  signal?: AbortSignal
): Promise<Array<Settled<T>>> {
  const chunk = Math.max(1, Math.min(size, MAX_BATCH_SIZE));
  const results: Array<Settled<T>> = [];
  for (let start = 0; start < ops.length; start += chunk) {
    throwIfAborted(signal);
    const [batchedSp, execute] = sp.batched();
    const pending = ops.slice(start, start + chunk).map((op) =>
      op(batchedSp).then(
        (value): Settled<T> => ({ ok: true, value }),
        (error): Settled<T> => ({ ok: false, error })
      )
    );
    await execute();
    const settled = await Promise.all(pending);
    for (let i = 0; i < settled.length; i++) {
      results.push(settled[i]);
    }
  }
  return results;
}
