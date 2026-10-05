import { SPQueryable, spPost, type SPFI } from '@pnp/sp';
import { body } from '@pnp/queryable';
import '@pnp/sp/webs';
import { AbortError, throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';

/**
 * Server-side copy within one tenant (spike 16): Site.CreateCopyJobs on the target site with absolute source
 * URLs, one job per URL, each followed with Site.GetCopyJobProgress until the import's JobEnd. Never
 * overwrites: NameConflictBehavior 0 (Fail) makes an existing target an error of that file only.
 */

export interface ICopyJobInfo {
  EncryptionKey?: string;
  JobId: string;
  JobQueueUri: string;
  SourceListItemUniqueIds?: string[];
}

interface ICopyJobLog {
  Event?: string;
  MigrationDirection?: string;
  ErrorCode?: string;
  Message?: string;
}

export type CopyStatus = 'copied' | 'exists' | 'failed';

export interface ICopyOutcome {
  sourceUrl: string;
  status: CopyStatus;
  message?: string;
}

export interface ICopyOptions {
  /** Earlier versions go along (IgnoreVersionHistory false). */
  includeVersions: boolean;
  signal?: AbortSignal;
  /** Wait between progress requests (default 2 s). */
  pollMs?: number;
  /** A job not finished after this long counts as failed (default 10 min). */
  timeoutMs?: number;
}

/** "A file or folder with this name already exists at the destination" (spike 16, JobError). */
export const COPY_ALREADY_EXISTS = '-2147024713';

/** URLs per CreateCopyJobs request; they share the destination folder. */
export const COPY_BATCH = 50;

/** Progress requests in flight at once. */
const POLLING = 4;

/** Progress requests answering JobState 0 without the import's end before the job counts as over. */
const IDLE_POLLS = 15;

const webUrl = (sp: SPFI): string => sp.web.toUrl().replace(/\/_api\/web\/?$/i, '');
const q = (sp: SPFI, path: string): ReturnType<typeof SPQueryable> => SPQueryable([sp.web, `${webUrl(sp)}/_api/${path}`]);
const asArray = <T>(r: unknown): T[] => (Array.isArray(r) ? (r as T[]) : ((r as { value?: T[] }) || {}).value || []);

function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  throwIfAborted(signal);
  return new Promise((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout> | undefined = undefined;
    const stop = (): void => {
      if (timer !== undefined) clearTimeout(timer);
      reject(new AbortError());
    };
    timer = setTimeout(() => {
      if (signal) signal.removeEventListener('abort', stop);
      resolve();
    }, ms);
    if (signal) signal.addEventListener('abort', stop);
  });
}

/** The outcome a job's log entries tell. `ended` is false while the import has not finished. */
export function copyJobOutcome(logs: ICopyJobLog[]): { ended: boolean; status: CopyStatus; message?: string } {
  const fatal = logs.filter((l) => /JobFatalError|JobCancel/i.test(l.Event || ''))[0];
  const error = logs.filter((l) => l.Event === 'JobError')[0];
  const ended = !!fatal || logs.some((l) => l.Event === 'JobEnd' && l.MigrationDirection === 'Import');
  if (fatal) return { ended, status: 'failed', message: fatal.Message || fatal.Event };
  if (error) return { ended, status: String(error.ErrorCode) === COPY_ALREADY_EXISTS ? 'exists' : 'failed', message: error.Message };
  return { ended, status: 'copied' };
}

/**
 * Copies files (absolute source URLs) into one target folder (absolute URL). The outcomes keep the input
 * order; a job that cannot be started or followed is 'failed' with the reason.
 */
export async function copyFiles(sp: SPFI, sourceUrls: string[], destinationUrl: string, opts: ICopyOptions): Promise<ICopyOutcome[]> {
  throwIfAborted(opts.signal);
  const out: ICopyOutcome[] = [];
  for (let i = 0; i < sourceUrls.length; i += COPY_BATCH) {
    const urls = sourceUrls.slice(i, i + COPY_BATCH);
    const jobs = asArray<ICopyJobInfo>(
      await spPost(
        q(sp, 'site/CreateCopyJobs'),
        body({
          exportObjectUris: urls,
          destinationUri: destinationUrl,
          options: {
            AllowSchemaMismatch: true,
            IgnoreVersionHistory: !opts.includeVersions,
            IsMoveMode: false,
            NameConflictBehavior: 0,
            IncludeItemPermissions: false,
            ExcludeChildren: false,
            SameWebCopyMoveOptimization: true
          }
        })
      )
    );
    const followed = await limitConcurrency(
      urls.map((url, n) => () => (jobs[n] ? follow(sp, jobs[n], opts) : Promise.resolve({ status: 'failed' as CopyStatus, message: 'No copy job was created.' }))),
      POLLING,
      opts.signal
    );
    followed.forEach((r, n) =>
      out.push(r.ok ? { sourceUrl: urls[n], ...r.value } : { sourceUrl: urls[n], status: 'failed', message: r.error instanceof Error ? r.error.message : String(r.error) })
    );
  }
  return out;
}

async function follow(sp: SPFI, job: ICopyJobInfo, opts: ICopyOptions): Promise<{ status: CopyStatus; message?: string }> {
  const logs: ICopyJobLog[] = [];
  const started = Date.now();
  let idle = 0;
  while (Date.now() - started < (opts.timeoutMs || 600000)) {
    await sleep(opts.pollMs === undefined ? 2000 : opts.pollMs, opts.signal);
    const p = await spPost<{ JobState?: number; Logs?: string[] }>(q(sp, 'site/GetCopyJobProgress'), body({ copyJobInfo: job }));
    (p.Logs || []).forEach((l) => {
      try {
        logs.push(JSON.parse(l) as ICopyJobLog);
      } catch {
        // not a JSON entry; nothing to learn from it
      }
    });
    const outcome = copyJobOutcome(logs);
    idle = p.JobState === 0 ? idle + 1 : 0;
    if ((p.JobState === 0 && outcome.ended) || idle >= IDLE_POLLS) {
      if (!outcome.ended) return { status: 'failed', message: 'The copy job ended without a result.' };
      return outcome.message === undefined ? { status: outcome.status } : { status: outcome.status, message: outcome.message };
    }
  }
  return { status: 'failed', message: 'The copy job did not finish in time.' };
}
