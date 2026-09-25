import type { SPFI } from '@pnp/sp';
import { AbortError, CopyJetError, isAbortError } from '../errors';
import { limitConcurrency } from '../http/concurrency';
import type { ArtifactKind, ConflictMode, DiffStatus, IArtifactRef, IInstallContext, IProvider } from '../model';
import type { IPlan, IPlanStep, StepDef } from '../planner';

export type ProviderMap = Partial<Record<ArtifactKind, IProvider<StepDef>>>;

export type StepStatus = 'created' | 'updated' | 'skipped' | 'failed' | 'blocked' | 'cancelled';

export interface IStepResult {
  ref: IArtifactRef;
  level: number;
  status: StepStatus;
  error?: unknown;
  /** For 'blocked': the failed/blocked dependency that stopped this step. */
  blockedBy?: string;
}

export interface IProgressEvent {
  done: number;
  total: number;
  ref: IArtifactRef;
  status: StepStatus;
}

export interface IRunOptions {
  providers: ProviderMap;
  /** Conflict mode for existing, different artifacts (default 'skip'); `modes` overrides per artifact key. */
  mode?: ConflictMode;
  modes?: { [key: string]: ConflictMode };
  concurrency?: number;
  onProgress?: (event: IProgressEvent) => void;
}

export interface IRunResult {
  steps: IStepResult[];
  counts: Record<StepStatus, number>;
  aborted: boolean;
}

const isAbort = isAbortError;

function provider(providers: ProviderMap, step: IPlanStep): IProvider<StepDef> {
  const p = providers[step.ref.kind];
  if (!p) {
    throw new CopyJetError('NO_PROVIDER', `No provider for ${step.ref.kind}.`);
  }
  return p;
}

/**
 * Runs a plan level by level (at most `concurrency` steps in flight within a level). A failed step blocks
 * its dependents; independent steps keep running. A step whose dependency was created in this run runs in
 * 'update' mode even when the mode is 'skip'. On abort, steps not started are 'cancelled' and the
 * partial result is returned. Providers register the identifiers they create in ctx.tokens themselves.
 */
export async function runPlan(sp: SPFI, plan: IPlan, ctx: IInstallContext, options: IRunOptions): Promise<IRunResult> {
  const results: { [key: string]: IStepResult } = {};
  const total = plan.steps.length;
  let done = 0;
  let aborted = false;

  const finish = (step: IPlanStep, status: StepStatus, extra: Partial<IStepResult> = {}): void => {
    results[step.ref.key] = { ref: step.ref, level: step.level, status, ...extra };
    done++;
    if (options.onProgress) options.onProgress({ done, total, ref: step.ref, status });
  };

  for (const level of plan.levels) {
    if (aborted || (ctx.signal && ctx.signal.aborted)) {
      aborted = true;
      break;
    }
    // Steps sharing a lock run in sequence; different locks run in parallel.
    const groups: IPlanStep[][] = [];
    const byLock: { [lock: string]: IPlanStep[] } = {};
    level.forEach((step) => {
      if (!byLock[step.lock]) groups.push((byLock[step.lock] = []));
      byLock[step.lock].push(step);
    });
    const runStep = async (step: IPlanStep): Promise<void> => {
      const failedDep = step.dependsOn.filter((d) => results[d] && ['failed', 'blocked', 'cancelled'].indexOf(results[d].status) >= 0)[0];
      if (failedDep) {
        ctx.log.warn(`Skipped because ${failedDep} did not install.`, { artifact: step.ref, code: 'STEP_BLOCKED', detail: failedDep });
        finish(step, 'blocked', { blockedBy: failedDep });
        return;
      }
      try {
        let mode = (options.modes && options.modes[step.ref.key]) || options.mode || 'skip';
        // Inside something this run created (e.g. the default view of a new list) the template decides:
        // nothing there predates the install, so 'skip' would only keep SharePoint's defaults.
        if (mode === 'skip' && step.dependsOn.some((d) => results[d] && results[d].status === 'created')) {
          mode = 'update';
        }
        const r = await provider(options.providers, step).apply(sp, step.def, mode, ctx);
        finish(step, r.outcome);
      } catch (e) {
        if (isAbort(e)) throw e;
        ctx.log.error(e instanceof Error ? e.message : String(e), {
          artifact: step.ref,
          code: e instanceof CopyJetError ? e.code : 'STEP_FAILED',
          detail: e instanceof CopyJetError ? e.detail : e
        });
        finish(step, 'failed', { error: e });
      }
    };
    try {
      await limitConcurrency(
        groups.map((group) => async () => {
          for (const step of group) {
            if (ctx.signal && ctx.signal.aborted) throw new AbortError();
            await runStep(step);
          }
        }),
        options.concurrency || 4,
        ctx.signal
      );
    } catch (e) {
      if (!isAbort(e)) throw e;
      aborted = true;
    }
  }

  plan.steps.filter((s) => !results[s.ref.key]).forEach((s) => (results[s.ref.key] = { ref: s.ref, level: s.level, status: 'cancelled' }));
  if (aborted) ctx.log.warn('Installation stopped by the user.', { code: 'RUN_ABORTED' });

  const steps = plan.steps.map((s) => results[s.ref.key]);
  const counts: Record<StepStatus, number> = { created: 0, updated: 0, skipped: 0, failed: 0, blocked: 0, cancelled: 0 };
  steps.forEach((s) => counts[s.status]++);
  return { steps, counts, aborted };
}

export type PreviewStatus = DiffStatus | 'error';

export interface IPreviewResult {
  ref: IArtifactRef;
  status: PreviewStatus;
  changes?: string[];
  error?: unknown;
}

/**
 * Diff of every step for the preview. A step whose dependency will be created ('new') is 'new' itself:
 * asking the target about it now would only report the missing parent (e.g. a column of a new list).
 */
export async function diffPlan(sp: SPFI, plan: IPlan, ctx: IInstallContext, providers: ProviderMap, concurrency: number = 4): Promise<IPreviewResult[]> {
  const results: { [key: string]: IPreviewResult } = {};
  for (const level of plan.levels) {
    await limitConcurrency(
      level.map((step) => async () => {
        if (step.dependsOn.some((d) => results[d] && results[d].status === 'new')) {
          results[step.ref.key] = { ref: step.ref, status: 'new' };
          return;
        }
        try {
          const d = await provider(providers, step).diff(sp, step.def, ctx);
          results[step.ref.key] = { ref: step.ref, status: d.status, changes: d.changes };
        } catch (e) {
          if (isAbort(e)) throw e;
          results[step.ref.key] = { ref: step.ref, status: 'error', error: e };
        }
      }),
      concurrency,
      ctx.signal
    );
  }
  return plan.steps.map((s) => results[s.ref.key]);
}
