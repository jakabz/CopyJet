import type { SPFI } from '@pnp/sp';
import { diffPlan, runPlan, type ProviderMap } from '../../src/core/engine';
import { CopyJetError } from '../../src/core/errors';
import { Logger } from '../../src/core/logger';
import type { ArtifactKind, ConflictMode, IApplyResult, IDiffResult, IInstallContext, IProvider } from '../../src/core/model';
import type { IPlan, IPlanStep, StepDef } from '../../src/core/planner';
import { TokenContext } from '../../src/core/tokenizer';

const sp = {} as SPFI;

function step(kind: ArtifactKind, key: string, level: number, dependsOn: string[] = []): IPlanStep {
  return { ref: { kind, key }, def: { key } as unknown as StepDef, dependsOn, level };
}

function plan(levels: IPlanStep[][]): IPlan {
  return { levels, steps: ([] as IPlanStep[]).concat(...levels), excluded: [] };
}

/** a, b (level 0) → c depends on a, d depends on b (level 1) → e depends on c (level 2). */
function sample(): IPlan {
  return plan([
    [step('group', 'a', 0), step('siteField', 'b', 0)],
    [step('list', 'c', 1, ['a']), step('list', 'd', 1, ['b'])],
    [step('view', 'e', 2, ['c'])]
  ]);
}

interface IFake {
  calls: Array<{ key: string; mode: ConflictMode }>;
  providers: ProviderMap;
  maxInFlight: number;
}

/** Providers that answer from a table: key → outcome, 'throw' or a diff status. */
function fakeProviders(behaviour: { [key: string]: string } = {}, delayMs = 1): IFake {
  const fake: IFake = { calls: [], providers: {}, maxInFlight: 0 };
  let inFlight = 0;
  const p: IProvider<StepDef> = {
    kind: 'group',
    diff: async (_sp, def): Promise<IDiffResult> => {
      const key = (def as unknown as { key: string }).key;
      if (behaviour[key] === 'throw') throw new CopyJetError('BOOM', `diff ${key} failed`);
      return { ref: { kind: 'group', key }, status: (behaviour[key] as IDiffResult['status']) || 'same' };
    },
    apply: async (_sp, def, mode): Promise<IApplyResult> => {
      const key = (def as unknown as { key: string }).key;
      fake.calls.push({ key, mode });
      inFlight++;
      fake.maxInFlight = Math.max(fake.maxInFlight, inFlight);
      await new Promise((r) => setTimeout(r, delayMs));
      inFlight--;
      if (behaviour[key] === 'throw') throw new CopyJetError('BOOM', `${key} failed`, { key });
      return { ref: { kind: 'group', key }, outcome: (behaviour[key] as IApplyResult['outcome']) || 'created' };
    }
  };
  (['group', 'siteField', 'contentType', 'list', 'listField', 'view'] as ArtifactKind[]).forEach((k) => (fake.providers[k] = p));
  return fake;
}

function ctx(signal?: AbortSignal): IInstallContext {
  return { targetSiteUrl: 'https://x/sites/cel', tokens: new TokenContext(), log: new Logger(), signal };
}

describe('runPlan', () => {
  it('runs level by level and reports progress and counts', async () => {
    const fake = fakeProviders({ b: 'skipped', d: 'updated' });
    const progress: string[] = [];
    const r = await runPlan(sp, sample(), ctx(), { providers: fake.providers, onProgress: (e) => progress.push(`${e.done}/${e.total} ${e.ref.key}:${e.status}`) });
    expect(fake.calls.map((c) => c.key)).toEqual(['a', 'b', 'c', 'd', 'e']);
    expect(progress).toHaveLength(5);
    expect(progress[4]).toBe('5/5 e:created');
    expect(r.counts).toEqual({ created: 3, updated: 1, skipped: 1, failed: 0, blocked: 0, cancelled: 0 });
    expect(r.aborted).toBe(false);
  });

  it('blocks dependents of a failed step, keeps independent ones running and logs both', async () => {
    const fake = fakeProviders({ a: 'throw' });
    const c = ctx();
    const r = await runPlan(sp, sample(), c, { providers: fake.providers });
    expect(r.steps.map((s) => `${s.ref.key}:${s.status}`)).toEqual(['a:failed', 'b:created', 'c:blocked', 'd:created', 'e:blocked']);
    expect(r.steps[2].blockedBy).toBe('a');
    expect(r.steps[4].blockedBy).toBe('c');
    expect(c.log.entries.map((e) => e.code)).toEqual(['BOOM', 'STEP_BLOCKED', 'STEP_BLOCKED']);
    expect(fake.calls.map((x) => x.key)).toEqual(['a', 'b', 'd']);
  });

  it('respects the concurrency limit and per-artifact conflict modes', async () => {
    const wide = plan([[1, 2, 3, 4, 5, 6].map((n) => step('list', `l${n}`, 0))]);
    const fake = fakeProviders({}, 5);
    await runPlan(sp, wide, ctx(), { providers: fake.providers, concurrency: 2, mode: 'update', modes: { l3: 'rename' } });
    expect(fake.maxInFlight).toBe(2);
    expect(fake.calls.find((c) => c.key === 'l3')!.mode).toBe('rename');
    expect(fake.calls.find((c) => c.key === 'l1')!.mode).toBe('update');
  });

  it('stops on abort: finished steps kept, the rest cancelled', async () => {
    const ac = new AbortController();
    const fake = fakeProviders();
    const r = await runPlan(sp, sample(), ctx(ac.signal), {
      providers: fake.providers,
      onProgress: (e) => {
        if (e.ref.key === 'b') ac.abort();
      }
    });
    expect(r.aborted).toBe(true);
    expect(r.steps.map((s) => `${s.ref.key}:${s.status}`)).toEqual(['a:created', 'b:created', 'c:cancelled', 'd:cancelled', 'e:cancelled']);
  });

  it('fails a step whose kind has no provider', async () => {
    const r = await runPlan(sp, plan([[step('page', 'p', 0)]]), ctx(), { providers: {} });
    expect(r.steps[0].status).toBe('failed');
    expect((r.steps[0].error as CopyJetError).code).toBe('NO_PROVIDER');
  });
});

describe('diffPlan', () => {
  it('marks steps depending on a new artifact as new without asking the target, and captures errors', async () => {
    const fake = fakeProviders({ a: 'new', b: 'throw', d: 'different' });
    const r = await diffPlan(sp, sample(), ctx(), fake.providers);
    expect(r.map((x) => `${x.ref.key}:${x.status}`)).toEqual(['a:new', 'b:error', 'c:new', 'd:different', 'e:new']);
    expect((r[1].error as CopyJetError).code).toBe('BOOM');
  });
});
