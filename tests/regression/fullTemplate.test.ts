import { readFileSync, readdirSync } from 'fs';
import { join } from 'path';
import type { SPFI } from '@pnp/sp';
import { runPlan, type ProviderMap } from '../../src/core/engine';
import { Logger } from '../../src/core/logger';
import type { ArtifactKind, ICopyJetTemplate, IInstallContext, IProvider, ITemplateReader } from '../../src/core/model';
import { ZipTemplateWriter, openTemplate } from '../../src/core/packager';
import { buildPlan, templateNodes, type IPlan, type StepDef } from '../../src/core/planner';
import { finishedSteps, MemoryRunStore, restoreRunState, RunTracker, templateChecksum } from '../../src/core/state';
import { TokenContext } from '../../src/core/tokenizer';

/**
 * Regression suite: the whole install flow on the full example template (schema/examples) – package, open,
 * validate, plan, run, rerun and resume – with recording providers in place of SharePoint. The providers
 * themselves are covered by their own tests against REST mocks; this suite guards how the pieces fit.
 */

const ROOT = join(__dirname, '..', '..');
const json = <T>(...path: string[]): T => JSON.parse(readFileSync(join(ROOT, ...path), 'utf8')) as T;
const sp = {} as SPFI;

/** The example manifest with the content entries it points to. */
async function examplePackage(): Promise<ITemplateReader> {
  const writer = new ZipTemplateWriter(json<ICopyJetTemplate>('schema', 'examples', 'manifest.example.json'));
  writer.addJson('items/Ugyfelek.json', { listKey: 'Ugyfelek', items: [] });
  writer.addJson('items/Projektek.json', json('schema', 'examples', 'items.Projektek.example.json'));
  writer.addJson('files/Dokumentumok/_meta.json', { listKey: 'Dokumentumok', files: [] });
  writer.addJson('pages/Home.json', { name: 'Home.aspx', canvasContent: [] });
  writer.addBlob('assets/SiteAssets/fejlec.jpg', new Blob([new Uint8Array([1, 2, 3])]));
  writer.addBlob('attachments/Projektek/1/specifikacio.pdf', new Blob([new Uint8Array([4, 5])]));
  return openTemplate(await writer.finalize());
}

function context(signal?: AbortSignal): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: 'https://fabrikam.sharepoint.com/sites/Cel', serverRelativeUrl: '/sites/Cel', title: 'Cél' });
  return { targetSiteUrl: 'https://fabrikam.sharepoint.com/sites/Cel', tokens, log: new Logger(), signal };
}

interface IRecorder {
  providers: ProviderMap;
  applied: string[];
}

/** Providers for every kind of the plan: diff says 'new', apply records the key and creates. */
function recorder(plan: IPlan, onApply?: (count: number) => void): IRecorder {
  const rec: IRecorder = { providers: {}, applied: [] };
  const kinds = plan.steps.map((s) => s.ref.kind).filter((k, i, all) => all.indexOf(k) === i);
  kinds.forEach((kind: ArtifactKind) => {
    const p: IProvider<StepDef> = {
      kind,
      diff: async (_sp, def) => ({ ref: { kind, key: keyOf(plan, def) }, status: 'new' }),
      apply: async (_sp, def) => {
        const key = keyOf(plan, def);
        rec.applied.push(key);
        if (onApply) onApply(rec.applied.length);
        await new Promise((r) => setTimeout(r, 1));
        return { ref: { kind, key }, outcome: 'created' };
      }
    };
    rec.providers[kind] = p;
  });
  return rec;
}

const keyOf = (plan: IPlan, def: StepDef): string => plan.steps.filter((s) => s.def === def)[0].ref.key;

describe('full example template', () => {
  it('packs, opens and validates with a stable checksum', async () => {
    const a = await examplePackage();
    const b = await examplePackage();
    expect(a.manifest.meta.checksum).toMatch(/^sha256:[0-9a-f]{64}$/);
    expect(await templateChecksum(a.storedManifest)).toBe(await templateChecksum(b.storedManifest));
  });

  it('plans every artifact of the template, dependencies first', async () => {
    const reader = await examplePackage();
    const plan = buildPlan(reader.manifest);
    expect(plan.excluded).toEqual([]);
    expect(plan.steps.map((s) => s.ref.key).sort()).toEqual(templateNodes(reader.manifest).map((n) => n.ref.key).sort());
    const level: { [key: string]: number } = {};
    plan.steps.forEach((s) => (level[s.ref.key] = s.level));
    plan.steps.forEach((s) => s.dependsOn.forEach((d) => expect(level[d]).toBeLessThan(s.level)));
    // Kinds of the example: a change of the planned order shows up here.
    expect(plan.levels.map((l) => l.map((s) => s.ref.key))).toMatchSnapshot();
  });

  it('installs every step once; a rerun with the finished steps applies nothing', async () => {
    const reader = await examplePackage();
    const plan = buildPlan(reader.manifest);
    const store = new MemoryRunStore();
    const checksum = await templateChecksum(reader.storedManifest);
    const ctx = context();
    const tracker = RunTracker.start(store, ctx, reader.manifest.meta.name, checksum);
    const first = recorder(plan);
    const r = await runPlan(sp, plan, ctx, { providers: first.providers, onStepDone: tracker.stepDone });
    await tracker.finish('completed');
    expect(r.counts.created).toBe(plan.steps.length);
    expect(first.applied.slice().sort()).toEqual(plan.steps.map((s) => s.ref.key).sort());
    expect(await store.findUnfinished(checksum)).toBeUndefined();

    const again = recorder(plan);
    const r2 = await runPlan(sp, plan, context(), { providers: again.providers, previous: finishedSteps(tracker.state) });
    expect(again.applied).toEqual([]);
    expect(r2.counts.previous).toBe(plan.steps.length);
  });

  it('resumes an interrupted install from the saved state without repeating finished steps', async () => {
    const reader = await examplePackage();
    const plan = buildPlan(reader.manifest);
    const store = new MemoryRunStore();
    const checksum = await templateChecksum(reader.storedManifest);
    const ac = new AbortController();
    const ctx = context(ac.signal);
    ctx.tokens.set('listkey', 'Ugyfelek', '11111111-1111-4111-8111-111111111111');
    const tracker = RunTracker.start(store, ctx, reader.manifest.meta.name, checksum);
    const first = recorder(plan, (n) => n === 4 && ac.abort());
    const r = await runPlan(sp, plan, ctx, { providers: first.providers, onStepDone: tracker.stepDone });
    expect(r.aborted).toBe(true);
    await tracker.finish('aborted');

    const saved = await store.findUnfinished(checksum);
    expect(saved).toBeDefined();
    const done = finishedSteps(saved!);
    expect(Object.keys(done).length).toBeGreaterThan(0);
    expect(Object.keys(done).length).toBeLessThan(plan.steps.length);

    const ctx2 = context();
    restoreRunState(saved!, ctx2);
    expect(ctx2.tokens.get('listkey', 'Ugyfelek')).toBe('11111111-1111-4111-8111-111111111111');
    const second = recorder(plan);
    const r2 = await runPlan(sp, plan, ctx2, { providers: second.providers, previous: done });
    expect(r2.aborted).toBe(false);
    expect(second.applied.filter((k) => !!done[k])).toEqual([]);
    expect(Object.keys(done).concat(second.applied).sort()).toEqual(plan.steps.map((s) => s.ref.key).sort());
  });
});

describe('templates exported by the Setup (tests/fixtures/templates)', () => {
  const dir = join(__dirname, '..', 'fixtures', 'templates');
  const files = readdirSync(dir).filter((f) => /\.(json|zip)$/i.test(f));

  it.each(files)('%s opens, validates and plans without exclusions', async (file) => {
    const reader = await openTemplate(new Blob([readFileSync(join(dir, file))]));
    const plan = buildPlan(reader.manifest);
    expect(plan.excluded).toEqual([]);
    expect(plan.steps.length).toBe(templateNodes(reader.manifest).length);
  });
});
