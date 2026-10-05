import { Logger } from '../../src/core/logger';
import type { ICopyJetTemplate, IInstallContext } from '../../src/core/model';
import { finishedSteps, MemoryRunStore, restoreRunState, RunTracker, SpRunStore, templateChecksum, type IRunState, type IRunStore } from '../../src/core/state';
import { createMockSp } from '../helpers/mockSp';
import { TokenContext } from '../../src/core/tokenizer';

function ctx(): IInstallContext {
  const tokens = TokenContext.forSite({ absoluteUrl: 'https://fabrikam.sharepoint.com/sites/Cel', serverRelativeUrl: '/sites/Cel', title: 'Cél' });
  return { targetSiteUrl: 'https://fabrikam.sharepoint.com/sites/Cel', tokens, log: new Logger(), content: { reader: {} as never, principals: {} as never, idMaps: {} } };
}

const ref = (key: string) => ({ kind: 'list' as const, key });

describe('templateChecksum', () => {
  it('is a stable SHA-256 hex of the manifest', async () => {
    const t = (name: string, checksum?: string) => ({ schemaVersion: '1.1', meta: { name, checksum } }) as unknown as ICopyJetTemplate;
    const a = await templateChecksum(t('A'));
    expect(a).toMatch(/^[0-9a-f]{64}$/);
    expect(await templateChecksum(t('A', 'sha256:x'))).toBe(a);
    expect(await templateChecksum(t('B'))).not.toBe(a);
  });
});

describe('RunTracker', () => {
  it('saves after every step with the tokens and ID maps of the context', async () => {
    const store = new MemoryRunStore();
    const c = ctx();
    const t = RunTracker.start(store, c, 'Projekt', 'abc');
    c.tokens.set('listkey', 'L1', '11111111-1111-1111-1111-111111111111');
    c.content!.idMaps.L1 = { 5: 105 };
    c.content!.pendingItems = { L2: [1, 2] };
    await t.stepDone({ ref: ref('list:L1'), level: 0, status: 'created' });
    const saved = store.runs[t.state.runId];
    expect(saved.status).toBe('running');
    expect(saved.steps).toEqual({ 'list:L1': 'created' });
    expect(saved.tokens.find((x) => x.name === 'listkey')!.value).toBe('11111111-1111-1111-1111-111111111111');
    expect(saved.idMaps).toEqual({ L1: { 5: 105 } });
    expect(saved.pendingItems).toEqual({ L2: [1, 2] });
    await t.finish('completed');
    expect(store.runs[t.state.runId].status).toBe('completed');
    expect(await store.findUnfinished('abc')).toBeUndefined();
  });

  it('serializes parallel saves and writes the latest state last', async () => {
    const written: string[][] = [];
    let release: () => void = () => undefined;
    const slow: IRunStore = {
      findUnfinished: async () => undefined,
      save: async (s: IRunState) => {
        if (!written.length) await new Promise<void>((r) => (release = r));
        written.push(Object.keys(s.steps));
      }
    };
    const t = RunTracker.start(slow, ctx(), 'P', 'x');
    const p1 = t.stepDone({ ref: ref('a'), level: 0, status: 'created' });
    const p2 = t.stepDone({ ref: ref('b'), level: 0, status: 'created' });
    const p3 = t.stepDone({ ref: ref('c'), level: 0, status: 'skipped' });
    release();
    await Promise.all([p1, p2, p3]);
    expect(written).toEqual([['a'], ['a', 'b', 'c']]);
  });
});

describe('resume helpers', () => {
  const earlier: IRunState = {
    runId: 'r1',
    templateName: 'P',
    checksum: 'abc',
    status: 'aborted',
    started: '2026-10-05T07:00:00Z',
    updated: '2026-10-05T07:05:00Z',
    steps: { a: 'created', b: 'failed', c: 'skipped', d: 'cancelled' },
    tokens: [
      { name: 'site', value: 'https://old/sites/x' },
      { name: 'listkey', arg: 'L1', value: '22222222-2222-2222-2222-222222222222' }
    ],
    idMaps: { L1: { 1: 7 } },
    pendingItems: { L2: [3] }
  };

  it('offers the latest unfinished run of the same template', async () => {
    const store = new MemoryRunStore();
    await store.save(earlier);
    await store.save({ ...earlier, runId: 'r2', updated: '2026-10-05T08:00:00Z' });
    await store.save({ ...earlier, runId: 'r3', updated: '2026-10-05T09:00:00Z', status: 'completed' });
    await store.save({ ...earlier, runId: 'r4', checksum: 'other', updated: '2026-10-05T10:00:00Z' });
    expect((await store.findUnfinished('abc'))!.runId).toBe('r2');
  });

  it('skips only finished steps and restores tokens and ID maps without the site tokens', () => {
    expect(finishedSteps(earlier)).toEqual({ a: 'created', c: 'skipped' });
    const c = ctx();
    restoreRunState(earlier, c);
    expect(c.tokens.get('listkey', 'L1')).toBe('22222222-2222-2222-2222-222222222222');
    expect(c.tokens.get('site')).toBe('https://fabrikam.sharepoint.com/sites/Cel');
    expect(c.content!.idMaps).toEqual({ L1: { 1: 7 } });
    expect(c.content!.pendingItems).toEqual({ L2: [3] });
  });

  it('keeps the run ID when a run is resumed', () => {
    const t = RunTracker.start(new MemoryRunStore(), ctx(), 'P', 'abc', earlier);
    expect(t.state.runId).toBe('r1');
    expect(t.state.status).toBe('running');
  });
});

describe('SpRunStore', () => {
  /** A CopyJetLog list in memory: 404 until created, then items with a state.json attachment. */
  function fakeSite() {
    let exists = false;
    const fields: string[] = ['Title', 'RunId', 'Template', 'Checksum', 'Status'];
    const items: Array<{ Id: number; fields: Record<string, unknown>; file?: unknown }> = [];
    const { sp, requests } = createMockSp((req) => {
      const u = req.url;
      const L = "/_api/web/lists/getByTitle('CopyJetLog')";
      if (/\/_api\/web\/lists$/.test(u) && req.method === 'POST') {
        exists = true;
        return { body: { Id: 'l1' } };
      }
      if (u.indexOf(L) < 0) {
        if (/associatedOwnerGroup/i.test(u)) return { body: { Id: 3 } };
        if (/associatedMemberGroup|associatedVisitorGroup/i.test(u)) return { body: { Id: 4 } };
        if (/roleDefinitions\/getByType\(5\)/i.test(u)) return { body: { Id: 1073741829 } };
        return undefined;
      }
      if (!exists) return { status: 404, body: { 'odata.error': { message: { value: 'List does not exist' } } } };
      const rest = u.slice(u.indexOf(L) + L.length);
      if (/^\?\$select=Id$/.test(rest)) return { body: { Id: 'l1' } };
      if (/^\/fields\/createFieldAsXml/i.test(rest)) {
        fields.push(/Name="(\w+)"/.exec((req.body as { parameters: { SchemaXml: string } }).parameters.SchemaXml)![1]);
        return { body: {} };
      }
      if (/^\/fields\?\$select=InternalName$/i.test(rest)) return { body: fields.map((InternalName) => ({ InternalName })) };
      if (/^\/breakRoleInheritance/i.test(rest) || /^\/roleAssignments\/addRoleAssignment/i.test(rest)) return { status: 204 };
      if (/^\/items$/.test(rest) && req.method === 'POST') {
        const it = { Id: items.length + 1, fields: req.body as Record<string, unknown> };
        items.push(it);
        return { body: { Id: it.Id } };
      }
      const byId = /^\/items\((?:getById\()?(\d+)\)?/i.exec(rest) || /^\/items\/getById\((\d+)\)/i.exec(rest);
      if (byId) {
        const it = items.find((x) => x.Id === Number(byId[1]))!;
        if (/attachmentFiles\/add/i.test(rest)) {
          it.file = req.rawBody;
          return { body: {} };
        }
        if (/\$value$/.test(rest) && req.method === 'POST') {
          it.file = req.rawBody;
          return { status: 204 };
        }
        if (/\$value$/.test(rest)) return { body: JSON.parse(String(it.file)) };
        if (req.method === 'POST') {
          Object.assign(it.fields, req.body);
          return { status: 204 };
        }
      }
      if (/^\/items\?/.test(rest)) {
        const checksum = /CJChecksum eq '([^']+)'/.exec(rest);
        const runId = /CJRunId eq '([^']+)'/.exec(rest);
        const hit = items.filter((x) => (checksum ? x.fields.CJChecksum === checksum[1] && x.fields.CJStatus !== 'completed' : x.fields.CJRunId === runId![1]));
        return { body: hit.slice(-1).map((x) => ({ Id: x.Id, ...x.fields })) };
      }
      return undefined;
    });
    return { sp, requests, items, fields, setExists: (v: boolean) => (exists = v) };
  }

  async function blobText(b: unknown): Promise<string> {
    return b instanceof Blob ? b.text() : String(b);
  }

  it('creates the hidden owners-only list on the first save, then overwrites state.json in place', async () => {
    const site = fakeSite();
    const store = new SpRunStore(site.sp);
    expect(await store.findUnfinished('abc')).toBeUndefined();
    const state: IRunState = { runId: 'r1', templateName: 'P', checksum: 'abc', status: 'running', started: 's', updated: 'u', steps: { a: 'created' }, tokens: [], idMaps: {} };
    await store.save(state);
    const create = site.requests.find((r) => /\/_api\/web\/lists$/.test(r.url))!;
    expect(create.body).toMatchObject({ Title: 'CopyJetLog', BaseTemplate: 100, Hidden: true, NoCrawl: true });
    const added = site.requests.filter((r) => /createFieldAsXml/i.test(r.url)).map((r) => (r.body as { parameters: { SchemaXml: string; Options: number } }).parameters);
    expect(added.map((p) => /Name="(\w+)"/.exec(p.SchemaXml)![1])).toEqual(['CJRunId', 'CJTemplate', 'CJChecksum', 'CJStatus']);
    expect(added.every((p) => p.Options === 8)).toBe(true);
    expect(site.requests.some((r) => /breakRoleInheritance\(copyroleassignments=false, ?clearsubscopes=true\)/i.test(r.url))).toBe(true);
    expect(site.requests.some((r) => /addRoleAssignment\(principalid=3, ?roledefid=1073741829\)/i.test(r.url))).toBe(true);
    expect(site.items[0].fields).toMatchObject({ CJRunId: 'r1', CJChecksum: 'abc', CJStatus: 'running' });
    site.items[0].file = await blobText(site.items[0].file);

    const before = site.requests.length;
    await store.save({ ...state, steps: { a: 'created', b: 'skipped' } });
    const second = site.requests.slice(before);
    expect(second.map((r) => r.method)).toEqual(['POST']); // only the PUT of state.json
    site.items[0].file = await blobText(site.items[0].file);

    const loaded = await new SpRunStore(site.sp).findUnfinished('abc');
    expect(loaded!.steps).toEqual({ a: 'created', b: 'skipped' });

    await store.save({ ...state, status: 'completed' });
    expect(site.items[0].fields.CJStatus).toBe('completed');
  });

  it('adds the CJ… columns to a CopyJetLog list 1.8.0.0 created under display names, before reading it', async () => {
    const site = fakeSite();
    site.setExists(true);
    expect(await new SpRunStore(site.sp).findUnfinished('abc')).toBeUndefined();
    expect(site.fields).toEqual(['Title', 'RunId', 'Template', 'Checksum', 'Status', 'CJRunId', 'CJTemplate', 'CJChecksum', 'CJStatus']);
    expect(site.requests.some((r) => /\/_api\/web\/lists$/.test(r.url))).toBe(false);
  });
});
