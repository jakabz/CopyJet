import { COPY_ALREADY_EXISTS, copyFiles, copyJobOutcome } from '../../src/core/files';
import { createMockSp } from '../helpers/mockSp';

const TARGET = 'https://contoso.sharepoint.com/sites/Cel';
const SRC = 'https://contoso.sharepoint.com/sites/Forras/Shared Documents';

const importEnd = { Event: 'JobEnd', MigrationDirection: 'Import' };

describe('copyJobOutcome', () => {
  it('waits for the import, not the export', () => {
    expect(copyJobOutcome([{ Event: 'JobStart' }, { Event: 'JobEnd', MigrationDirection: 'Export' }]).ended).toBe(false);
    expect(copyJobOutcome([{ Event: 'JobEnd', MigrationDirection: 'Export' }, importEnd])).toEqual({ ended: true, status: 'copied' });
  });

  it('tells an existing target from other errors', () => {
    expect(copyJobOutcome([{ Event: 'JobError', ErrorCode: COPY_ALREADY_EXISTS, Message: 'Már létezik' }, importEnd])).toEqual({ ended: true, status: 'exists', message: 'Már létezik' });
    expect(copyJobOutcome([{ Event: 'JobError', ErrorCode: '-1', Message: 'Hozzáférés megtagadva' }, importEnd])).toMatchObject({ status: 'failed', message: 'Hozzáférés megtagadva' });
    expect(copyJobOutcome([{ Event: 'JobFatalError', Message: 'Nem érhető el' }])).toEqual({ ended: true, status: 'failed', message: 'Nem érhető el' });
  });
});

describe('copyFiles', () => {
  it('starts one job per URL with Fail on conflict and follows each until the import ends', async () => {
    const progress: { [job: string]: string[][] } = {
      j0: [[JSON.stringify({ Event: 'JobStart' })], [JSON.stringify({ Event: 'JobFinishedObjectInfo' }), JSON.stringify(importEnd)]],
      j1: [[JSON.stringify({ Event: 'JobError', ErrorCode: COPY_ALREADY_EXISTS, Message: 'Már létezik' }), JSON.stringify(importEnd)]]
    };
    const { sp, requests } = createMockSp((req) => {
      if (req.method === 'POST' && /\/_api\/site\/CreateCopyJobs$/i.test(req.url)) return { body: { value: [{ JobId: 'j0', JobQueueUri: 'q' }, { JobId: 'j1', JobQueueUri: 'q' }] } };
      if (req.method === 'POST' && /\/_api\/site\/GetCopyJobProgress$/i.test(req.url)) {
        const id = (req.body as { copyJobInfo: { JobId: string } }).copyJobInfo.JobId;
        return { body: { JobState: progress[id].length > 1 ? 4 : 0, Logs: progress[id].shift() || [] } };
      }
      return undefined;
    }, TARGET);

    const out = await copyFiles(sp, [`${SRC}/a.docx`, `${SRC}/b.docx`], `${TARGET}/Shared Documents`, { includeVersions: false, pollMs: 0 });
    expect(out).toEqual([
      { sourceUrl: `${SRC}/a.docx`, status: 'copied' },
      { sourceUrl: `${SRC}/b.docx`, status: 'exists', message: 'Már létezik' }
    ]);
    const create = requests.filter((r) => /CreateCopyJobs/.test(r.url))[0].body as { exportObjectUris: string[]; destinationUri: string; options: { [k: string]: unknown } };
    expect(create.exportObjectUris).toEqual([`${SRC}/a.docx`, `${SRC}/b.docx`]);
    expect(create.destinationUri).toBe(`${TARGET}/Shared Documents`);
    expect(create.options).toMatchObject({ NameConflictBehavior: 0, IgnoreVersionHistory: true, IsMoveMode: false });
    expect(requests.filter((r) => /GetCopyJobProgress/.test(r.url)).length).toBe(3);
  });

  it('reports a job that never ends as failed', async () => {
    const { sp } = createMockSp((req) => {
      if (/CreateCopyJobs$/i.test(req.url)) return { body: { value: [{ JobId: 'j0', JobQueueUri: 'q' }] } };
      if (/GetCopyJobProgress$/i.test(req.url)) return { body: { JobState: 0, Logs: [] } };
      return undefined;
    }, TARGET);
    const out = await copyFiles(sp, [`${SRC}/a.docx`], `${TARGET}/Shared Documents`, { includeVersions: true, pollMs: 0 });
    expect(out[0]).toMatchObject({ status: 'failed' });
  });
});
