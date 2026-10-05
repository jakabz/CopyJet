import type { IRunState, IRunStore } from './runState';

/** Run states in memory: for tests and for an install that cannot write to the target's log list. */
export class MemoryRunStore implements IRunStore {
  public readonly runs: { [runId: string]: IRunState } = {};
  public saves = 0;

  public async findUnfinished(checksum: string): Promise<IRunState | undefined> {
    return Object.keys(this.runs)
      .map((k) => this.runs[k])
      .filter((r) => r.checksum === checksum && r.status !== 'completed')
      .sort((a, b) => (a.updated < b.updated ? 1 : -1))[0];
  }

  public async save(state: IRunState): Promise<void> {
    this.saves++;
    this.runs[state.runId] = JSON.parse(JSON.stringify(state));
  }
}
