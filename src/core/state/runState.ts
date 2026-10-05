import type { IStepResult, StepStatus } from '../engine/engine';
import { FINISHED_STATUSES } from '../engine/engine';
import type { ICopyJetTemplate, IInstallContext } from '../model';
import { manifestChecksumBytes, sha256Hex } from '../packager/checksum';
import type { ITokenEntry } from '../tokenizer/TokenContext';

export type RunStatus = 'running' | 'completed' | 'failed' | 'aborted';

/**
 * What a resume needs from an earlier run: the outcome of every step that ran, the identifiers the providers
 * registered (tokens) and the item ID maps. Steps finished by the earlier run are not run again, so their
 * tokens and ID maps must come from here.
 */
export interface IRunState {
  runId: string;
  templateName: string;
  /** Checksum of the template manifest: a resume is offered only for the same template. */
  checksum: string;
  status: RunStatus;
  started: string;
  updated: string;
  steps: { [key: string]: StepStatus };
  tokens: ITokenEntry[];
  idMaps: { [listKey: string]: { [sourceId: number]: number } };
  /** Item chunks being written when the state was saved (source IDs per list key). */
  pendingItems?: { [listKey: string]: number[] };
}

/** Where run states live (the target site's CopyJetLog list; in tests, memory). */
export interface IRunStore {
  /** The latest unfinished run of this template, if any. */
  findUnfinished(checksum: string, signal?: AbortSignal): Promise<IRunState | undefined>;
  save(state: IRunState): Promise<void>;
}

/** SHA-256 of the manifest as stored in the package (without meta.checksum): the identity of a template. */
export async function templateChecksum(manifest: ICopyJetTemplate): Promise<string> {
  return sha256Hex(manifestChecksumBytes(manifest));
}

export function newRunId(now: Date = new Date()): string {
  const rand = Math.random().toString(36).slice(2, 8);
  return `${now.toISOString().replace(/[-:]/g, '').replace(/\..*$/, '')}-${rand}`;
}

/** Steps of a state the engine may skip on resume. */
export function finishedSteps(state: IRunState): { [key: string]: StepStatus } {
  const out: { [key: string]: StepStatus } = {};
  Object.keys(state.steps).forEach((k) => {
    if (FINISHED_STATUSES.indexOf(state.steps[k]) >= 0) out[k] = state.steps[k];
  });
  return out;
}

/** Tokens read from the target or set by the mapping step of this session: never taken from an earlier run. */
const SESSION_TOKENS = ['site', 'siterelative', 'sitename', 'webid', 'siteid', 'associatedownergroup', 'associatedmembergroup', 'associatedvisitorgroup', 'principal'];

/** Puts an earlier run's identifiers (lists, fields, views, groups …) and ID maps into a fresh install context. */
export function restoreRunState(state: IRunState, ctx: IInstallContext): void {
  state.tokens.forEach((t) => {
    if (SESSION_TOKENS.indexOf(t.name) >= 0) return;
    try {
      ctx.tokens.set(t.name, t.arg, t.value);
    } catch {
      // A token this version no longer knows: the step that needs it reruns or fails visibly.
    }
  });
  if (ctx.content) {
    Object.keys(state.idMaps).forEach((k) => {
      ctx.content!.idMaps[k] = { ...state.idMaps[k], ...(ctx.content!.idMaps[k] || {}) };
    });
    if (state.pendingItems) ctx.content.pendingItems = { ...state.pendingItems };
  }
}

/**
 * Keeps the state of the running install and saves it after every step. Saves are serialized and coalesced:
 * while one is in flight, further steps only mark the state dirty, and one more save follows with the latest
 * snapshot – parallel steps never write over each other.
 */
export class RunTracker {
  public readonly state: IRunState;
  private _inFlight: Promise<void> | undefined;
  private _dirty = false;

  constructor(
    private readonly _store: IRunStore,
    private readonly _ctx: IInstallContext,
    state: IRunState
  ) {
    this.state = state;
  }

  public static start(store: IRunStore, ctx: IInstallContext, templateName: string, checksum: string, resumed?: IRunState): RunTracker {
    const now = new Date().toISOString();
    const state: IRunState = resumed
      ? { ...resumed, status: 'running', updated: now }
      : { runId: newRunId(), templateName, checksum, status: 'running', started: now, updated: now, steps: {}, tokens: [], idMaps: {} };
    return new RunTracker(store, ctx, state);
  }

  /** For runPlan's onStepDone. */
  public readonly stepDone = async (result: IStepResult): Promise<void> => {
    this.state.steps[result.ref.key] = result.status;
    await this._save();
  };

  /** For ctx.checkpoint: saves the ID maps written so far within a step. */
  public readonly checkpoint = async (): Promise<void> => {
    await this._save();
  };

  public async finish(status: RunStatus): Promise<void> {
    this.state.status = status;
    await this._save();
  }

  private async _save(): Promise<void> {
    if (this._inFlight) {
      this._dirty = true;
      return this._inFlight;
    }
    this._inFlight = (async () => {
      try {
        do {
          this._dirty = false;
          this._snapshot();
          // Not cancelled with the run: the state of a stopped install must still be written.
          await this._store.save(JSON.parse(JSON.stringify(this.state)));
        } while (this._dirty);
      } finally {
        this._inFlight = undefined;
      }
    })();
    return this._inFlight;
  }

  private _snapshot(): void {
    this.state.updated = new Date().toISOString();
    this.state.tokens = this._ctx.tokens.entries();
    if (this._ctx.content) {
      this.state.idMaps = this._ctx.content.idMaps;
      this.state.pendingItems = this._ctx.content.pendingItems;
    }
  }
}
