import type { IArtifactRef } from '../model';

export type LogLevel = 'info' | 'warn' | 'error';

export interface ILogEntry {
  /** ISO 8601 */
  time: string;
  level: LogLevel;
  artifact?: IArtifactRef;
  step?: string;
  /** Stable, English, developer-facing text; `code` (if any) is what the UI localizes. */
  message: string;
  code?: string;
  detail?: unknown;
}

export type LogListener = (entry: ILogEntry) => void;

export type LogCounts = Record<LogLevel, number>;

/** In-memory, subscribable run log with per-level counters and CSV/JSON export. */
export class Logger {
  private readonly _entries: ILogEntry[] = [];
  private readonly _listeners: LogListener[] = [];
  private readonly _counts: LogCounts = { info: 0, warn: 0, error: 0 };
  private readonly _now: () => Date;

  constructor(now: () => Date = () => new Date()) {
    this._now = now;
  }

  public get entries(): ReadonlyArray<ILogEntry> {
    return this._entries;
  }

  public get counts(): Readonly<LogCounts> {
    return this._counts;
  }

  public log(level: LogLevel, message: string, extra: Partial<Omit<ILogEntry, 'time' | 'level' | 'message'>> = {}): ILogEntry {
    const entry: ILogEntry = { time: this._now().toISOString(), level, message, ...extra };
    this._entries.push(entry);
    this._counts[level]++;
    this._listeners.slice().forEach((l) => {
      try {
        l(entry);
      } catch {
        // A broken UI listener must not break the run being logged.
      }
    });
    return entry;
  }

  public info(message: string, extra?: Partial<Omit<ILogEntry, 'time' | 'level' | 'message'>>): ILogEntry {
    return this.log('info', message, extra);
  }

  public warn(message: string, extra?: Partial<Omit<ILogEntry, 'time' | 'level' | 'message'>>): ILogEntry {
    return this.log('warn', message, extra);
  }

  public error(message: string, extra?: Partial<Omit<ILogEntry, 'time' | 'level' | 'message'>>): ILogEntry {
    return this.log('error', message, extra);
  }

  /** Returns an unsubscribe function. */
  public subscribe(listener: LogListener): () => void {
    this._listeners.push(listener);
    return () => {
      const i = this._listeners.indexOf(listener);
      if (i >= 0) {
        this._listeners.splice(i, 1);
      }
    };
  }

  public toJson(): string {
    return JSON.stringify(this._entries, null, 2);
  }

  /** RFC 4180 CSV with a header row; `detail` is serialized as JSON. */
  public toCsv(): string {
    const header = ['time', 'level', 'artifactKind', 'artifactKey', 'step', 'code', 'message', 'detail'];
    const rows = this._entries.map((e) => [
      e.time,
      e.level,
      e.artifact ? e.artifact.kind : '',
      e.artifact ? e.artifact.key : '',
      e.step || '',
      e.code || '',
      e.message,
      e.detail === undefined ? '' : safeJson(e.detail)
    ]);
    return [header].concat(rows).map((r) => r.map(csvCell).join(',')).join('\r\n');
  }
}

function csvCell(value: string): string {
  return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

function safeJson(value: unknown): string {
  if (value instanceof Error) {
    return JSON.stringify({ name: value.name, message: value.message });
  }
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}
