import { throwIfAborted } from '../errors';
import { limitConcurrency } from '../http/concurrency';

const DAY = 24 * 60 * 60 * 1000;

/** "2026-03-02T09:15:00" (no zone) or an ISO UTC string → epoch ms, reading the wall-clock digits as UTC. */
function wallClockMs(value: string): number {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/.exec(value);
  if (!m) throw new Error(`Not an ISO date-time: ${value}`);
  return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6] || 0));
}

const toWallClock = (ms: number): string => new Date(ms).toISOString().slice(0, 19);

/**
 * UTC → the target web's local time, via SharePoint's own time zone (utcToLocalTime handles daylight saving,
 * spike 08 C/D; TimeZone.Information has no transition dates). One call per distinct UTC day instead of per
 * value: when the offset at the start of a day equals the next day's, it holds for the whole day; days with a
 * transition fall back to exact calls.
 */
export class LocalTimeConverter {
  private readonly _offsets: { [dayStartMs: number]: number } = {};
  private readonly _exact: { [utcIso: string]: string } = {};
  private readonly _utcToLocal: (utcIso: string) => Promise<string>;

  /** `utcToLocal` answers like SharePoint's RegionalSettings/TimeZone/utcToLocalTime. */
  constructor(utcToLocal: (utcIso: string) => Promise<string>) {
    this._utcToLocal = utcToLocal;
  }

  /** Loads what local() needs for these UTC values. */
  public async prepare(utcIsos: string[], signal?: AbortSignal): Promise<void> {
    const days: { [ms: number]: boolean } = {};
    utcIsos.forEach((iso) => {
      const start = Math.floor(Date.parse(iso) / DAY) * DAY;
      days[start] = true;
      days[start + DAY] = true;
    });
    await this._load(
      Object.keys(days)
        .map(Number)
        .filter((d) => this._offsets[d] === undefined),
      signal
    );
    const exact = utcIsos.filter((iso, i, all) => {
      const start = Math.floor(Date.parse(iso) / DAY) * DAY;
      return all.indexOf(iso) === i && this._offsets[start] !== this._offsets[start + DAY] && this._exact[iso] === undefined;
    });
    const results = await limitConcurrency(exact.map((iso) => () => this._utcToLocal(new Date(Date.parse(iso)).toISOString())), 4, signal);
    results.forEach((r, i) => {
      if (!r.ok) throw r.error;
      this._exact[exact[i]] = toWallClock(wallClockMs(r.value));
    });
  }

  /** Local wall-clock time ("2026-03-02T09:15:00") of a prepared UTC value. */
  public local(utcIso: string): string {
    if (this._exact[utcIso] !== undefined) return this._exact[utcIso];
    const ms = Date.parse(utcIso);
    const offset = this._offsets[Math.floor(ms / DAY) * DAY];
    if (offset === undefined) throw new Error(`Local time not prepared for ${utcIso}.`);
    return toWallClock(ms + offset);
  }

  private async _load(dayStarts: number[], signal?: AbortSignal): Promise<void> {
    throwIfAborted(signal);
    const results = await limitConcurrency(dayStarts.map((d) => () => this._utcToLocal(new Date(d).toISOString())), 4, signal);
    results.forEach((r, i) => {
      if (!r.ok) throw r.error;
      this._offsets[dayStarts[i]] = wallClockMs(r.value) - dayStarts[i];
    });
  }
}
