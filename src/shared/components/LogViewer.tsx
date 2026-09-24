import * as React from 'react';
import type { ILogEntry, Logger, LogLevel } from '../../core/logger';
import { downloadBlob } from './download';
import { Tag, type TagKind } from './ui';
import styles from './ui.module.scss';
import log from './LogViewer.module.scss';

export interface ILogViewerLabels {
  title: string;
  allLevels: string;
  levels: Record<LogLevel, string>;
  empty: string;
  exportCsv: string;
  exportJson: string;
}

export interface ILogViewerProps {
  logger: Logger;
  labels: ILogViewerLabels;
  /** Show at most this many latest entries (exports contain all). */
  maxRows?: number;
}

const TAG: Record<LogLevel, TagKind> = { info: 'info', warn: 'diff', error: 'err' };

const time = (iso: string): string => {
  const d = new Date(iso);
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
};

/** Live log panel of the mockups: header with level filter, rows with time, level tag and message. */
export const LogViewer: React.FC<ILogViewerProps> = ({ logger, labels, maxRows = 300 }) => {
  const [entries, setEntries] = React.useState<ReadonlyArray<ILogEntry>>(logger.entries.slice());
  const [level, setLevel] = React.useState<LogLevel | 'all'>('all');

  React.useEffect(() => {
    setEntries(logger.entries.slice());
    return logger.subscribe(() => setEntries(logger.entries.slice()));
  }, [logger]);

  const shown = entries.filter((e) => level === 'all' || e.level === level).slice(-maxRows);

  return (
    <div className={`${styles.box} ${log.panel}`}>
      <div className={log.head}>
        <span className={log.title}>{labels.title}</span>
        <select className={styles.select} aria-label={labels.allLevels} value={level} onChange={(e) => setLevel(e.target.value as LogLevel | 'all')}>
          <option value="all">{labels.allLevels}</option>
          {(['info', 'warn', 'error'] as LogLevel[]).map((l) => (
            <option key={l} value={l}>
              {`${labels.levels[l]} (${logger.counts[l]})`}
            </option>
          ))}
        </select>
      </div>
      <div className={log.rows} role="log" aria-live="polite">
        {shown.length === 0 && <div className={`${log.row} ${styles.muted}`}>{labels.empty}</div>}
        {shown.map((e, i) => (
          <div key={i} className={log.row}>
            <span className={`${styles.muted} ${log.time}`}>{time(e.time)}</span>
            <span className={log.level}>
              <Tag kind={TAG[e.level]}>{labels.levels[e.level]}</Tag>
            </span>
            <span>
              {e.artifact ? `${e.artifact.key}: ` : ''}
              {e.message}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
};

/** Downloads the log as CSV or JSON. */
export function downloadLog(logger: Logger, fileBaseName: string, format: 'csv' | 'json'): void {
  if (format === 'csv') {
    downloadBlob(new Blob([logger.toCsv()], { type: 'text/csv' }), `${fileBaseName}.csv`);
  } else {
    downloadBlob(new Blob([logger.toJson()], { type: 'application/json' }), `${fileBaseName}.json`);
  }
}
