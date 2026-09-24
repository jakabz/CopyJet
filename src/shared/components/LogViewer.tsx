import * as React from 'react';
import { DefaultButton, Dropdown, Stack, Text, type IDropdownOption } from '@fluentui/react';
import type { ILogEntry, Logger, LogLevel } from '../../core/logger';
import { downloadBlob } from './download';
import styles from './LogViewer.module.scss';

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
  /** File name without extension for the exports. */
  fileBaseName: string;
  /** Show at most this many latest entries (all are exported). */
  maxRows?: number;
}

const time = (iso: string): string => iso.slice(11, 19);

/** Live run log: follows the Logger, filters by level, exports CSV/JSON. */
export const LogViewer: React.FC<ILogViewerProps> = ({ logger, labels, fileBaseName, maxRows = 300 }) => {
  const [entries, setEntries] = React.useState<ReadonlyArray<ILogEntry>>(logger.entries.slice());
  const [level, setLevel] = React.useState<LogLevel | 'all'>('all');

  React.useEffect(() => {
    setEntries(logger.entries.slice());
    return logger.subscribe(() => setEntries(logger.entries.slice()));
  }, [logger]);

  const options: IDropdownOption[] = [{ key: 'all', text: labels.allLevels }].concat(
    (['info', 'warn', 'error'] as LogLevel[]).map((l) => ({ key: l, text: `${labels.levels[l]} (${logger.counts[l]})` }))
  );
  const shown = entries.filter((e) => level === 'all' || e.level === level).slice(-maxRows);

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }} wrap>
        <Text variant="mediumPlus" className={styles.title}>
          {labels.title}
        </Text>
        <Dropdown options={options} selectedKey={level} onChange={(_, o) => o && setLevel(o.key as LogLevel | 'all')} styles={{ root: { minWidth: 160 } }} />
        <DefaultButton text={labels.exportCsv} disabled={entries.length === 0} onClick={() => downloadBlob(new Blob([logger.toCsv()], { type: 'text/csv' }), `${fileBaseName}.csv`)} />
        <DefaultButton text={labels.exportJson} disabled={entries.length === 0} onClick={() => downloadBlob(new Blob([logger.toJson()], { type: 'application/json' }), `${fileBaseName}.json`)} />
      </Stack>
      <div className={styles.log} role="log" aria-live="polite">
        {shown.length === 0 && <div className={styles.empty}>{labels.empty}</div>}
        {shown.map((e, i) => (
          <div key={i} className={styles[e.level]}>
            <span className={styles.time}>{time(e.time)}</span>
            <span className={styles.level}>{labels.levels[e.level]}</span>
            <span>
              {e.artifact ? `${e.artifact.key}: ` : ''}
              {e.message}
            </span>
          </div>
        ))}
      </div>
    </Stack>
  );
};
