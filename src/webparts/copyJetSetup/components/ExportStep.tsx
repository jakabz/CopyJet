import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, Stack, Text } from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetSetupWebPartStrings';
import { extractTemplate, type IExtractResult } from '../../../core/engine';
import { CopyJetError } from '../../../core/errors';
import { Logger } from '../../../core/logger';
import type { ArtifactKind, IArtifactRef, IDiscoveredArtifact } from '../../../core/model';
import { formatIssues, type IValidationIssue } from '../../../core/schema';
import { LogViewer } from '../../../shared/components/LogViewer';
import { downloadBlob, templateFileName } from '../../../shared/components/download';
import { format } from './selection';

export interface IExportStepProps {
  sp: SPFI;
  refs: IArtifactRef[];
  discovered: IDiscoveredArtifact[];
  name: string;
  description: string;
  createdBy: string;
  /** Changing it starts a new extraction. */
  runId: number;
  kindLabel: (kind: ArtifactKind) => string;
  onAddMissing: (keys: string[]) => void;
  onNewTemplate: () => void;
}

type Status = 'running' | 'done' | 'failed' | 'stopped';

function message(e: unknown): string {
  if (e instanceof CopyJetError && e.code === 'TEMPLATE_INVALID') {
    return `${strings.TemplateInvalid}: ${formatIssues((e.detail as IValidationIssue[]) || []).split('\n').slice(0, 5).join('; ')}`;
  }
  return e instanceof Error ? e.message : String(e);
}

/** Step 4: extraction with progress and log, missing dependencies, download of the .json template. */
export const ExportStep: React.FC<IExportStepProps> = (props) => {
  const { sp, refs, discovered, name, description, createdBy, runId, kindLabel, onAddMissing, onNewTemplate } = props;
  const [logger, setLogger] = React.useState(() => new Logger());
  const [status, setStatus] = React.useState<Status>('running');
  const [progress, setProgress] = React.useState<{ done: number; total: number; kind?: ArtifactKind }>({ done: 0, total: 1 });
  const [result, setResult] = React.useState<IExtractResult | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const abortRef = React.useRef<AbortController | undefined>(undefined);

  React.useEffect(() => {
    const log = new Logger();
    const ac = new AbortController();
    abortRef.current = ac;
    setLogger(log);
    setStatus('running');
    setResult(undefined);
    setError(undefined);
    extractTemplate(sp, {
      refs,
      discovered,
      name,
      description,
      createdBy,
      log,
      signal: ac.signal,
      onProgress: (done, total, kind) => setProgress({ done, total, kind })
    }).then(
      (r) => {
        if (ac.signal.aborted) return;
        setResult(r);
        setStatus('done');
      },
      (e: unknown) => {
        if (ac.signal.aborted) {
          setStatus('stopped');
          return;
        }
        log.error(message(e), { code: e instanceof CopyJetError ? e.code : 'EXPORT_FAILED' });
        setError(message(e));
        setStatus('failed');
      }
    );
    return () => ac.abort();
    // Only a new runId restarts the extraction.
  }, [runId]);

  const download = (): void => {
    if (!result) return;
    result.writer.finalize().then(
      (blob) => downloadBlob(blob, templateFileName(name, new Date(), 'json')),
      (e: unknown) => setError(message(e))
    );
  };

  const stop = (): void => {
    if (abortRef.current) abortRef.current.abort();
    setStatus('stopped');
  };

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      {status === 'running' && (
        <ProgressIndicator
          label={format(strings.ExportRunning, progress.kind ? kindLabel(progress.kind) : '…')}
          percentComplete={progress.total ? progress.done / progress.total : undefined}
        />
      )}
      {status === 'done' && <MessageBar messageBarType={MessageBarType.success}>{strings.ExportDone}</MessageBar>}
      {status === 'stopped' && <MessageBar>{strings.ExportStopped}</MessageBar>}
      {error && (
        <MessageBar messageBarType={MessageBarType.error}>
          {strings.ExportFailed}: {error}
        </MessageBar>
      )}
      {result && result.missing.length > 0 && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          actions={<DefaultButton text={strings.AddMissingAndRebuild} onClick={() => onAddMissing(result.missing.map((m) => m.ref.key))} />}
        >
          <Text>{format(strings.MissingTitle, result.missing.length)}</Text>
          <ul>
            {result.missing.map((m) => (
              <li key={m.ref.key}>
                {kindLabel(m.ref.kind)}: {m.title}
              </li>
            ))}
          </ul>
        </MessageBar>
      )}
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        {status === 'running' && <DefaultButton text={strings.Stop} onClick={stop} />}
        {status === 'done' && <PrimaryButton text={strings.Download} onClick={download} />}
        {status !== 'running' && <DefaultButton text={strings.NewTemplate} onClick={onNewTemplate} />}
      </Stack>
      <LogViewer
        logger={logger}
        fileBaseName={templateFileName(name, new Date(), 'log').replace(/\.log$/, '-log')}
        labels={{
          title: strings.LogTitle,
          allLevels: strings.LogAll,
          levels: { info: strings.LogInfo, warn: strings.LogWarn, error: strings.LogError },
          empty: strings.LogEmpty,
          exportCsv: strings.LogExportCsv,
          exportJson: strings.LogExportJson
        }}
      />
    </Stack>
  );
};
