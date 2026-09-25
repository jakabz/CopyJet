import * as React from 'react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetSetupWebPartStrings';
import { extractTemplate, type IExtractResult } from '../../../core/engine';
import { CopyJetError } from '../../../core/errors';
import { Logger } from '../../../core/logger';
import type { ArtifactKind, IArtifactRef, IDiscoveredArtifact } from '../../../core/model';
import { formatIssues, type IValidationIssue } from '../../../core/schema';
import { LogViewer } from '../../../shared/components/LogViewer';
import { downloadBlob, templateFileName } from '../../../shared/components/download';
import { Button, Message, ProgressBar, ui } from '../../../shared/components/ui';
import { format } from './selection';
import styles from './CopyJetSetup.module.scss';

export type ExportStatus = 'running' | 'done' | 'failed' | 'stopped';

export interface IExportStepProps {
  sp: SPFI;
  refs: IArtifactRef[];
  discovered: IDiscoveredArtifact[];
  name: string;
  description: string;
  createdBy: string;
  preserveAuthors: boolean;
  kindLabel: (kind: ArtifactKind) => string;
  logLabels: React.ComponentProps<typeof LogViewer>['labels'];
  /** Reports the status so the wizard can show Stop / New template in its footer. */
  onStatus: (status: ExportStatus, stop: () => void) => void;
}

function message(e: unknown): string {
  if (e instanceof CopyJetError && e.code === 'TEMPLATE_INVALID') {
    return `${strings.TemplateInvalid}: ${formatIssues((e.detail as IValidationIssue[]) || []).split('\n').slice(0, 5).join('; ')}`;
  }
  return e instanceof Error ? e.message : String(e);
}

/** Step 4 (docs/ui setup-4): progress, log, output card with download. */
export const ExportStep: React.FC<IExportStepProps> = ({ sp, refs, discovered, name, description, createdBy, preserveAuthors, kindLabel, logLabels, onStatus }) => {
  const [logger] = React.useState(() => new Logger());
  const [status, setStatus] = React.useState<ExportStatus>('running');
  const [progress, setProgress] = React.useState<{ done: number; total: number; kind?: ArtifactKind }>({ done: 0, total: 1 });
  const [result, setResult] = React.useState<IExtractResult | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  // Content (list items) makes a .zip package; structure alone stays a readable .json.
  const extension = refs.some((r) => r.kind === 'items') ? 'zip' : 'json';
  const fileName = React.useMemo(() => templateFileName(name, new Date(), extension), [name, extension]);

  React.useEffect(() => {
    const ac = new AbortController();
    const report = (s: ExportStatus): void => {
      setStatus(s);
      onStatus(s, () => ac.abort());
    };
    report('running');
    extractTemplate(sp, { refs, discovered, name, description, createdBy, preserveAuthors, log: logger, signal: ac.signal, onProgress: (done, total, kind) => setProgress({ done, total, kind }) }).then(
      (r) => {
        if (ac.signal.aborted) return;
        setResult(r);
        report('done');
      },
      (e: unknown) => {
        if (ac.signal.aborted) {
          report('stopped');
          return;
        }
        logger.error(message(e), { code: e instanceof CopyJetError ? e.code : 'EXPORT_FAILED' });
        setError(message(e));
        report('failed');
      }
    );
    return () => ac.abort();
    // One extraction per mount: the wizard remounts this step for a new export.
  }, []);

  const download = (): void => {
    if (!result) return;
    result.writer.finalize().then(
      (blob) => downloadBlob(blob, fileName),
      (e: unknown) => setError(message(e))
    );
  };

  return (
    <>
      {status === 'running' ? (
        <ProgressBar label={format(strings.ExportRunning, progress.kind ? kindLabel(progress.kind) : '…')} detail={format(strings.ExportSteps, progress.done, progress.total)} fraction={progress.total ? progress.done / progress.total : 0} />
      ) : (
        <ProgressBar label={strings.ExportDoneLabel} detail={format(strings.ExportSteps, progress.total, progress.total)} fraction={status === 'done' ? 1 : progress.total ? progress.done / progress.total : 0} />
      )}
      {status === 'done' && <Message kind="success">{strings.ExportDone}</Message>}
      {status === 'stopped' && <Message kind="note">{strings.ExportStopped}</Message>}
      {error && <Message kind="error" title={strings.ExportFailed}>{error}</Message>}
      <LogViewer logger={logger} labels={logLabels} />
      <div className={`${ui.box} ${styles.output}`}>
        <div className={styles.outputText}>
          <span className={ui.strong}>{strings.Output}</span>
          <span className={ui.muted} style={{ fontSize: 13 }}>
            {status === 'done' ? fileName : format(strings.OutputPending, fileName)}
          </span>
        </div>
        <span>
          <input className={ui.check} type="checkbox" id="cj-save-lib" disabled /> <label htmlFor="cj-save-lib">{strings.SaveToLibrary}</label>
          <span className={ui.muted}> {strings.Phase2}</span>
        </span>
        <Button text={strings.Download} kind={status === 'done' ? 'primary' : 'default'} disabled={status !== 'done'} onClick={download} />
      </div>
    </>
  );
};
