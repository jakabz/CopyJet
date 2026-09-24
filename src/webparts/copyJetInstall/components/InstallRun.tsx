import * as React from 'react';
import { DefaultButton, Link, MessageBar, MessageBarType, ProgressIndicator, Stack, Text } from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetInstallWebPartStrings';
import { createInstallContext, createProviders, runPlan, type IRunResult, type StepStatus } from '../../../core/engine';
import { Logger } from '../../../core/logger';
import type { ConflictMode, ICopyJetTemplate, IInstallContext } from '../../../core/model';
import { buildPlan } from '../../../core/planner';
import { LogViewer } from '../../../shared/components/LogViewer';
import { templateFileName } from '../../../shared/components/download';
import { format, logLabels } from './labels';
import styles from './CopyJetInstall.module.scss';

export interface IInstallRunProps {
  sp: SPFI;
  template: ICopyJetTemplate;
  disabled: string[];
  mode: ConflictMode;
  onFinished: () => void;
  onNewInstall: () => void;
}

const COUNT_LABELS: Array<[StepStatus, () => string]> = [
  ['created', () => strings.CountCreated],
  ['updated', () => strings.CountUpdated],
  ['skipped', () => strings.CountSkipped],
  ['failed', () => strings.CountFailed],
  ['blocked', () => strings.CountBlocked],
  ['cancelled', () => strings.CountCancelled]
];

/** Steps 3–4: runs the plan with progress, stop and log; then shows the result on the same log. */
export const InstallRun: React.FC<IInstallRunProps> = ({ sp, template, disabled, mode, onFinished, onNewInstall }) => {
  const [logger] = React.useState(() => new Logger());
  const [progress, setProgress] = React.useState<{ done: number; total: number }>({ done: 0, total: 0 });
  const [result, setResult] = React.useState<IRunResult | undefined>(undefined);
  const [ctx, setCtx] = React.useState<IInstallContext | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const abortRef = React.useRef<AbortController | undefined>(undefined);

  React.useEffect(() => {
    const ac = new AbortController();
    abortRef.current = ac;
    const plan = buildPlan(template, { disabled });
    setProgress({ done: 0, total: plan.steps.length });
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      e.returnValue = strings.LeaveWarning;
      return strings.LeaveWarning;
    };
    window.addEventListener('beforeunload', warn);
    createInstallContext(sp, logger, ac.signal)
      .then((c) => {
        setCtx(c);
        return runPlan(sp, plan, c, { providers: createProviders(), mode, onProgress: (e) => setProgress({ done: e.done, total: e.total }) });
      })
      .then(
        (r) => {
          setResult(r);
          onFinished();
        },
        (e: unknown) => {
          const message = e instanceof Error ? e.message : String(e);
          logger.error(message, { code: 'INSTALL_FAILED' });
          setError(message);
          onFinished();
        }
      )
      .then(() => window.removeEventListener('beforeunload', warn), () => window.removeEventListener('beforeunload', warn));
    return () => {
      window.removeEventListener('beforeunload', warn);
      ac.abort();
    };
    // Runs once per mount: the wizard mounts this component for a new installation.
  }, []);

  const running = !result && !error;
  const listSteps = result ? result.steps.filter((s) => s.ref.kind === 'list' && s.status !== 'failed' && s.status !== 'blocked' && s.status !== 'cancelled') : [];
  const siteUrl = ctx ? ctx.targetSiteUrl.replace(/\/+$/, '') : '';

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      {running && (
        <Stack tokens={{ childrenGap: 8 }}>
          <ProgressIndicator label={format(strings.Installing, progress.done, progress.total)} percentComplete={progress.total ? progress.done / progress.total : undefined} />
          <Stack horizontal>
            <DefaultButton text={strings.Stop} onClick={() => abortRef.current && abortRef.current.abort()} />
          </Stack>
        </Stack>
      )}
      {error && (
        <MessageBar messageBarType={MessageBarType.error}>
          {strings.InstallFailed}: {error}
        </MessageBar>
      )}
      {result && (
        <Stack tokens={{ childrenGap: 8 }}>
          <MessageBar messageBarType={result.counts.failed ? MessageBarType.warning : result.aborted ? MessageBarType.info : MessageBarType.success}>
            {result.aborted ? strings.ResultStopped : strings.ResultTitle}
          </MessageBar>
          <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
            {COUNT_LABELS.filter(([k]) => result.counts[k] > 0).map(([k, label]) => (
              <div key={k} className={styles.count}>
                <Text variant="xLarge">{result.counts[k]}</Text>
                <Text>{label()}</Text>
              </div>
            ))}
          </Stack>
          {listSteps.length > 0 && (
            <div>
              <Text>{strings.CreatedLists}</Text>
              <ul>
                {listSteps.map((s) => {
                  const key = s.ref.key.replace(/^list:/, '');
                  const def = template.lists.filter((l) => l.key === key)[0];
                  const url = (ctx && ctx.tokens.get('listurl', key)) || (def && def.url);
                  return (
                    <li key={s.ref.key}>
                      <Link href={`${siteUrl}/${url}`} target="_blank">
                        {def ? def.title : key}
                      </Link>
                    </li>
                  );
                })}
              </ul>
            </div>
          )}
          <Stack horizontal>
            <DefaultButton text={strings.NewInstall} onClick={onNewInstall} />
          </Stack>
        </Stack>
      )}
      <LogViewer logger={logger} labels={logLabels} fileBaseName={templateFileName(`${template.meta.name}-install`, new Date(), 'log').replace(/\.log$/, '')} />
    </Stack>
  );
};
