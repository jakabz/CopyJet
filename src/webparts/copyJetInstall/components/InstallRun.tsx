import * as React from 'react';
import { Icon } from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetInstallWebPartStrings';
import { createInstallContext, createProviders, runPlan, type IRunResult, type StepStatus } from '../../../core/engine';
import type { Logger } from '../../../core/logger';
import type { PrincipalMapper } from '../../../core/mapping';
import type { ArtifactKind, ConflictMode, IInstallContext, ITemplateReader } from '../../../core/model';
import { buildPlan, type IPlanStep } from '../../../core/planner';
import { LogViewer } from '../../../shared/components/LogViewer';
import { Message, ProgressBar, Stats, ui } from '../../../shared/components/ui';
import { format, kindLabel, logLabels, stepTitle } from './labels';
import styles from './CopyJetInstall.module.scss';

export type RunStatus = 'running' | 'finished';

export interface IInstallRunProps {
  sp: SPFI;
  reader: ITemplateReader;
  /** The mapper of the mapping step: people found there are not looked up again. */
  principals: PrincipalMapper;
  targetSiteTitle: string;
  disabled: string[];
  mode: ConflictMode;
  modes: { [key: string]: ConflictMode };
  logger: Logger;
  /** 'progress' = step 4 (install), 'result' = step 5. */
  view: 'progress' | 'result';
  onStatus: (status: RunStatus, stop: () => void) => void;
}

/** Install phases shown on the left (docs/ui install-4), in plan order. */
const PHASES: Array<[ArtifactKind, () => string]> = [
  ['group', () => strings.PhaseGroups],
  ['siteField', () => strings.PhaseSiteFields],
  ['contentType', () => strings.PhaseContentTypes],
  ['list', () => strings.PhaseLists],
  ['listField', () => strings.PhaseListFields],
  ['view', () => strings.PhaseViews],
  ['items', () => strings.PhaseItems],
  ['itemLookups', () => strings.PhaseItemLookups]
];

const FAILED: StepStatus[] = ['failed', 'blocked', 'cancelled'];

const duration = (ms: number): string => format(strings.Duration, Math.floor(ms / 60000), Math.floor((ms % 60000) / 1000));

/** Steps 4–5 (docs/ui install-4, install-5): the run with phases, progress and log, then the result. */
export const InstallRun: React.FC<IInstallRunProps> = ({ sp, reader, principals, targetSiteTitle, disabled, mode, modes, logger, view, onStatus }) => {
  const template = reader.manifest;
  const plan = React.useMemo(() => buildPlan(template, { disabled }), [template, disabled]);
  const [statuses, setStatuses] = React.useState<{ [key: string]: StepStatus }>({});
  const [current, setCurrent] = React.useState<IPlanStep | undefined>(undefined);
  const [result, setResult] = React.useState<IRunResult | undefined>(undefined);
  const [ctx, setCtx] = React.useState<IInstallContext | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [elapsed, setElapsed] = React.useState(0);

  React.useEffect(() => {
    const ac = new AbortController();
    const started = Date.now();
    onStatus('running', () => ac.abort());
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      e.returnValue = strings.LeaveWarning;
      return strings.LeaveWarning;
    };
    window.addEventListener('beforeunload', warn);
    const done = (): void => {
      window.removeEventListener('beforeunload', warn);
      setElapsed(Date.now() - started);
      onStatus('finished', () => undefined);
    };
    createInstallContext(sp, logger, ac.signal, { reader, principals })
      .then((c) => {
        setCtx(c);
        setCurrent(plan.steps[0]);
        return runPlan(sp, plan, c, {
          providers: createProviders(),
          mode,
          modes,
          onProgress: (e) => {
            setStatuses((s) => ({ ...s, [e.ref.key]: e.status }));
            setCurrent(plan.steps[e.done]);
          }
        });
      })
      .then(
        (r) => {
          setResult(r);
          done();
        },
        (e: unknown) => {
          const message = e instanceof Error ? e.message : String(e);
          logger.error(message, { code: 'INSTALL_FAILED' });
          setError(message);
          done();
        }
      );
    return () => {
      window.removeEventListener('beforeunload', warn);
      ac.abort();
    };
    // One run per mount: the wizard remounts this component for a new installation.
  }, []);

  const doneCount = Object.keys(statuses).length;

  if (view === 'result' && (result || error)) {
    const counts = result ? result.counts : { created: 0, updated: 0, skipped: 0, failed: 1, blocked: 0, cancelled: 0 };
    const failed = counts.failed + counts.blocked;
    const created = result ? result.steps.filter((s) => s.status === 'created') : [];
    const siteUrl = ctx ? ctx.targetSiteUrl.replace(/\/+$/, '') : '';
    const warnings = logger.entries.filter((e) => e.level !== 'info');
    const titleOf = (key: string): string => {
      const step = plan.steps.filter((s) => s.ref.key === key)[0];
      return step ? `${stepTitle(step, template, targetSiteTitle)} (${kindLabel(step.ref.kind, step.def).toLowerCase()})` : key;
    };
    return (
      <>
        {error && <Message kind="error" title={strings.InstallFailed}>{error}</Message>}
        {result && (
          <Message kind={result.aborted ? 'note' : failed ? 'warning' : 'success'} title={result.aborted ? strings.ResultStopped : failed ? strings.ResultWithErrors : strings.ResultTitle}>
            · {duration(elapsed)}
          </Message>
        )}
        <Stats
          items={[
            { label: strings.StatCreated, value: counts.created, tone: 'good' },
            { label: strings.StatUpdated, value: counts.updated },
            { label: strings.StatSkipped, value: counts.skipped },
            { label: strings.StatWarnError, value: `${logger.counts.warn} / ${failed}`, tone: logger.counts.warn || failed ? 'warn' : undefined }
          ]}
        />
        <div className={ui.split}>
          <div className={`${ui.grow} ${ui.box} ${styles.resultBox}`}>
            <span className={ui.strong}>{strings.CreatedItems}</span>
            {created.length === 0 && <span className={ui.muted}>{strings.NothingCreated}</span>}
            {created.slice(0, 8).map((s) => {
              const key = s.ref.key.replace(/^list:/, '');
              const url = s.ref.kind === 'list' ? (ctx && ctx.tokens.get('listurl', key)) || (template.lists.filter((l) => l.key === key)[0] || { url: '' }).url : undefined;
              return url ? (
                <a key={s.ref.key} className={ui.link} href={`${siteUrl}/${url}`} target="_blank" rel="noreferrer">
                  {titleOf(s.ref.key)}
                </a>
              ) : (
                <span key={s.ref.key}>{titleOf(s.ref.key)}</span>
              );
            })}
            {created.length > 8 && <span className={ui.muted}>{format(strings.AndMore, created.length - 8)}</span>}
          </div>
          <div className={`${ui.grow} ${ui.box} ${styles.resultBox}`}>
            <span className={ui.strong}>{strings.Warnings}</span>
            {warnings.length === 0 && <span className={ui.muted}>{strings.NoWarnings}</span>}
            {warnings.slice(0, 6).map((w, i) => (
              <span key={i}>
                {w.artifact ? `${titleOf(w.artifact.key)}: ` : ''}
                {w.message}
              </span>
            ))}
            {warnings.length > 6 && <span className={ui.muted}>{format(strings.AndMore, warnings.length - 6)}</span>}
          </div>
        </div>
      </>
    );
  }

  const phaseState = (kind: ArtifactKind): 'done' | 'current' | 'waiting' | 'failed' => {
    const steps = plan.steps.filter((s) => s.ref.kind === kind);
    const finished = steps.filter((s) => statuses[s.ref.key]);
    if (finished.some((s) => FAILED.indexOf(statuses[s.ref.key]) >= 0) && finished.length === steps.length) return 'failed';
    if (finished.length === steps.length) return 'done';
    if (finished.length > 0 || (current && current.ref.kind === kind)) return 'current';
    return 'waiting';
  };
  const ICONS = { done: 'CompletedSolid', current: 'Sync', failed: 'ErrorBadge', waiting: '' };

  return (
    <>
      <ProgressBar
        label={current ? `${kindLabel(current.ref.kind, current.def)}: ${stepTitle(current, template, targetSiteTitle)}` : strings.Preparing}
        detail={format(strings.InstallSteps, doneCount, plan.steps.length)}
        fraction={plan.steps.length ? doneCount / plan.steps.length : 0}
      />
      <div className={ui.split}>
        <div className={`${ui.sideWide} ${ui.box} ${styles.phases}`}>
          {PHASES.map(([kind, label], i) => {
            const n = plan.steps.filter((s) => s.ref.kind === kind).length;
            if (!n) return null;
            const state = phaseState(kind);
            return (
              <div key={kind} className={`${styles.phase} ${styles[`phase_${state}`]}`}>
                <span className={styles.phaseIcon}>{ICONS[state] ? <Icon iconName={ICONS[state]} /> : <span className={styles.waitingDot} />}</span>
                <span className={`${ui.muted} ${styles.phaseNo}`}>{i + 1}</span>
                <span>{format(label(), n)}</span>
              </div>
            );
          })}
        </div>
        <div className={ui.grow}>
          <LogViewer logger={logger} labels={logLabels} />
        </div>
      </div>
    </>
  );
};
