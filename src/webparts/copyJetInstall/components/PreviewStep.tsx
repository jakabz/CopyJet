import * as React from 'react';
import { Spinner } from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetInstallWebPartStrings';
import { createInstallContext, createProviders, diffPlan, type IPreviewResult } from '../../../core/engine';
import { Logger } from '../../../core/logger';
import type { ConflictMode, ICopyJetTemplate } from '../../../core/model';
import { buildPlan } from '../../../core/planner';
import { Message, Tag, ui, type TagKind } from '../../../shared/components/ui';
import { format, kindLabel, reasonText, stepTitle } from './labels';
import styles from './CopyJetInstall.module.scss';

export interface IPreviewStepProps {
  sp: SPFI;
  template: ICopyJetTemplate;
  targetSiteTitle: string;
  disabled: string[];
  mode: ConflictMode;
  modes: { [key: string]: ConflictMode };
  onChange: (change: { disabled?: string[]; mode?: ConflictMode; modes?: { [key: string]: ConflictMode } }) => void;
  /** Whether the plan has anything to install (enables Start). */
  onReady: (runnable: boolean) => void;
}

const TAG: Record<string, TagKind> = { new: 'new', same: 'same', different: 'diff', unsupported: 'err', error: 'err' };

function statusLabel(r: IPreviewResult): string {
  switch (r.status) {
    case 'new':
      return strings.StatusNew;
    case 'same':
      return strings.StatusSame;
    case 'different':
      return strings.StatusDifferent;
    case 'unsupported':
      return strings.StatusUnsupported;
    default:
      return strings.StatusError;
  }
}

const MODES: Array<[ConflictMode, () => string]> = [
  ['skip', () => strings.ModeSkip],
  ['update', () => strings.ModeUpdate],
  ['rename', () => strings.ModeRename]
];

/** Step 3 (docs/ui install-3): counters, global conflict mode, table with per-item mode, problem boxes. */
export const PreviewStep: React.FC<IPreviewStepProps> = ({ sp, template, targetSiteTitle, disabled, mode, modes, onChange, onReady }) => {
  const full = React.useMemo(() => buildPlan(template), [template]);
  const plan = React.useMemo(() => buildPlan(template, { disabled }), [template, disabled]);
  const [preview, setPreview] = React.useState<{ [key: string]: IPreviewResult } | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => onReady(!!preview && plan.steps.length > 0), [preview, plan]);

  React.useEffect(() => {
    const ac = new AbortController();
    createInstallContext(sp, new Logger(), ac.signal)
      .then((ctx) => diffPlan(sp, full, ctx, createProviders()))
      .then(
        (results) => {
          if (ac.signal.aborted) return;
          const byKey: { [key: string]: IPreviewResult } = {};
          results.forEach((r) => (byKey[r.ref.key] = r));
          setPreview(byKey);
        },
        (e: unknown) => !ac.signal.aborted && setError(e instanceof Error ? e.message : String(e))
      );
    return () => ac.abort();
  }, [sp, full]);

  if (error) return <Message kind="error" title={strings.PreviewFailed}>{error}</Message>;
  if (!preview) return <Spinner label={strings.PreviewLoading} />;

  const excluded: { [key: string]: string } = {};
  plan.excluded.forEach((x) => {
    excluded[x.ref.key] = x.reason === 'disabled' ? strings.ExcludedDisabled : x.reason === 'cycle' ? strings.ExcludedCycle : format(strings.ExcludedDependency, x.cause || '');
  });
  const results = full.steps.map((s) => preview[s.ref.key]).filter((r) => !!r);
  const count = (st: string): number => results.filter((r) => r.status === st).length;
  const problems = full.steps.filter((s) => preview[s.ref.key] && (preview[s.ref.key].status === 'unsupported' || preview[s.ref.key].status === 'error'));
  const setIncluded = (key: string, on: boolean): void => onChange({ disabled: on ? disabled.filter((k) => k !== key) : disabled.concat(key) });

  return (
    <>
      <div className={styles.previewBar}>
        <Tag kind="new">{format(strings.CountNew, count('new'))}</Tag>
        <Tag kind="same">{format(strings.CountSame, count('same'))}</Tag>
        <Tag kind="diff">{format(strings.CountDifferent, count('different'))}</Tag>
        <Tag kind="err">{format(strings.CountProblem, count('unsupported') + count('error'))}</Tag>
        <div className={ui.spacer} />
        <label htmlFor="cj-global-mode">{strings.GlobalMode}</label>
        <select id="cj-global-mode" className={ui.select} value={mode} onChange={(e) => onChange({ mode: e.target.value as ConflictMode, modes: {} })}>
          {MODES.map(([m, label]) => (
            <option key={m} value={m}>
              {label()}
            </option>
          ))}
        </select>
      </div>
      <div className={`${ui.box} ${styles.previewTable}`}>
        <table className={ui.table}>
          <thead>
            <tr>
              <th style={{ width: 32 }}>
                <span className={styles.srOnly}>{strings.ColSelect}</span>
              </th>
              <th>{strings.ColKind}</th>
              <th>{strings.ColName}</th>
              <th>{strings.ColStatus}</th>
              <th>{strings.ColConflict}</th>
            </tr>
          </thead>
          <tbody>
            {full.steps.map((s) => {
              const r = preview[s.ref.key];
              const off = excluded[s.ref.key];
              const title = stepTitle(s, template, targetSiteTitle);
              const exists = r && (r.status === 'same' || r.status === 'different');
              return (
                <tr key={s.ref.key} className={off ? styles.off : undefined}>
                  <td>
                    <input
                      className={ui.check}
                      type="checkbox"
                      aria-label={title}
                      checked={disabled.indexOf(s.ref.key) < 0}
                      disabled={!!off && disabled.indexOf(s.ref.key) < 0}
                      onChange={(e) => setIncluded(s.ref.key, e.target.checked)}
                    />
                  </td>
                  <td className={ui.muted}>{kindLabel(s.ref.kind, s.def)}</td>
                  <td className={ui.strong}>{title}</td>
                  <td>
                    {off ? (
                      <span className={ui.muted}>
                        <Tag kind="same">{strings.Excluded}</Tag> {off}
                      </span>
                    ) : r ? (
                      <span title={(r.changes || []).join(', ')}>
                        <Tag kind={TAG[r.status]}>{statusLabel(r)}</Tag>
                        {r.status === 'different' && r.changes && <span className={`${ui.muted} ${styles.changes}`}> {r.changes.map(reasonText).join(', ')}</span>}
                      </span>
                    ) : null}
                  </td>
                  <td>
                    {exists && !off ? (
                      <select
                        className={ui.select}
                        aria-label={`${strings.ColConflict}: ${title}`}
                        value={modes[s.ref.key] || mode}
                        onChange={(e) => onChange({ modes: { ...modes, [s.ref.key]: e.target.value as ConflictMode } })}
                      >
                        {MODES.map(([m, label]) => (
                          <option key={m} value={m}>
                            {label()}
                          </option>
                        ))}
                      </select>
                    ) : (
                      <span className={ui.muted}>—</span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {problems.map((s) => {
        const r = preview[s.ref.key];
        const why = r.status === 'error' ? (r.error instanceof Error ? r.error.message : String(r.error)) : (r.changes || []).map(reasonText).join(', ');
        return (
          <Message key={s.ref.key} kind="error">
            <strong>{stepTitle(s, template, targetSiteTitle)}</strong> – {why}
          </Message>
        );
      })}
    </>
  );
};
