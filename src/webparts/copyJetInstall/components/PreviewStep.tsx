import * as React from 'react';
import { Checkbox, ChoiceGroup, MessageBar, MessageBarType, PrimaryButton, Spinner, Stack, Text } from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetInstallWebPartStrings';
import { createInstallContext, createProviders, diffPlan, type IPreviewResult } from '../../../core/engine';
import { Logger } from '../../../core/logger';
import type { ConflictMode, ICopyJetTemplate } from '../../../core/model';
import { buildPlan } from '../../../core/planner';
import { format, kindLabel, stepTitle } from './labels';
import styles from './CopyJetInstall.module.scss';

export interface IPreviewStepProps {
  sp: SPFI;
  template: ICopyJetTemplate;
  targetSiteTitle: string;
  disabled: string[];
  mode: ConflictMode;
  onChange: (disabled: string[], mode: ConflictMode) => void;
  onStart: () => void;
}

function statusText(r: IPreviewResult | undefined): string {
  if (!r) return '…';
  const changes = (r.changes || []).join(', ');
  switch (r.status) {
    case 'new':
      return strings.StatusNew;
    case 'same':
      return strings.StatusSame;
    case 'different':
      return format(strings.StatusDifferent, changes);
    case 'unsupported':
      return format(strings.StatusUnsupported, changes);
    default:
      return format(strings.StatusError, r.error instanceof Error ? r.error.message : String(r.error));
  }
}

/** Step 2: what the template would do on this site; switch items off, pick the conflict mode. */
export const PreviewStep: React.FC<IPreviewStepProps> = ({ sp, template, targetSiteTitle, disabled, mode, onChange, onStart }) => {
  const full = React.useMemo(() => buildPlan(template), [template]);
  const plan = React.useMemo(() => buildPlan(template, { disabled }), [template, disabled]);
  const [preview, setPreview] = React.useState<{ [key: string]: IPreviewResult } | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);

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

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>
        {strings.PreviewFailed}: {error}
      </MessageBar>
    );
  }
  if (!preview) {
    return <Spinner label={strings.PreviewLoading} />;
  }

  const excluded: { [key: string]: string } = {};
  plan.excluded.forEach((x) => {
    excluded[x.ref.key] =
      x.reason === 'disabled' ? strings.ExcludedDisabled : x.reason === 'cycle' ? strings.ExcludedCycle : format(strings.ExcludedDependency, x.cause || '');
  });
  const results = full.steps.map((s) => preview[s.ref.key]).filter((r) => !!r);
  const count = (st: string): number => results.filter((r) => r.status === st).length;
  const setIncluded = (key: string, on: boolean): void => onChange(on ? disabled.filter((k) => k !== key) : disabled.concat(key), mode);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Text>{format(strings.PreviewCounts, count('new'), count('same'), count('different'), count('unsupported') + count('error'))}</Text>
      <table className={styles.preview}>
        <thead>
          <tr>
            <th>{strings.ColumnInclude}</th>
            <th>{strings.ColumnKind}</th>
            <th>{strings.ColumnName}</th>
            <th>{strings.ColumnStatus}</th>
          </tr>
        </thead>
        <tbody>
          {full.steps.map((s) => {
            const r = preview[s.ref.key];
            const off = excluded[s.ref.key];
            return (
              <tr key={s.ref.key} className={off ? styles.off : r ? (styles as unknown as Record<string, string>)[r.status] : undefined}>
                <td>
                  <Checkbox
                    ariaLabel={stepTitle(s, template, targetSiteTitle)}
                    checked={disabled.indexOf(s.ref.key) < 0}
                    disabled={!!off && disabled.indexOf(s.ref.key) < 0}
                    onChange={(_, on) => setIncluded(s.ref.key, !!on)}
                  />
                </td>
                <td>{kindLabel(s.ref.kind)}</td>
                <td>{stepTitle(s, template, targetSiteTitle)}</td>
                <td>{off || statusText(r)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
      <ChoiceGroup
        label={strings.ConflictModeLabel}
        selectedKey={mode}
        options={[
          { key: 'skip', text: strings.ModeSkip },
          { key: 'update', text: strings.ModeUpdate },
          { key: 'rename', text: strings.ModeRename }
        ]}
        onChange={(_, o) => o && onChange(disabled, o.key as ConflictMode)}
      />
      <Stack horizontal horizontalAlign="end">
        <PrimaryButton text={strings.StartInstall} disabled={plan.steps.length === 0} onClick={onStart} />
      </Stack>
    </Stack>
  );
};
