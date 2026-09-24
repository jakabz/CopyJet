import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton, Spinner, Stack, Text, TextField } from '@fluentui/react';
import * as strings from 'CopyJetSetupWebPartStrings';
import { discoverSite, type IDiscovery } from '../../../core/engine';
import type { ArtifactKind } from '../../../core/model';
import { WizardShell } from '../../../shared/components/WizardShell';
import { ExportStep } from './ExportStep';
import { SelectStep } from './SelectStep';
import { countByKind, format, selectedRefs } from './selection';
import styles from './CopyJetSetup.module.scss';
import type { ICopyJetSetupProps } from './ICopyJetSetupProps';

const KIND_ORDER: ArtifactKind[] = ['group', 'siteField', 'contentType', 'list', 'listField', 'view'];

export function kindLabel(kind: ArtifactKind): string {
  const labels: Partial<Record<ArtifactKind, string>> = {
    group: strings.KindGroup,
    siteField: strings.KindSiteField,
    contentType: strings.KindContentType,
    list: strings.KindList,
    listField: strings.KindListField,
    view: strings.KindView
  };
  return labels[kind] || kind;
}

interface ISetupState {
  step: number;
  discovery?: IDiscovery;
  loadError?: string;
  selected: string[];
  name: string;
  description: string;
  /** Incremented to (re)start the extraction in the export step. */
  runId: number;
}

type Action =
  | { type: 'loaded'; discovery: IDiscovery; siteTitle: string }
  | { type: 'loadFailed'; message: string }
  | { type: 'select'; selected: string[] }
  | { type: 'step'; step: number }
  | { type: 'name'; name: string }
  | { type: 'description'; description: string }
  | { type: 'export' }
  | { type: 'addMissing'; keys: string[] }
  | { type: 'reset' };

function reducer(state: ISetupState, action: Action): ISetupState {
  switch (action.type) {
    case 'loaded':
      return { ...state, discovery: action.discovery, name: state.name || format(strings.DefaultTemplateName, action.siteTitle) };
    case 'loadFailed':
      return { ...state, loadError: action.message };
    case 'select':
      return { ...state, selected: action.selected };
    case 'step':
      return { ...state, step: action.step };
    case 'name':
      return { ...state, name: action.name };
    case 'description':
      return { ...state, description: action.description };
    case 'export':
      return { ...state, step: 3, runId: state.runId + 1 };
    case 'addMissing':
      return { ...state, selected: state.selected.concat(action.keys.filter((k) => state.selected.indexOf(k) < 0)), runId: state.runId + 1 };
    case 'reset':
      return { ...state, step: 0, selected: [] };
    default:
      return state;
  }
}

/** Setup wizard: select → options → summary → export (phase 1: structure only, .json template). */
const CopyJetSetup: React.FC<ICopyJetSetupProps> = ({ sp, siteTitle, createdBy }) => {
  const [state, dispatch] = React.useReducer(reducer, { step: 0, selected: [], name: '', description: '', runId: 0 });

  React.useEffect(() => {
    const ac = new AbortController();
    discoverSite(sp, ac.signal).then(
      (discovery) => !ac.signal.aborted && dispatch({ type: 'loaded', discovery, siteTitle }),
      (e: unknown) => !ac.signal.aborted && dispatch({ type: 'loadFailed', message: e instanceof Error ? e.message : String(e) })
    );
    return () => ac.abort();
  }, [sp, siteTitle]);

  const artifacts = state.discovery ? state.discovery.artifacts : [];
  const refs = selectedRefs(artifacts, state.selected);
  const counts = countByKind(refs);
  const steps = [strings.StepSelect, strings.StepOptions, strings.StepSummary, strings.StepExport];
  const go = (step: number): void => dispatch({ type: 'step', step });

  const summary = (
    <ul className={styles.summary}>
      {KIND_ORDER.filter((k) => counts[k]).map((k) => (
        <li key={k}>
          {kindLabel(k)}: <strong>{counts[k]}</strong>
        </li>
      ))}
    </ul>
  );

  let body: React.ReactNode;
  let footer: React.ReactNode;
  if (state.loadError) {
    body = (
      <MessageBar messageBarType={MessageBarType.error}>
        {strings.LoadError}: {state.loadError}
      </MessageBar>
    );
  } else if (!state.discovery) {
    body = <Spinner label={strings.Loading} />;
  } else if (state.step === 0) {
    body = (
      <Stack horizontal tokens={{ childrenGap: 24 }} wrap>
        <Stack.Item grow className={styles.tree}>
          {state.discovery.errors.length > 0 && (
            <MessageBar messageBarType={MessageBarType.warning}>{format(strings.DiscoverPartialError, state.discovery.errors.map((e) => kindLabel(e.kind)).join(', '))}</MessageBar>
          )}
          <SelectStep artifacts={artifacts} selected={state.selected} onChange={(selected) => dispatch({ type: 'select', selected })} />
        </Stack.Item>
        <Stack.Item className={styles.side}>
          <Text variant="mediumPlus">{strings.SelectionSummary}</Text>
          {refs.length ? summary : <Text className={styles.meta}>{strings.NothingSelected}</Text>}
        </Stack.Item>
      </Stack>
    );
    footer = <PrimaryButton text={strings.Next} disabled={refs.length === 0} onClick={() => go(1)} />;
  } else if (state.step === 1) {
    body = (
      <Stack tokens={{ childrenGap: 12 }} className={styles.form}>
        <TextField
          label={strings.NameLabel}
          required
          value={state.name}
          onChange={(_, v) => dispatch({ type: 'name', name: v || '' })}
          errorMessage={state.name.trim() ? undefined : strings.NameRequired}
        />
        <TextField label={strings.DescriptionLabel} multiline rows={3} value={state.description} onChange={(_, v) => dispatch({ type: 'description', description: v || '' })} />
        <MessageBar>{strings.ContentLaterInfo}</MessageBar>
      </Stack>
    );
    footer = (
      <>
        <DefaultButton text={strings.Back} onClick={() => go(0)} />
        <PrimaryButton text={strings.Next} disabled={!state.name.trim()} onClick={() => go(2)} />
      </>
    );
  } else if (state.step === 2) {
    body = (
      <Stack tokens={{ childrenGap: 8 }}>
        <Text variant="mediumPlus">{state.name}</Text>
        <Text>{strings.SummaryIntro}</Text>
        {summary}
      </Stack>
    );
    footer = (
      <>
        <DefaultButton text={strings.Back} onClick={() => go(1)} />
        <PrimaryButton text={strings.CreateTemplate} onClick={() => dispatch({ type: 'export' })} />
      </>
    );
  } else {
    body = (
      <ExportStep
        sp={sp}
        refs={refs}
        discovered={artifacts}
        name={state.name}
        description={state.description}
        createdBy={createdBy}
        runId={state.runId}
        kindLabel={kindLabel}
        onAddMissing={(keys) => dispatch({ type: 'addMissing', keys })}
        onNewTemplate={() => dispatch({ type: 'reset' })}
      />
    );
  }

  return (
    <WizardShell title={strings.Title} steps={steps} current={state.step} footer={footer}>
      {body}
    </WizardShell>
  );
};

export default CopyJetSetup;
