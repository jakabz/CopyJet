import * as React from 'react';
import { Spinner } from '@fluentui/react';
import * as strings from 'CopyJetSetupWebPartStrings';
import { discoverSite, type IDiscovery } from '../../../core/engine';
import type { ArtifactKind } from '../../../core/model';
import { WizardShell } from '../../../shared/components/WizardShell';
import { Button, Message } from '../../../shared/components/ui';
import { ExportStep, type ExportStatus } from './ExportStep';
import { OptionsStep } from './OptionsStep';
import { SelectStep } from './SelectStep';
import { SummaryStep } from './SummaryStep';
import { format, selectedArtifacts, selectedRefs } from './selection';
import type { ICopyJetSetupProps } from './ICopyJetSetupProps';

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

const LOG_LABELS = {
  title: strings.LogTitle,
  allLevels: strings.LogAll,
  levels: { info: strings.LogInfo, warn: strings.LogWarn, error: strings.LogError },
  empty: strings.LogEmpty,
  exportCsv: strings.LogExportCsv,
  exportJson: strings.LogExportJson
};

interface ISetupState {
  step: number;
  discovery?: IDiscovery;
  loadError?: string;
  selected: string[];
  name: string;
  description: string;
  /** Remounts the export step for a new extraction. */
  runId: number;
  exportStatus?: ExportStatus;
}

type Action =
  | { type: 'loaded'; discovery: IDiscovery; siteTitle: string }
  | { type: 'loadFailed'; message: string }
  | { type: 'select'; selected: string[] }
  | { type: 'step'; step: number }
  | { type: 'name'; name: string }
  | { type: 'description'; description: string }
  | { type: 'export' }
  | { type: 'exportStatus'; status: ExportStatus }
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
      return { ...state, step: 3, runId: state.runId + 1, exportStatus: 'running' };
    case 'exportStatus':
      return { ...state, exportStatus: action.status };
    case 'addMissing':
      return { ...state, selected: state.selected.concat(action.keys.filter((k) => state.selected.indexOf(k) < 0)) };
    case 'reset':
      return { ...state, step: 0, selected: [], exportStatus: undefined };
    default:
      return state;
  }
}

/** Setup wizard as in docs/ui (setup-1 … setup-4); phase 1 copies structure into a .json template. */
const CopyJetSetup: React.FC<ICopyJetSetupProps> = ({ sp, siteTitle, createdBy }) => {
  const [state, dispatch] = React.useReducer(reducer, { step: 0, selected: [], name: '', description: '', runId: 0 });
  const stopRef = React.useRef<(() => void) | undefined>(undefined);

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
  const steps = [strings.StepSelect, strings.StepOptions, strings.StepSummary, strings.StepExport];
  const subtitles = [strings.SubtitleSelect, strings.SubtitleOptions, strings.SubtitleSummary, strings.SubtitleExport];
  const go = (step: number): void => dispatch({ type: 'step', step });

  let body: React.ReactNode;
  let footerStart: React.ReactNode;
  let footerEnd: React.ReactNode;
  if (state.loadError) {
    body = <Message kind="error" title={strings.LoadError}>{state.loadError}</Message>;
  } else if (!state.discovery) {
    body = <Spinner label={strings.Loading} />;
  } else if (state.step === 0) {
    body = (
      <>
        {state.discovery.errors.length > 0 && (
          <Message kind="warning">{format(strings.DiscoverPartialError, state.discovery.errors.map((e) => kindLabel(e.kind)).join(', '))}</Message>
        )}
        <SelectStep artifacts={artifacts} selected={state.selected} onChange={(selected) => dispatch({ type: 'select', selected })} />
      </>
    );
    footerEnd = <Button text={strings.Next} kind="primary" disabled={refs.length === 0} onClick={() => go(1)} />;
  } else if (state.step === 1) {
    body = (
      <OptionsStep
        items={selectedArtifacts(artifacts, state.selected).filter((a) => a.ref.kind === 'list' || a.ref.kind === 'group')}
        name={state.name}
        description={state.description}
        onName={(name) => dispatch({ type: 'name', name })}
        onDescription={(description) => dispatch({ type: 'description', description })}
      />
    );
    footerEnd = (
      <>
        <Button text={strings.Back} onClick={() => go(0)} />
        <Button text={strings.Next} kind="primary" disabled={!state.name.trim()} onClick={() => go(2)} />
      </>
    );
  } else if (state.step === 2) {
    body = (
      <SummaryStep
        sp={sp}
        refs={refs}
        discovered={artifacts}
        name={state.name}
        createdBy={createdBy}
        kindLabel={kindLabel}
        onAddMissing={(keys) => dispatch({ type: 'addMissing', keys })}
      />
    );
    footerEnd = (
      <>
        <Button text={strings.Back} onClick={() => go(1)} />
        <Button text={strings.CreateTemplate} kind="primary" onClick={() => dispatch({ type: 'export' })} />
      </>
    );
  } else {
    body = (
      <ExportStep
        key={state.runId}
        sp={sp}
        refs={refs}
        discovered={artifacts}
        name={state.name}
        description={state.description}
        createdBy={createdBy}
        kindLabel={kindLabel}
        logLabels={LOG_LABELS}
        onStatus={(status, stop) => {
          stopRef.current = stop;
          dispatch({ type: 'exportStatus', status });
        }}
      />
    );
    const running = state.exportStatus === 'running';
    footerStart = running ? <Button text={strings.Stop} kind="danger" onClick={() => stopRef.current && stopRef.current()} /> : undefined;
    footerEnd = running ? undefined : <Button text={strings.NewTemplate} kind="primary" onClick={() => dispatch({ type: 'reset' })} />;
  }

  return (
    <WizardShell title={strings.Title} subtitle={subtitles[state.step]} steps={steps} current={state.step} footerStart={footerStart} footerEnd={footerEnd}>
      {body}
    </WizardShell>
  );
};

export default CopyJetSetup;
