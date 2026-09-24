import * as React from 'react';
import * as strings from 'CopyJetInstallWebPartStrings';
import { Logger } from '../../../core/logger';
import type { ConflictMode, ITemplateReader } from '../../../core/model';
import { downloadLog } from '../../../shared/components/LogViewer';
import { WizardShell } from '../../../shared/components/WizardShell';
import { templateFileName } from '../../../shared/components/download';
import { Button } from '../../../shared/components/ui';
import { InstallRun, type RunStatus } from './InstallRun';
import { LoadStep } from './LoadStep';
import { MappingStep } from './MappingStep';
import { PreviewStep } from './PreviewStep';
import type { ICopyJetInstallProps } from './ICopyJetInstallProps';

interface IInstallState {
  step: number;
  reader?: ITemplateReader;
  fileName?: string;
  permissionsOk?: boolean;
  disabled: string[];
  mode: ConflictMode;
  modes: { [key: string]: ConflictMode };
  previewReady: boolean;
  /** Remounts the install run for a new installation. */
  runId: number;
  runStatus?: RunStatus;
  logger?: Logger;
}

const INITIAL: IInstallState = { step: 0, disabled: [], mode: 'skip', modes: {}, previewReady: false, runId: 0 };

/** Install wizard as in docs/ui (install-1 … install-5). */
const CopyJetInstall: React.FC<ICopyJetInstallProps> = ({ sp, siteTitle, siteUrl }) => {
  const [state, setState] = React.useState<IInstallState>(INITIAL);
  const stopRef = React.useRef<() => void>(() => undefined);
  const update = (patch: Partial<IInstallState>): void => setState((s) => ({ ...s, ...patch }));
  const steps = [strings.StepLoad, strings.StepMapping, strings.StepPreview, strings.StepInstall, strings.StepResult];
  const subtitles = [strings.SubtitleLoad, strings.SubtitleMapping, strings.SubtitlePreview, strings.SubtitleInstall, strings.SubtitleResult];
  const template = state.reader && state.reader.manifest;
  const logName = template ? templateFileName(`${template.meta.name}-install`, new Date(), 'log').replace(/\.log$/, '') : 'copyjet-install';

  let body: React.ReactNode;
  let footerStart: React.ReactNode;
  let footerEnd: React.ReactNode;
  if (state.step === 0 || !template) {
    body = (
      <LoadStep
        sp={sp}
        targetUrl={siteUrl}
        reader={state.reader}
        fileName={state.fileName}
        onLoaded={(reader, fileName) => update({ reader, fileName, disabled: [], modes: {} })}
        onPermissions={(ok) => update({ permissionsOk: ok })}
      />
    );
    footerEnd = <Button text={strings.Next} kind="primary" disabled={!template || !state.permissionsOk} onClick={() => update({ step: 1 })} />;
  } else if (state.step === 1) {
    body = <MappingStep template={template} />;
    footerEnd = (
      <>
        <Button text={strings.Back} onClick={() => update({ step: 0 })} />
        <Button text={strings.Next} kind="primary" onClick={() => update({ step: 2, previewReady: false })} />
      </>
    );
  } else if (state.step === 2) {
    body = (
      <PreviewStep
        sp={sp}
        template={template}
        targetSiteTitle={siteTitle}
        disabled={state.disabled}
        mode={state.mode}
        modes={state.modes}
        onChange={(c) => update(c)}
        onReady={(previewReady) => update({ previewReady })}
      />
    );
    footerEnd = (
      <>
        <Button text={strings.Back} onClick={() => update({ step: 1 })} />
        <Button text={strings.StartInstall} kind="primary" disabled={!state.previewReady} onClick={() => update({ step: 3, runId: state.runId + 1, runStatus: 'running', logger: new Logger() })} />
      </>
    );
  } else {
    body = (
      <InstallRun
        key={state.runId}
        sp={sp}
        template={template}
        targetSiteTitle={siteTitle}
        disabled={state.disabled}
        mode={state.mode}
        modes={state.modes}
        logger={state.logger!}
        view={state.step === 4 ? 'result' : 'progress'}
        onStatus={(runStatus, stop) => {
          stopRef.current = stop;
          update({ runStatus });
        }}
      />
    );
    if (state.step === 3) {
      footerStart = <Button text={strings.Stop} kind="danger" disabled={state.runStatus !== 'running'} onClick={() => stopRef.current()} />;
      footerEnd = <Button text={strings.Next} kind="primary" disabled={state.runStatus !== 'finished'} onClick={() => update({ step: 4 })} />;
    } else {
      footerEnd = (
        <>
          <Button text={strings.DownloadCsv} onClick={() => state.logger && downloadLog(state.logger, logName, 'csv')} />
          <Button text={strings.DownloadJson} onClick={() => state.logger && downloadLog(state.logger, logName, 'json')} />
          <Button text={strings.NewInstall} kind="primary" onClick={() => setState({ ...INITIAL, permissionsOk: state.permissionsOk, runId: state.runId })} />
        </>
      );
    }
  }

  return (
    <WizardShell title={strings.Title} subtitle={subtitles[state.step]} steps={steps} current={state.step} footerStart={footerStart} footerEnd={footerEnd}>
      {body}
    </WizardShell>
  );
};

export default CopyJetInstall;
