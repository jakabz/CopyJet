import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react';
import * as strings from 'CopyJetInstallWebPartStrings';
import type { ConflictMode, ITemplateReader } from '../../../core/model';
import { WizardShell } from '../../../shared/components/WizardShell';
import { InstallRun } from './InstallRun';
import { LoadStep } from './LoadStep';
import { PreviewStep } from './PreviewStep';
import type { ICopyJetInstallProps } from './ICopyJetInstallProps';

interface IInstallState {
  step: number;
  reader?: ITemplateReader;
  permissionsOk?: boolean;
  disabled: string[];
  mode: ConflictMode;
  /** Remounts the install run for a new installation. */
  runId: number;
}

/** Install wizard: load → preview → install → result (phase 1: structure, .json templates). */
const CopyJetInstall: React.FC<ICopyJetInstallProps> = ({ sp, siteTitle }) => {
  const [state, setState] = React.useState<IInstallState>({ step: 0, disabled: [], mode: 'skip', runId: 0 });
  const update = (patch: Partial<IInstallState>): void => setState((s) => ({ ...s, ...patch }));
  const steps = [strings.StepLoad, strings.StepPreview, strings.StepInstall, strings.StepResult];
  const template = state.reader && state.reader.manifest;

  let body: React.ReactNode;
  let footer: React.ReactNode;
  if (state.step === 0) {
    body = (
      <LoadStep
        sp={sp}
        reader={state.reader}
        onLoaded={(reader) => update({ reader, disabled: [] })}
        onPermissions={(ok) => update({ permissionsOk: ok })}
      />
    );
    footer = <PrimaryButton text={strings.Next} disabled={!template || !state.permissionsOk} onClick={() => update({ step: 1 })} />;
  } else if (state.step === 1 && template) {
    body = (
      <PreviewStep
        sp={sp}
        template={template}
        targetSiteTitle={siteTitle}
        disabled={state.disabled}
        mode={state.mode}
        onChange={(disabled, mode) => update({ disabled, mode })}
        onStart={() => update({ step: 2, runId: state.runId + 1 })}
      />
    );
    footer = <DefaultButton text={strings.Back} onClick={() => update({ step: 0 })} />;
  } else if (template) {
    body = (
      <InstallRun
        key={state.runId}
        sp={sp}
        template={template}
        disabled={state.disabled}
        mode={state.mode}
        onFinished={() => update({ step: 3 })}
        onNewInstall={() => update({ step: 0, reader: undefined, disabled: [], mode: 'skip' })}
      />
    );
  }

  return (
    <WizardShell title={strings.Title} steps={steps} current={state.step} footer={footer}>
      {body}
    </WizardShell>
  );
};

export default CopyJetInstall;
