import * as React from 'react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetInstallWebPartStrings';
import type { ITemplateReader } from '../../../core/model';
import { finishedSteps, SpRunStore, templateChecksum, type IRunState } from '../../../core/state';
import { Button, Message } from '../../../shared/components/ui';
import { format } from './labels';

export interface IResumeBannerProps {
  sp: SPFI;
  reader: ITemplateReader;
  /** The chosen earlier run, or undefined for a new installation. */
  resume?: IRunState;
  onChange: (resume: IRunState | undefined) => void;
}

const when = (iso: string): string => new Date(iso).toLocaleString();

/** Offers to resume the latest unfinished run of the loaded template on this site (CopyJetLog). */
export const ResumeBanner: React.FC<IResumeBannerProps> = ({ sp, reader, resume, onChange }) => {
  const [found, setFound] = React.useState<IRunState | undefined>(undefined);
  const [dismissed, setDismissed] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    let live = true;
    setFound(undefined);
    setDismissed(false);
    setError(undefined);
    templateChecksum(reader.storedManifest)
      .then((checksum) => new SpRunStore(sp).findUnfinished(checksum))
      .then(
        (run) => live && setFound(run),
        (e: unknown) => live && setError(e instanceof Error ? e.message : String(e))
      );
    return () => {
      live = false;
    };
  }, [reader]);

  if (error) return <Message kind="warning">{format(strings.ResumeCheckFailed, error)}</Message>;
  if (!found || dismissed) return null;
  if (resume) {
    return (
      <Message kind="note" title={strings.ResumeChosen} action={<Button text={strings.ResumeNew} onClick={() => onChange(undefined)} />}>
        {format(strings.ResumeText, when(resume.started), when(resume.updated), Object.keys(finishedSteps(resume)).length, Object.keys(resume.steps).length - Object.keys(finishedSteps(resume)).length)}
      </Message>
    );
  }
  const finished = Object.keys(finishedSteps(found)).length;
  return (
    <Message
      kind="warning"
      title={strings.ResumeTitle}
      action={
        <>
          <Button text={strings.ResumeContinue} kind="primary" onClick={() => onChange(found)} />
          <Button
            text={strings.ResumeNew}
            onClick={() => {
              setDismissed(true);
              onChange(undefined);
            }}
          />
        </>
      }
    >
      {format(strings.ResumeText, when(found.started), when(found.updated), finished, Object.keys(found.steps).length - finished)}
    </Message>
  );
};
