import * as React from 'react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetSetupWebPartStrings';
import { extractTemplate } from '../../../core/engine';
import { Logger } from '../../../core/logger';
import type { ArtifactKind, IArtifactRef, ICopyJetTemplate, IDiscoveredArtifact } from '../../../core/model';
import { Button, Message, Stats, ui } from '../../../shared/components/ui';
import { Spinner } from '@fluentui/react';
import { format } from './selection';

export interface ISummaryStepProps {
  sp: SPFI;
  refs: IArtifactRef[];
  discovered: IDiscoveredArtifact[];
  name: string;
  createdBy: string;
  kindLabel: (kind: ArtifactKind) => string;
  onAddMissing: (keys: string[]) => void;
}

interface IAnalysis {
  template: ICopyJetTemplate;
  missing: IDiscoveredArtifact[];
}

/**
 * Step 3 (docs/ui setup-3): a dry run of the extraction shows what the template will contain and which
 * dependencies are missing from the selection.
 */
export const SummaryStep: React.FC<ISummaryStepProps> = ({ sp, refs, discovered, name, createdBy, kindLabel, onAddMissing }) => {
  const [analysis, setAnalysis] = React.useState<IAnalysis | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const selectionKey = refs.map((r) => r.key).join('|');

  React.useEffect(() => {
    const ac = new AbortController();
    setAnalysis(undefined);
    setError(undefined);
    extractTemplate(sp, { refs, discovered, name: name || 'CopyJet', createdBy, log: new Logger(), signal: ac.signal }).then(
      (r) => !ac.signal.aborted && setAnalysis({ template: r.writer.manifest, missing: r.missing }),
      (e: unknown) => !ac.signal.aborted && setError(e instanceof Error ? e.message : String(e))
    );
    return () => ac.abort();
    // The selection is what matters; refs/discovered are rebuilt on every render.
  }, [sp, selectionKey]);

  if (error) {
    return <Message kind="error" title={strings.CheckFailed}>{error}</Message>;
  }
  if (!analysis) {
    return <Spinner label={strings.Checking} />;
  }

  const t = analysis.template;
  const listFields = t.lists.reduce((n, l) => n + (l.fields || []).length, 0);
  const views = t.lists.reduce((n, l) => n + (l.views || []).length, 0);

  return (
    <>
      {analysis.missing.length > 0 && (
        <Message
          kind="warning"
          title={format(strings.MissingTitle, analysis.missing.length)}
          action={<Button text={strings.AddAll} onClick={() => onAddMissing(analysis.missing.map((m) => m.ref.key))} />}
        >
          {strings.MissingNote}{' '}
          {analysis.missing.map((m, i) => (
            <React.Fragment key={m.ref.key}>
              {i > 0 && ', '}
              <em>{m.title}</em> ({kindLabel(m.ref.kind).toLowerCase()})
            </React.Fragment>
          ))}
        </Message>
      )}
      <Stats
        items={[
          { label: strings.StatLists, value: t.lists.length },
          { label: strings.StatColumns, value: t.siteFields.length + listFields },
          { label: strings.StatViews, value: views },
          { label: strings.StatGroups, value: t.groups.length }
        ]}
      />
      <table className={ui.table}>
        <thead>
          <tr>
            <th>{strings.ColArea}</th>
            <th>{strings.ColContent}</th>
            <th>{strings.ColNote}</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td className={ui.strong}>{strings.AreaStructure}</td>
            <td>{format(strings.StructureText, t.siteFields.length, t.contentTypes.length, t.lists.length, listFields, views)}</td>
            <td className={ui.muted}>{strings.StructureNote}</td>
          </tr>
          {t.groups.length > 0 && (
            <tr>
              <td className={ui.strong}>{strings.AreaGroups}</td>
              <td>{t.groups.map((g) => g.title.replace('{sitename} ', '')).join(', ')}</td>
              <td className={ui.muted}>{strings.GroupsNote}</td>
            </tr>
          )}
        </tbody>
      </table>
      <Message kind="note">{strings.PersonalNote}</Message>
    </>
  );
};
