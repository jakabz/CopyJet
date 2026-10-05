import * as React from 'react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetSetupWebPartStrings';
import { extractTemplate } from '../../../core/engine';
import { lookupTargetsWithoutContent } from '../../../core/items';
import { countNodes } from '../../../core/navigation';
import { Logger } from '../../../core/logger';
import type { ArtifactKind, IArtifactRef, ICopyJetTemplate, IDiscoveredArtifact } from '../../../core/model';
import { Button, Message, Stats, ui } from '../../../shared/components/ui';
import { Spinner } from '@fluentui/react';
import { canCopyContent, format } from './selection';

export interface ISummaryStepProps {
  sp: SPFI;
  /** The structure to dry-run (pages are left out: their images would be downloaded). */
  refs: IArtifactRef[];
  /** Selected pages, shown as they are. */
  pages: IDiscoveredArtifact[];
  discovered: IDiscoveredArtifact[];
  name: string;
  createdBy: string;
  kindLabel: (kind: ArtifactKind) => string;
  /** Lists copied with items and their item count (from the discovery); items are not read in the dry run. */
  content: { lists: number; items: number; libraries: number; files: number; withVersions: number; personal: boolean };
  /** List ref keys ('list:K') copied with content. */
  contentKeys: string[];
  /** Switches content on for these list ref keys. */
  onAddContent: (keys: string[]) => void;
  onAddMissing: (keys: string[]) => void;
}

interface IAnalysis {
  template: ICopyJetTemplate;
  missing: IDiscoveredArtifact[];
}

/**
 * Step 3 (docs/ui setup-3): a dry run of the extraction shows what the template will contain and which
 * dependencies are missing from the selection. It covers the structure only: items are read at export.
 */
export const SummaryStep: React.FC<ISummaryStepProps> = ({ sp, refs, pages, discovered, name, createdBy, kindLabel, content, contentKeys, onAddContent, onAddMissing }) => {
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
  // Lookup targets of content lists that travel without items: their lookup values would stay empty.
  const withoutContent = lookupTargetsWithoutContent(
    t,
    contentKeys.map((k) => k.replace(/^list:/, ''))
  )
    .map((key) => discovered.filter((a) => a.ref.key === `list:${key}`)[0])
    .filter((a) => !!a && canCopyContent(a));
  const nav = t.navigation;
  const navParts: string[] = [];
  if (nav && nav.quickLaunch && nav.quickLaunch.length) navParts.push(format(strings.NavQuickLaunchText, countNodes(nav.quickLaunch)));
  if (nav && nav.topNavigation && nav.topNavigation.length) navParts.push(format(strings.NavTopNavigationText, countNodes(nav.topNavigation)));
  if (nav && nav.homePage) {
    const home = nav.homePage.toLowerCase();
    // Pages are not dry-run here; the selected ones are what the template will carry.
    const inTemplate = pages.some((p) => p.ref.key.toLowerCase() === `page:${home}`);
    navParts.push(format(inTemplate ? strings.NavHomePageText : strings.NavHomePageOutside, nav.homePage));
  }
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
      {withoutContent.length > 0 && (
        <Message
          kind="warning"
          title={format(strings.ContentMissingTitle, withoutContent.length)}
          action={<Button text={strings.AddContent} onClick={() => onAddContent(withoutContent.map((a) => a.ref.key))} />}
        >
          {strings.ContentMissingNote}{' '}
          {withoutContent.map((a, i) => (
            <React.Fragment key={a.ref.key}>
              {i > 0 && ', '}
              <em>{a.title}</em>
            </React.Fragment>
          ))}
        </Message>
      )}
      <Stats
        items={[
          { label: strings.StatLists, value: t.lists.length },
          ...(content.lists > 0 ? [{ label: strings.StatItems, value: content.items }] : []),
          ...(content.libraries > 0 ? [{ label: strings.StatFiles, value: content.files }] : []),
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
          {content.lists > 0 && (
            <tr>
              <td className={ui.strong}>{strings.AreaContent}</td>
              <td>{format(strings.ContentText, content.lists, content.items)}</td>
              <td className={ui.muted}>{strings.ContentNote}</td>
            </tr>
          )}
          {content.libraries > 0 && (
            <tr>
              <td className={ui.strong}>{strings.AreaFiles}</td>
              <td>
                {format(strings.FilesText, content.libraries, content.files)}
                {content.withVersions > 0 ? format(strings.FilesWithVersions, content.withVersions) : ''}
              </td>
              <td className={ui.muted}>{strings.FilesNote}</td>
            </tr>
          )}
          {pages.length > 0 && (
            <tr>
              <td className={ui.strong}>{strings.AreaPages}</td>
              <td>{pages.map((p) => p.title).join(', ')}</td>
              <td className={ui.muted}>{strings.PagesNote}</td>
            </tr>
          )}
          {navParts.length > 0 && (
            <tr>
              <td className={ui.strong}>{strings.AreaNavigation}</td>
              <td>{navParts.join(', ')}</td>
              <td className={ui.muted}>{strings.NavigationNote}</td>
            </tr>
          )}
          {t.groups.length > 0 && (
            <tr>
              <td className={ui.strong}>{strings.AreaGroups}</td>
              <td>{t.groups.map((g) => g.title.replace('{sitename} ', '')).join(', ')}</td>
              <td className={ui.muted}>{strings.GroupsNote}</td>
            </tr>
          )}
        </tbody>
      </table>
      <Message kind="note">{content.personal ? strings.PersonalNoteContent : strings.PersonalNote}</Message>
    </>
  );
};
