import * as React from 'react';
import { Icon } from '@fluentui/react';
import * as strings from 'CopyJetSetupWebPartStrings';
import type { IDiscoveredArtifact } from '../../../core/model';
import { ui } from '../../../shared/components/ui';
import { canCopyItems, categoryOf } from './selection';
import styles from './CopyJetSetup.module.scss';

export interface IOptionsStepProps {
  /** Selected lists, libraries and groups (the rows of the options table). */
  items: IDiscoveredArtifact[];
  name: string;
  description: string;
  onName: (name: string) => void;
  onDescription: (description: string) => void;
  /** List keys switched to "structure + content". */
  content: string[];
  onContent: (key: string, on: boolean) => void;
  preserveAuthors: boolean;
  onPreserveAuthors: (on: boolean) => void;
}

function typeLabel(a: IDiscoveredArtifact): string {
  switch (categoryOf(a)) {
    case 'libraries':
      return strings.TypeLibrary;
    case 'groups':
      return strings.TypeGroup;
    default:
      return strings.TypeList;
  }
}

const Phase2: React.FC = () => <span className={ui.muted}> {strings.Phase2}</span>;

/**
 * Step 2 (docs/ui setup-2): per-item copy options with a side panel, global settings. Lists can carry their
 * items; library files, versions and group members follow later in phase 2 and are shown switched off.
 */
export const OptionsStep: React.FC<IOptionsStepProps> = ({ items, name, description, onName, onDescription, content, onContent, preserveAuthors, onPreserveAuthors }) => {
  const [open, setOpen] = React.useState<string | undefined>(undefined);
  const current = items.filter((i) => i.ref.key === open)[0];
  const withContent = (key: string): boolean => content.indexOf(key) >= 0;
  const hasContent = items.some((i) => withContent(i.ref.key));

  return (
    <div className={ui.split}>
      <div className={ui.grow} style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {items.length > 0 && (
          <table className={ui.table}>
            <thead>
              <tr>
                <th>{strings.ColName}</th>
                <th>{strings.ColType}</th>
                <th>{strings.ColCopy}</th>
                <th>{strings.ColVersionsMembers}</th>
              </tr>
            </thead>
            <tbody>
              {items.map((i) => {
                const group = i.ref.kind === 'group';
                return (
                  <tr
                    key={i.ref.key}
                    className={`${ui.clickable} ${open === i.ref.key ? ui.selected : ''}`}
                    onClick={() => setOpen(i.ref.key)}
                    tabIndex={0}
                    onKeyDown={(e) => (e.key === 'Enter' || e.key === ' ') && setOpen(i.ref.key)}
                    aria-selected={open === i.ref.key}
                  >
                    <td className={ui.strong}>{i.title}</td>
                    <td className={ui.muted}>{typeLabel(i)}</td>
                    <td>{group ? strings.CopyGroup : withContent(i.ref.key) ? strings.CopyStructureContent : strings.CopyStructure}</td>
                    <td className={ui.muted}>{group ? strings.MembersOff : categoryOf(i) === 'libraries' ? strings.Off : strings.Dash}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
        <div className={ui.box} style={{ display: 'flex', flexDirection: 'column', gap: 12, padding: 16 }}>
          <div className={ui.panelTitle}>{strings.GlobalSettings}</div>
          <div className={ui.field}>
            <label htmlFor="cj-name" className={ui.label}>
              {strings.NameLabel}
            </label>
            <input id="cj-name" className={ui.input} type="text" value={name} onChange={(e) => onName(e.target.value)} aria-invalid={!name.trim()} style={{ maxWidth: 480 }} />
            {!name.trim() && <span className={styles.error}>{strings.NameRequired}</span>}
          </div>
          <div className={ui.field}>
            <label htmlFor="cj-desc" className={ui.label}>
              {strings.DescriptionLabel}
            </label>
            <textarea id="cj-desc" className={ui.textarea} rows={2} value={description} onChange={(e) => onDescription(e.target.value)} style={{ maxWidth: 480 }} />
          </div>
          <div className={styles.inlineSettings}>
            <span>
              <input className={ui.check} type="checkbox" id="cj-authors" checked={preserveAuthors} disabled={!hasContent} onChange={(e) => onPreserveAuthors(e.target.checked)} />{' '}
              <label htmlFor="cj-authors">{strings.PreserveAuthors}</label>
            </span>
            <span>
              <label htmlFor="cj-maxsize">{strings.MaxFileSize}</label>{' '}
              <select id="cj-maxsize" className={ui.select} disabled>
                <option>250 MB</option>
              </select>
              <Phase2 />
            </span>
            <span>
              <label htmlFor="cj-format">{strings.FormatLabel}</label>{' '}
              <select id="cj-format" className={ui.select} value={hasContent ? 'zip' : 'json'} disabled aria-describedby="cj-format-note">
                <option value="json">{strings.FormatJson}</option>
                <option value="zip">{strings.FormatZip}</option>
              </select>{' '}
              <span id="cj-format-note" className={ui.muted}>
                ({strings.FormatAuto})
              </span>
            </span>
          </div>
        </div>
      </div>
      <div className={`${ui.sideWide} ${ui.box} ${styles.optionsPanel}`}>
        {!current && <span className={ui.muted}>{strings.SelectRowHint}</span>}
        {current && (
          <>
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <div className={styles.panelHeading}>{current.title}</div>
              <button type="button" className={`${ui.button} ${styles.iconButton}`} aria-label={strings.ClosePanel} onClick={() => setOpen(undefined)}>
                <Icon iconName="Cancel" />
              </button>
            </div>
            {current.ref.kind === 'group' ? (
              <>
                <span>
                  <input className={ui.check} type="checkbox" id="cj-members" disabled /> <label htmlFor="cj-members">{strings.IncludeMembers}</label>
                  <Phase2 />
                </span>
                <span className={ui.muted} style={{ fontSize: 13 }}>
                  {strings.MembersNote}
                </span>
              </>
            ) : (
              <>
                <fieldset className={styles.fieldset}>
                  <legend className={ui.label}>{strings.CopyMode}</legend>
                  <span>
                    <input type="radio" name="cj-mode" id="cj-m1" checked={!withContent(current.ref.key)} onChange={() => onContent(current.ref.key, false)} className={ui.check} />{' '}
                    <label htmlFor="cj-m1">{strings.CopyStructure}</label>
                  </span>
                  <span>
                    <input
                      type="radio"
                      name="cj-mode"
                      id="cj-m2"
                      checked={withContent(current.ref.key)}
                      disabled={!canCopyItems(current)}
                      onChange={() => onContent(current.ref.key, true)}
                      className={ui.check}
                    />{' '}
                    <label htmlFor="cj-m2">{strings.CopyStructureContent}</label>
                    {!canCopyItems(current) && <Phase2 />}
                  </span>
                </fieldset>
                <div className={ui.field}>
                  <label htmlFor="cj-folder" className={ui.label}>
                    {strings.FolderFilter}
                  </label>
                  <input id="cj-folder" className={ui.input} type="text" placeholder={strings.FolderFilterPlaceholder} disabled />
                </div>
                <span>
                  <input className={ui.check} type="checkbox" id="cj-versions" disabled /> <label htmlFor="cj-versions">{strings.CopyVersions}</label>
                  <Phase2 />
                </span>
                <span className={ui.muted} style={{ fontSize: 13 }}>
                  {strings.VersionsNote}
                </span>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
};
