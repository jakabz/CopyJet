import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, Spinner, Stack, Text } from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetInstallWebPartStrings';
import { missingInstallPermissions } from '../../../core/engine';
import { CopyJetError } from '../../../core/errors';
import type { ITemplateReader } from '../../../core/model';
import { openTemplate } from '../../../core/packager';
import { formatIssues, type IValidationIssue } from '../../../core/schema';
import { format, kindLabel } from './labels';
import styles from './CopyJetInstall.module.scss';

export interface ILoadStepProps {
  sp: SPFI;
  reader?: ITemplateReader;
  onLoaded: (reader: ITemplateReader | undefined) => void;
  /** Whether the user may install here (undefined while checking). */
  onPermissions: (ok: boolean) => void;
}

function loadError(e: unknown): string {
  if (e instanceof CopyJetError) {
    switch (e.code) {
      case 'TEMPLATE_ZIP_UNSUPPORTED':
        return strings.ZipLater;
      case 'TEMPLATE_PARSE':
        return strings.ParseError;
      case 'TEMPLATE_INVALID':
        return `${strings.InvalidTemplate}: ${formatIssues((e.detail as IValidationIssue[]) || []).split('\n').slice(0, 5).join('; ')}`;
      default:
        if (e.code.indexOf('SCHEMA_') === 0) return `${strings.VersionError}: ${e.message}`;
    }
  }
  return e instanceof Error ? e.message : String(e);
}

/** Step 1: pick or drop a .json template, validate it, check the user's permissions on the target. */
export const LoadStep: React.FC<ILoadStepProps> = ({ sp, reader, onLoaded, onPermissions }) => {
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [missing, setMissing] = React.useState<string[] | undefined>(undefined);
  const [dragging, setDragging] = React.useState(false);
  const input = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    let live = true;
    missingInstallPermissions(sp).then(
      (m) => {
        if (!live) return;
        setMissing(m);
        onPermissions(m.length === 0);
      },
      () => {
        if (!live) return;
        setMissing(['?']);
        onPermissions(false);
      }
    );
    return () => {
      live = false;
    };
  }, [sp]);

  const load = (file: File | undefined): void => {
    if (!file) return;
    setBusy(true);
    setError(undefined);
    onLoaded(undefined);
    openTemplate(file).then(
      (r) => {
        setBusy(false);
        onLoaded(r);
      },
      (e: unknown) => {
        setBusy(false);
        setError(loadError(e));
      }
    );
  };

  const t = reader && reader.manifest;
  const counts: Array<[string, number]> | undefined = t && ([
    [kindLabel('group'), t.groups.length],
    [kindLabel('siteField'), t.siteFields.length],
    [kindLabel('contentType'), t.contentTypes.length],
    [kindLabel('list'), t.lists.length],
    [kindLabel('listField'), t.lists.reduce((n, l) => n + (l.fields || []).length, 0)],
    [kindLabel('view'), t.lists.reduce((n, l) => n + (l.views || []).length, 0)]
  ] as Array<[string, number]>).filter(([, n]) => n > 0);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      {missing === undefined && <Spinner label={strings.PermissionsChecking} />}
      {missing && missing.length > 0 && <MessageBar messageBarType={MessageBarType.blocked}>{format(strings.PermissionsMissing, missing.join(', '))}</MessageBar>}
      <div
        className={`${styles.drop} ${dragging ? styles.dragging : ''}`}
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragging(false);
          load(e.dataTransfer.files[0]);
        }}
      >
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} wrap>
          <Text>{strings.DropHint}</Text>
          <DefaultButton text={strings.ChooseFile} onClick={() => input.current && input.current.click()} />
          <input ref={input} type="file" accept=".json,.zip,application/json" className={styles.hidden} onChange={(e) => load(e.target.files ? e.target.files[0] : undefined)} />
        </Stack>
      </div>
      {busy && <Spinner label={strings.Reading} />}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {t && (
        <table className={styles.facts}>
          <tbody>
            <tr>
              <th>{strings.TemplateLabel}</th>
              <td>
                <strong>{t.meta.name}</strong>
                {t.meta.description ? ` – ${t.meta.description}` : ''}
              </td>
            </tr>
            <tr>
              <th>{strings.SourceLabel}</th>
              <td>{t.meta.sourceSiteUrl}</td>
            </tr>
            <tr>
              <th>{strings.CreatedLabel}</th>
              <td>
                {new Date(t.meta.createdAt).toLocaleString()} · {t.meta.createdBy}
              </td>
            </tr>
            <tr>
              <th>{strings.ContentsLabel}</th>
              <td>{counts!.map(([label, n]) => `${n} ${label}`).join(', ')}</td>
            </tr>
          </tbody>
        </table>
      )}
    </Stack>
  );
};
