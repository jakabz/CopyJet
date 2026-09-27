import * as React from 'react';
import { Icon, Spinner } from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import * as strings from 'CopyJetInstallWebPartStrings';
import { missingInstallPermissions } from '../../../core/engine';
import { CopyJetError } from '../../../core/errors';
import type { ITemplateReader } from '../../../core/model';
import { openTemplate, openTemplateFromUrl } from '../../../core/packager';
import { formatIssues, type IValidationIssue } from '../../../core/schema';
import { Button, Message, Tag, ui } from '../../../shared/components/ui';
import { format } from './labels';
import styles from './CopyJetInstall.module.scss';

export interface ILoadStepProps {
  sp: SPFI;
  targetUrl: string;
  reader?: ITemplateReader;
  fileName?: string;
  onLoaded: (reader: ITemplateReader | undefined, fileName?: string) => void;
  onPermissions: (ok: boolean) => void;
}

function loadError(e: unknown, fromUrl: boolean): string {
  if (e instanceof CopyJetError) {
    switch (e.code) {
      case 'TEMPLATE_CONTENT_MISSING':
        return strings.ContentMissing;
      case 'TEMPLATE_PARSE':
        return strings.ParseError;
      case 'TEMPLATE_INVALID':
        return `${strings.InvalidTemplate}: ${formatIssues((e.detail as IValidationIssue[]) || []).split('\n').slice(0, 5).join('; ')}`;
      default:
        if (e.code.indexOf('SCHEMA_') === 0) return `${strings.VersionError}: ${e.message}`;
    }
  }
  const text = e instanceof Error ? e.message : String(e);
  return fromUrl ? `${strings.UrlError}: ${text}` : text;
}

const host = (url: string): string => url.replace(/^https?:\/\//i, '').split('/')[0].toLowerCase();

/** Step 1 (docs/ui install-1): drop zone, library/URL panel, template card; permission check. */
export const LoadStep: React.FC<ILoadStepProps> = ({ sp, targetUrl, reader, fileName, onLoaded, onPermissions }) => {
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [missing, setMissing] = React.useState<string[] | undefined>(undefined);
  const [dragging, setDragging] = React.useState(false);
  const [url, setUrl] = React.useState('');
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

  const run = (load: () => Promise<ITemplateReader>, name: string, fromUrl: boolean): void => {
    setBusy(true);
    setError(undefined);
    onLoaded(undefined);
    load().then(
      (r) => {
        setBusy(false);
        onLoaded(r, name);
      },
      (e: unknown) => {
        setBusy(false);
        setError(loadError(e, fromUrl));
      }
    );
  };
  const loadFile = (file: File | undefined): void => {
    if (file) run(() => openTemplate(file), file.name, false);
  };
  const loadUrl = (): void => {
    if (url.trim()) run(() => openTemplateFromUrl(sp, url), decodeURIComponent(url.trim().split('/').pop() || url), true);
  };

  const t = reader && reader.manifest;
  const listFields = t ? t.lists.reduce((n, l) => n + (l.fields || []).length, 0) + t.siteFields.length : 0;
  const views = t ? t.lists.reduce((n, l) => n + (l.views || []).length, 0) : 0;
  const items = t ? t.lists.reduce((n, l) => n + (l.content && l.content.mode === 'items' ? l.content.itemCount || 0 : 0), 0) : 0;
  const warnings: string[] = t
    ? (host(t.meta.sourceSiteUrl) !== host(targetUrl) && (t.principals.length > 0 || (t.terms || []).length > 0) ? [strings.TenantDiffers] : []).concat(
        (t.meta.warnings || []).map((w) => w.message)
      )
    : [];

  return (
    <>
      {missing === undefined && <Spinner label={strings.PermissionsChecking} />}
      {missing && missing.length > 0 && (
        <Message kind="error" title={strings.PermissionsMissing}>
          {format(strings.PermissionsMissingDetail, missing.join(', '))}
        </Message>
      )}
      <div className={ui.split}>
        <div
          className={`${ui.grow} ${styles.drop} ${dragging ? styles.dragging : ''}`}
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragging(false);
            loadFile(e.dataTransfer.files[0]);
          }}
        >
          <Icon iconName="CloudUpload" className={styles.dropIcon} />
          <span className={styles.dropText}>{strings.DropHint}</span>
          <Button text={strings.Browse} onClick={() => input.current && input.current.click()} />
          <input
            ref={input}
            type="file"
            accept=".json,.zip,application/json,application/zip"
            className={styles.hidden}
            onChange={(e) => {
              loadFile(e.target.files ? e.target.files[0] : undefined);
              e.target.value = '';
            }}
          />
        </div>
        <div className={`${styles.urlPanel} ${ui.box}`}>
          <span className={ui.strong}>{strings.OrFromLibrary}</span>
          <label htmlFor="cj-url" className={ui.muted} style={{ fontSize: 13 }}>
            {strings.TemplateUrl}
          </label>
          <input id="cj-url" className={ui.input} type="text" value={url} placeholder={strings.TemplateUrlPlaceholder} onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && loadUrl()} />
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
            <Button text={strings.LoadFromUrl} disabled={!url.trim() || busy} onClick={loadUrl} />
            <Button text={strings.PickFromLibrary} disabled />
            <span className={ui.muted}>{strings.Phase2}</span>
          </div>
        </div>
      </div>
      {busy && <Spinner label={strings.Reading} />}
      {error && <Message kind="error">{error}</Message>}
      {t && (
        <div className={`${ui.box} ${styles.templateCard}`}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <span className={styles.templateName}>{fileName || t.meta.name}</span>
            <Tag kind="new">{strings.Valid}</Tag>
          </div>
          <div className={styles.factsGrid}>
            <span className={ui.muted}>{strings.FactSource}</span>
            <span className={ui.muted}>{strings.FactSchema}</span>
            <span className={ui.muted}>{strings.FactContent}</span>
            <span className={ui.muted}>{strings.FactChecksum}</span>
            <span>{t.meta.sourceSiteUrl.replace(/^https?:\/\//i, '')}</span>
            <span>{t.schemaVersion}</span>
            <span>
              {format(strings.ContentSummary, t.lists.length, listFields, views)}
              {t.meta.includesContent ? format(strings.ContentItems, items) : ''}
            </span>
            <span>{t.meta.checksum ? t.meta.checksum.slice(0, 18) + '…' : strings.NoChecksum}</span>
          </div>
          {warnings.map((w) => (
            <div key={w} style={{ display: 'flex', gap: 8, alignItems: 'center', fontSize: 13 }}>
              <Tag kind="diff">{strings.Warning}</Tag>
              <span>{w}</span>
            </div>
          ))}
        </div>
      )}
    </>
  );
};
