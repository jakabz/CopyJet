import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './ui.module.scss';

/** Building blocks in the style of the docs/ui mockups, shared by the Setup and Install web parts. */

export type TagKind = 'new' | 'same' | 'diff' | 'err' | 'info';

export const Tag: React.FC<{ kind: TagKind; children?: React.ReactNode }> = ({ kind, children }) => (
  <span className={`${styles.tag} ${styles[kind]}`}>{children}</span>
);

export interface IStat {
  label: string;
  value: React.ReactNode;
  tone?: 'good' | 'warn';
}

export const Stats: React.FC<{ items: IStat[] }> = ({ items }) => (
  <div className={styles.stats}>
    {items.map((s) => (
      <div key={s.label} className={styles.stat}>
        <div className={`${styles.muted}`} style={{ fontSize: 13 }}>
          {s.label}
        </div>
        <div className={`${styles.statValue} ${s.tone === 'good' ? styles.statGood : s.tone === 'warn' ? styles.statWarn : ''}`}>{s.value}</div>
      </div>
    ))}
  </div>
);

export const ProgressBar: React.FC<{ label: string; detail?: string; fraction: number }> = ({ label, detail, fraction }) => (
  <div className={styles.progress} role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(fraction * 100)} aria-label={label}>
    <div className={styles.progressText}>
      <span>{label}</span>
      {detail && <span className={styles.muted}>{detail}</span>}
    </div>
    <div className={styles.bar}>
      <div style={{ width: `${Math.max(0, Math.min(1, fraction)) * 100}%` }} />
    </div>
  </div>
);

export type MessageKind = 'warning' | 'error' | 'success' | 'note';

const MESSAGE_ICON: Record<MessageKind, string> = { warning: 'Warning', error: 'ErrorBadge', success: 'Completed', note: 'Info' };

export const Message: React.FC<{ kind: MessageKind; title?: React.ReactNode; action?: React.ReactNode; children?: React.ReactNode }> = ({ kind, title, action, children }) => (
  <div className={`${styles.message} ${styles[kind]}`} role={kind === 'error' ? 'alert' : 'status'}>
    <Icon iconName={MESSAGE_ICON[kind]} />
    <div className={styles.messageBody}>
      {title && <strong>{title}</strong>}
      {children && <span>{children}</span>}
    </div>
    {action}
  </div>
);

export interface IButtonProps {
  text: string;
  onClick?: () => void;
  kind?: 'primary' | 'danger' | 'default';
  disabled?: boolean;
}

export const Button: React.FC<IButtonProps> = ({ text, onClick, kind = 'default', disabled }) => (
  <button type="button" className={`${styles.button} ${kind === 'primary' ? styles.primary : kind === 'danger' ? styles.danger : ''}`} onClick={onClick} disabled={disabled}>
    {text}
  </button>
);

export { styles as ui };
