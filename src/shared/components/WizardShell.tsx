import * as React from 'react';
import styles from './ui.module.scss';

export interface IWizardShellProps {
  /** "CopyJet Setup" / "CopyJet Install" */
  title: string;
  /** Name of the current screen next to the title ("Forrás kiválasztása"). */
  subtitle: string;
  steps: string[];
  /** 0-based index of the current step. */
  current: number;
  /** Buttons on the left of the footer (e.g. Stop). */
  footerStart?: React.ReactNode;
  /** Buttons on the right of the footer (Back / Next). */
  footerEnd?: React.ReactNode;
  children?: React.ReactNode;
}

/** Wizard card of the docs/ui mockups: title row, step indicator with lines, body, footer bar. */
export const WizardShell: React.FC<IWizardShellProps> = ({ title, subtitle, steps, current, footerStart, footerEnd, children }) => (
  <section className={styles.card}>
    <div className={styles.header}>
      <div className={styles.titleRow}>
        <h1>{title}</h1>
        <span className={styles.muted}>{subtitle}</span>
      </div>
      <ol className={styles.steps} style={{ listStyle: 'none', margin: 0, padding: 0 }}>
        {steps.map((s, i) => (
          <React.Fragment key={s}>
            {i > 0 && <li className={styles.stepLine} aria-hidden="true" />}
            <li className={`${styles.step} ${i === current ? styles.stepCurrent : i < current ? styles.stepDone : ''}`} aria-current={i === current ? 'step' : undefined}>
              <span className={styles.stepNumber}>{i + 1}</span>
              <span>{s}</span>
            </li>
          </React.Fragment>
        ))}
      </ol>
    </div>
    <div className={styles.body}>{children}</div>
    {(footerStart || footerEnd) && (
      <div className={styles.footer}>
        {footerStart}
        <div className={styles.spacer} />
        {footerEnd}
      </div>
    )}
  </section>
);
