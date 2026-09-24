import * as React from 'react';
import { Stack, Text } from '@fluentui/react';
import styles from './WizardShell.module.scss';

export interface IWizardShellProps {
  title: string;
  steps: string[];
  /** 0-based index of the current step. */
  current: number;
  footer?: React.ReactNode;
  children?: React.ReactNode;
}

/** Common wizard frame of the Setup and Install web parts: title, step indicator, content, footer. */
export const WizardShell: React.FC<IWizardShellProps> = ({ title, steps, current, footer, children }) => (
  <section className={styles.shell}>
    <Stack tokens={{ childrenGap: 16 }}>
      <Text variant="xLarge" as="h2">
        {title}
      </Text>
      <ol className={styles.steps}>
        {steps.map((s, i) => (
          <li key={s} className={i === current ? styles.current : i < current ? styles.done : styles.todo}>
            <span className={styles.number}>{i + 1}</span>
            {s}
          </li>
        ))}
      </ol>
      <div>{children}</div>
      {footer && (
        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
          {footer}
        </Stack>
      )}
    </Stack>
  </section>
);
