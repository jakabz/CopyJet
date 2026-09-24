import * as React from 'react';
import { MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';
import * as strings from 'CopyJetSetupWebPartStrings';
import { probeSite, type ISiteProbe } from '../../../core/http';
import styles from './CopyJetSetup.module.scss';
import type { ICopyJetSetupProps } from './ICopyJetSetupProps';

type ProbeState =
  | { status: 'idle' }
  | { status: 'loading' }
  | { status: 'done'; result: ISiteProbe }
  | { status: 'error'; message: string };

const CopyJetSetup: React.FC<ICopyJetSetupProps> = ({ sp }) => {
  const [state, setState] = React.useState<ProbeState>({ status: 'idle' });
  const abortRef = React.useRef<AbortController | undefined>(undefined);

  React.useEffect(() => () => abortRef.current?.abort(), []);

  const onProbe = React.useCallback((): void => {
    abortRef.current?.abort();
    const ac = new AbortController();
    abortRef.current = ac;
    setState({ status: 'loading' });
    probeSite(sp, ac.signal).then(
      (result) => {
        if (!ac.signal.aborted) setState({ status: 'done', result });
      },
      (err: unknown) => {
        if (!ac.signal.aborted) setState({ status: 'error', message: err instanceof Error ? err.message : String(err) });
      }
    );
  }, [sp]);

  return (
    <section className={styles.copyJetSetup}>
      <Stack tokens={{ childrenGap: 12 }}>
        <Text variant="xLarge" as="h2">
          {strings.Title}
        </Text>
        <Text>{strings.Intro}</Text>
        <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
          <PrimaryButton text={strings.ProbeButton} onClick={onProbe} disabled={state.status === 'loading'} />
          {state.status === 'loading' && <Spinner size={SpinnerSize.small} label={strings.Probing} labelPosition="right" />}
        </Stack>
        {state.status === 'done' && (
          <MessageBar messageBarType={MessageBarType.success}>
            <div>
              {strings.SiteTitleLabel}: <strong>{state.result.title}</strong>
            </div>
            <div>
              {strings.ListCountLabel}: <strong>{state.result.visibleListCount}</strong>
            </div>
          </MessageBar>
        )}
        {state.status === 'error' && (
          <MessageBar messageBarType={MessageBarType.error}>
            {strings.ProbeError}: {state.message}
          </MessageBar>
        )}
      </Stack>
    </section>
  );
};

export default CopyJetSetup;
