import * as React from 'react';
import { Stack, Text } from '@fluentui/react';
import * as strings from 'CopyJetInstallWebPartStrings';
import styles from './CopyJetInstall.module.scss';
import type { ICopyJetInstallProps } from './ICopyJetInstallProps';

const CopyJetInstall: React.FC<ICopyJetInstallProps> = ({ siteUrl }) => (
  <section className={styles.copyJetInstall}>
    <Stack tokens={{ childrenGap: 12 }}>
      <Text variant="xLarge" as="h2">
        {strings.Title}
      </Text>
      <Text>{strings.Intro}</Text>
      <Text>
        {strings.SiteLabel}: <strong>{siteUrl}</strong>
      </Text>
    </Stack>
  </section>
);

export default CopyJetInstall;
