import type { SPFI } from '@pnp/sp';

export interface ICopyJetInstallProps {
  /** PnPjs root for the current (target) site, created by the web part. */
  sp: SPFI;
  siteTitle: string;
  siteUrl: string;
}
