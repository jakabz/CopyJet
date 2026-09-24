import type { SPFI } from '@pnp/sp';

export interface ICopyJetSetupProps {
  /** PnPjs root for the current (source) site, created by the web part. */
  sp: SPFI;
}
