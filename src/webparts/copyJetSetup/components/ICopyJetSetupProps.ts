import type { SPFI } from '@pnp/sp';

export interface ICopyJetSetupProps {
  /** PnPjs root for the current (source) site, created by the web part. */
  sp: SPFI;
  siteTitle: string;
  /** Login or e-mail written into the template's meta.createdBy. */
  createdBy: string;
}
