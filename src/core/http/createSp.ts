import { spfi, type SPFI } from '@pnp/sp';
import { SPFx, type ISPFXContext } from '@pnp/sp/behaviors/spfx';
import { CopyJetBehavior } from './behavior';
import type { IRetryOptions } from './retry';

/**
 * Creates a PnPjs root for a site. `context` is the web part's WebPartContext; it is typed structurally
 * (ISPFXContext) so the core does not depend on @microsoft/sp-webpart-base.
 * Source and target sites each get their own instance.
 */
export function createSp(context: ISPFXContext, siteUrl?: string, retry?: Partial<IRetryOptions>): SPFI {
  const url = siteUrl || context.pageContext.web.absoluteUrl;
  return spfi(url).using(SPFx(context), CopyJetBehavior(retry));
}
