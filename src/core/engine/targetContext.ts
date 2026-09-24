import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import { throwIfAborted } from '../errors';
import { readAssociatedGroups } from '../groups';
import type { Logger } from '../logger/Logger';
import type { IInstallContext } from '../model';
import { TokenContext } from '../tokenizer';

/**
 * Install context for a target web: {site}, {siterelative}, {sitename} and the associated groups'
 * {associated…group} IDs. Providers add the identifiers they create as they run.
 */
export async function createInstallContext(sp: SPFI, log: Logger, signal?: AbortSignal): Promise<IInstallContext> {
  throwIfAborted(signal);
  const [web, associated] = await Promise.all([
    sp.web.select('Url', 'ServerRelativeUrl', 'Title')<{ Url: string; ServerRelativeUrl: string; Title: string }>(),
    readAssociatedGroups(sp)
  ]);
  const tokens = TokenContext.forSite({ absoluteUrl: web.Url, serverRelativeUrl: web.ServerRelativeUrl, title: web.Title });
  if (associated.owner !== undefined) tokens.set('associatedownergroup', undefined, String(associated.owner));
  if (associated.member !== undefined) tokens.set('associatedmembergroup', undefined, String(associated.member));
  if (associated.visitor !== undefined) tokens.set('associatedvisitorgroup', undefined, String(associated.visitor));
  return { targetSiteUrl: web.Url, tokens, log, signal };
}
