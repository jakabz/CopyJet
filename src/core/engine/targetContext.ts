import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import { throwIfAborted } from '../errors';
import { readAssociatedGroups } from '../groups';
import type { Logger } from '../logger/Logger';
import type { PrincipalMapper } from '../mapping/PrincipalMapper';
import type { IInstallContext, ITemplateReader } from '../model';
import { TokenContext } from '../tokenizer';

/** What content steps need: the package and the principal mapper (reused from the mapping step). */
export interface IInstallContent {
  reader: ITemplateReader;
  principals: PrincipalMapper;
}

/**
 * Install context for a target web: {site}, {siterelative}, {sitename} and the associated groups'
 * {associated…group} IDs. Providers add the identifiers they create as they run. With `content`, the
 * context also carries the package and fresh ID maps for item steps.
 */
export async function createInstallContext(sp: SPFI, log: Logger, signal?: AbortSignal, content?: IInstallContent): Promise<IInstallContext> {
  throwIfAborted(signal);
  const [web, associated] = await Promise.all([
    sp.web.select('Url', 'ServerRelativeUrl', 'Title')<{ Url: string; ServerRelativeUrl: string; Title: string }>(),
    readAssociatedGroups(sp)
  ]);
  const tokens = TokenContext.forSite({ absoluteUrl: web.Url, serverRelativeUrl: web.ServerRelativeUrl, title: web.Title });
  if (associated.owner !== undefined) tokens.set('associatedownergroup', undefined, String(associated.owner));
  if (associated.member !== undefined) tokens.set('associatedmembergroup', undefined, String(associated.member));
  if (associated.visitor !== undefined) tokens.set('associatedvisitorgroup', undefined, String(associated.visitor));
  const ctx: IInstallContext = { targetSiteUrl: web.Url, tokens, log, signal };
  if (content) ctx.content = { reader: content.reader, principals: content.principals, idMaps: {} };
  return ctx;
}
