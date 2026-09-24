import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { throwIfAborted } from '../errors';
import { TokenContext } from '../tokenizer';
import { assignListKeys, isUserList, toSiteRelativeUrl, LIST_SELECT, type IListInfoLike } from './listModel';

export interface ISourceSite {
  tokens: TokenContext;
  webServerRelativeUrl: string;
  /** User lists of the site with their template keys (same keys the ListExtractor writes). */
  lists: Array<{ info: IListInfoLike; key: string; url: string }>;
}

/**
 * Reads the source site once and builds the token context every extractor tokenizes with:
 * {site}, {siterelative}, {sitename}, and {listkey:K} / {listurl:K} for each user list.
 */
export async function loadSourceSite(sp: SPFI, signal?: AbortSignal): Promise<ISourceSite> {
  throwIfAborted(signal);
  const [web, infos] = await Promise.all([
    sp.web.select('Url', 'ServerRelativeUrl', 'Title')<{ Url: string; ServerRelativeUrl: string; Title: string }>(),
    sp.web.lists.select(...LIST_SELECT).expand('RootFolder')<IListInfoLike[]>()
  ]);
  throwIfAborted(signal);

  const tokens = TokenContext.forSite({ absoluteUrl: web.Url, serverRelativeUrl: web.ServerRelativeUrl, title: web.Title });
  const userLists = infos.filter(isUserList).map((info) => ({ info, url: toSiteRelativeUrl(info.RootFolder.ServerRelativeUrl, web.ServerRelativeUrl) }));
  const keys = assignListKeys(userLists.map((l) => l.url));
  const lists = userLists.map((l) => ({ ...l, key: keys[l.url] }));
  lists.forEach((l) => {
    tokens.set('listkey', l.key, l.info.Id);
    tokens.set('listurl', l.key, l.url);
  });
  return { tokens, webServerRelativeUrl: web.ServerRelativeUrl, lists };
}
