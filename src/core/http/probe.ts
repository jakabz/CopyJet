import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { throwIfAborted } from '../errors';

export interface ISiteProbe {
  title: string;
  url: string;
  /** Lists and libraries with Hidden = false. */
  visibleListCount: number;
}

/** Connection check used by the phase 0 UI: reads the site title and counts its visible lists. */
export async function probeSite(sp: SPFI, signal?: AbortSignal): Promise<ISiteProbe> {
  throwIfAborted(signal);
  const [web, lists] = await Promise.all([
    sp.web.select('Title', 'Url')<{ Title: string; Url: string }>(),
    sp.web.lists.filter('Hidden eq false').select('Id')<Array<{ Id: string }>>()
  ]);
  throwIfAborted(signal);
  return { title: web.Title, url: web.Url, visibleListCount: lists.length };
}
