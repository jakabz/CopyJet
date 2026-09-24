import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/files';
import { throwIfAborted } from '../errors';
import type { ITemplateReader } from '../model';
import { openTemplate } from './json';

/**
 * Opens a template stored in SharePoint (any site of the tenant, absolute or server-relative URL, sharing
 * links included) – e.g. the CopyJetTemplates library of the source site.
 */
export async function openTemplateFromUrl(sp: SPFI, url: string, signal?: AbortSignal): Promise<ITemplateReader> {
  throwIfAborted(signal);
  const blob = await sp.web.getFileByUrl(url.trim()).getBlob();
  return openTemplate(blob, signal);
}
