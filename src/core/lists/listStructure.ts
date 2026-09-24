import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/folders';
import '@pnp/sp/content-types';
import { normalizeContentTypeId, parentIdOf } from '../contentTypes';
import { throwIfAborted } from '../errors';

/** Folders SharePoint keeps in a list's root for itself (forms, attachments, picture thumbnails …). */
function isSystemRootFolder(name: string): boolean {
  return name === 'Forms' || name === 'Attachments' || name.charAt(0) === '_';
}

/** All user folders of a list, relative to its root ('2026', '2026/Q3'), parents before children. */
export async function listFolderPaths(sp: SPFI, listServerRelativeUrl: string, signal?: AbortSignal): Promise<string[]> {
  const out: string[] = [];
  const queue: string[] = [''];
  while (queue.length) {
    throwIfAborted(signal);
    const rel = queue.shift()!;
    const folders = await sp.web.getFolderByServerRelativePath(rel ? `${listServerRelativeUrl}/${rel}` : listServerRelativeUrl).folders.select('Name')<Array<{ Name: string }>>();
    folders
      .filter((f) => rel !== '' || !isSystemRootFolder(f.Name))
      .forEach((f) => {
        const path = rel ? `${rel}/${f.Name}` : f.Name;
        out.push(path);
        queue.push(path);
      });
  }
  return out;
}

/**
 * Folders to create so that every wanted path exists: wanted paths plus their parents, minus existing ones,
 * shallowest first (SharePoint does not create missing parents, spike 04).
 */
export function missingFolders(wanted: string[], existing: string[]): string[] {
  const have: { [p: string]: boolean } = {};
  existing.forEach((p) => (have[p.toLowerCase()] = true));
  const need: string[] = [];
  wanted.forEach((path) => {
    const parts = path.split('/').filter((x) => x);
    for (let i = 1; i <= parts.length; i++) {
      const p = parts.slice(0, i).join('/');
      if (!have[p.toLowerCase()]) {
        have[p.toLowerCase()] = true;
        need.push(p);
      }
    }
  });
  return need.sort((a, b) => a.split('/').length - b.split('/').length);
}

export interface IListContentTypes {
  /** Visible list content type IDs in order; the first is the default. */
  ordered: string[];
  /** Every list content type ID (including hidden ones such as Folder). */
  all: string[];
}

export async function readListContentTypes(sp: SPFI, listServerRelativeUrl: string): Promise<IListContentTypes> {
  const list = sp.web.getList(listServerRelativeUrl);
  const [root, cts] = await Promise.all([
    list.rootFolder.select('ContentTypeOrder')<{ ContentTypeOrder?: Array<{ StringValue: string }> }>(),
    list.contentTypes.select('StringId')<Array<{ StringId: string }>>()
  ]);
  return {
    ordered: (root.ContentTypeOrder || []).map((c) => normalizeContentTypeId(c.StringValue)),
    all: cts.map((c) => normalizeContentTypeId(c.StringId))
  };
}

/** A list content type's site-level ID: list types are site ID + "00" + GUID (spike 04). */
export function siteContentTypeIdOf(listContentTypeId: string): string {
  return parentIdOf(listContentTypeId) || normalizeContentTypeId(listContentTypeId);
}

/** The list content type (if any) that was created from the given site content type. */
export function listContentTypeFor(siteContentTypeId: string, listContentTypeIds: string[]): string | undefined {
  const site = normalizeContentTypeId(siteContentTypeId);
  return listContentTypeIds.filter((id) => siteContentTypeIdOf(id) === site)[0];
}
