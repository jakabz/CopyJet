import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/folders';
import '@pnp/sp/content-types';
import { normalizeContentTypeId, parentIdOf } from '../contentTypes';
import '@pnp/sp/items';
import { CopyJetError, throwIfAborted } from '../errors';

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
 * Folders of a custom list that are real list folders (items with FSObjType 1), relative to the list root.
 * A bare SPFolder (what web.folders.add makes in a custom list) is not one: items cannot be added to it
 * (spike 08 F). Falls back to the folder tree when the filter is refused (large list without index).
 */
export async function listFolderItemPaths(sp: SPFI, listServerRelativeUrl: string, signal?: AbortSignal): Promise<string[]> {
  throwIfAborted(signal);
  try {
    const rows = await sp.web.getList(listServerRelativeUrl).items.filter('FSObjType eq 1').select('FileRef').top(5000)<Array<{ FileRef: string }>>();
    const root = `${listServerRelativeUrl.toLowerCase()}/`;
    return rows
      .map((r) => r.FileRef || '')
      .filter((ref) => ref.toLowerCase().indexOf(root) === 0)
      .map((ref) => ref.slice(root.length))
      .sort((a, b) => a.split('/').length - b.split('/').length);
  } catch {
    return listFolderPaths(sp, listServerRelativeUrl, signal);
  }
}

/** Existing folders as CopyJet sees them: list folders for custom lists, the folder tree for libraries. */
export function existingFolderPaths(sp: SPFI, listServerRelativeUrl: string, isLibrary: boolean, signal?: AbortSignal): Promise<string[]> {
  return isLibrary ? listFolderPaths(sp, listServerRelativeUrl, signal) : listFolderItemPaths(sp, listServerRelativeUrl, signal);
}

/**
 * Creates one folder whose parent exists. Custom lists get a list folder through
 * AddValidateUpdateItemUsingPath (UnderlyingObjectType 1, spike 08 E); libraries keep web.folders.add.
 */
export async function createFolder(sp: SPFI, listServerRelativeUrl: string, path: string, isLibrary: boolean): Promise<void> {
  if (isLibrary) {
    await sp.web.folders.addUsingPath(`${listServerRelativeUrl}/${path}`);
    return;
  }
  const i = path.lastIndexOf('/');
  const parent = i < 0 ? listServerRelativeUrl : `${listServerRelativeUrl}/${path.slice(0, i)}`;
  const leaf = path.slice(i + 1);
  const result = await sp.web
    .getList(listServerRelativeUrl)
    .addValidateUpdateItemUsingPath([{ FieldName: 'Title', FieldValue: leaf }], parent, false, undefined, { leafName: leaf, objectType: 1 });
  const failed = (result || []).filter((v) => v.HasException);
  if (failed.length) {
    throw new CopyJetError('FOLDER_CREATE_FAILED', `Folder ${path} could not be created: ${failed.map((v) => `${v.FieldName}: ${v.ErrorMessage}`).join('; ')}`, failed);
  }
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
  /** List content type ID → the site content type it was made from (REST Parent). */
  parentOf: { [listContentTypeId: string]: string };
}

export async function readListContentTypes(sp: SPFI, listServerRelativeUrl: string): Promise<IListContentTypes> {
  const list = sp.web.getList(listServerRelativeUrl);
  const [root, cts] = await Promise.all([
    list.rootFolder.select('ContentTypeOrder')<{ ContentTypeOrder?: Array<{ StringValue: string }> }>(),
    list.contentTypes.select('StringId', 'Parent/StringId').expand('Parent')<Array<{ StringId: string; Parent?: { StringId?: string } }>>()
  ]);
  const parentOf: { [id: string]: string } = {};
  cts.forEach((c) => {
    if (c.Parent && c.Parent.StringId) parentOf[normalizeContentTypeId(c.StringId)] = normalizeContentTypeId(c.Parent.StringId);
  });
  return {
    ordered: (root.ContentTypeOrder || []).map((c) => normalizeContentTypeId(c.StringValue)),
    all: cts.map((c) => normalizeContentTypeId(c.StringId)),
    parentOf
  };
}

/**
 * A list content type's site-level ID. SharePoint's Parent is authoritative: the list copy of Item can be
 * 0x01 + 00 + GUID + 00 + GUID (spike 08 F), so cutting one "00" + GUID is only the fallback when the
 * parent is unknown (spike 04's lists: site ID + "00" + GUID).
 */
export function siteContentTypeIdOf(listContentTypeId: string, cts?: IListContentTypes): string {
  const id = normalizeContentTypeId(listContentTypeId);
  if (cts && cts.parentOf[id]) return cts.parentOf[id];
  return parentIdOf(listContentTypeId) || id;
}

/** The list content type (if any) that was created from the given site content type. */
export function listContentTypeFor(siteContentTypeId: string, cts: IListContentTypes): string | undefined {
  const site = normalizeContentTypeId(siteContentTypeId);
  return cts.all.filter((id) => siteContentTypeIdOf(id, cts) === site)[0];
}
