import type { ICopyJetTemplate } from '../model';
import { lookupTargets } from '../items/itemsModel';

/** The files of one document library (provider input of the 'files' step). */
export interface IListFilesDef {
  listKey: string;
  /** Site-relative library URL from the template; a renamed copy is found through the {listurl:K} token. */
  listUrl: string;
  /** Package entry holding the file list (files/<listkey>/_meta.json). */
  source: string;
  /** Keys of lists the library's lookup columns point to (itself included). */
  lookupTargets: string[];
}

/** Up to this size a file is uploaded in one request, above it in chunks of this size (rendszerterv §7). */
export const CHUNK_SIZE = 10 * 1024 * 1024;

/** Largest file the Setup puts into a package by default (the "Max. fájlméret" option). */
export const DEFAULT_MAX_FILE_BYTES = 250 * 1024 * 1024;

export const filesKey = (listKey: string): string => `files:${listKey}`;
/** The folder of a library's entries; manifest content.source. */
export const filesFolderPath = (listKey: string): string => `files/${listKey}/`;
export const filesMetaPath = (listKey: string): string => `files/${listKey}/_meta.json`;

/** ".." is not allowed in package paths (schema packagePath); runs of dots become one. */
const safe = (path: string): string => path.replace(/\.{2,}/g, '.');

/** Package path of a file's current content: files/<listkey>/<path in the library>. */
export function fileEntryPath(listKey: string, path: string): string {
  return `files/${listKey}/${safe(path)}`;
}

/** Package path of an earlier version: files/<listkey>/_v/<label>/<path in the library>. */
export function fileVersionEntryPath(listKey: string, label: string, path: string): string {
  return `files/${listKey}/_v/${label}/${safe(path)}`;
}

/** Libraries of the template whose files are in the package (embedded mode). */
export function listFilesDefs(template: ICopyJetTemplate): IListFilesDef[] {
  return template.lists
    .filter((l) => l.content && l.content.mode === 'files' && l.content.sourceMode !== 'reference' && !!l.content.source)
    .map((l) => ({ listKey: l.key, listUrl: l.url, source: filesMetaPath(l.key), lookupTargets: lookupTargets(l) }));
}
