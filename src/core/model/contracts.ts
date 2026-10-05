import type { SPFI } from '@pnp/sp';
import type { ArtifactKind, IArtifactRef, IDiscoveredArtifact } from './artifacts';
import type { ICopyJetTemplate } from './generated/copyjet.v1';
import type { TokenContext } from '../tokenizer/TokenContext';
import type { Logger } from '../logger/Logger';
import type { PrincipalMapper } from '../mapping/PrincipalMapper';
import type { TermMapper } from '../mapping/TermMapper';

/** Receives extracted data; hides the package format (.json / .zip) from the extractors. */
export interface ITemplateWriter {
  /** The manifest being built; extractors append their definitions to it. */
  readonly manifest: ICopyJetTemplate;
  /** Adds a JSON entry to the package (e.g. items/<listkey>.json). */
  addJson(path: string, value: unknown): void;
  /** Adds a binary entry to the package (e.g. files/<listkey>/<path>). */
  addBlob(path: string, blob: Blob): void;
  /** Validates the manifest and produces the downloadable template. */
  finalize(signal?: AbortSignal): Promise<Blob>;
}

/** Read access to a loaded template package; entries are loaded lazily. */
export interface ITemplateReader {
  readonly manifest: ICopyJetTemplate;
  /** The manifest as stored in the file, before migration (what meta.checksum was computed over). */
  readonly storedManifest: ICopyJetTemplate;
  /** Paths of the package entries besides manifest.json (none for a .json template). */
  entries(): string[];
  has(path: string): boolean;
  getJson<T>(path: string, signal?: AbortSignal): Promise<T>;
  getBlob(path: string, signal?: AbortSignal): Promise<Blob>;
}

export interface IExtractOptions {
  includeContent: boolean;
  includeVersions: boolean;
  includeMembers: boolean;
  /** Items keep their author, editor and dates (default true). */
  preserveAuthors?: boolean;
  /** Libraries (template keys) whose earlier file versions are copied too (default: none). */
  versionsFor?: string[];
  /** Files larger than this stay out of the package, with a warning (default 250 MB). */
  maxFileBytes?: number;
  /** Source-site values to tokenize (site URLs, list GUIDs → {listkey:X} ...). */
  tokens: TokenContext;
  log: Logger;
  signal?: AbortSignal;
}

export interface IExtractor<TDef> {
  kind: ArtifactKind;
  discover(sp: SPFI, signal?: AbortSignal): Promise<IDiscoveredArtifact[]>;
  /** Artifacts that must also be included for this definition to install. */
  dependencies(def: TDef): IArtifactRef[];
  extract(sp: SPFI, refs: IArtifactRef[], opts: IExtractOptions, out: ITemplateWriter): Promise<void>;
}

export type DiffStatus = 'new' | 'same' | 'different' | 'unsupported';
export type ConflictMode = 'skip' | 'update' | 'rename';

export interface IDiffResult {
  ref: IArtifactRef;
  status: DiffStatus;
  /** Human-readable differences for the preview (property names, not localized text). */
  changes?: string[];
}

export interface IApplyResult {
  ref: IArtifactRef;
  outcome: 'created' | 'updated' | 'skipped';
  /** Created identifiers to register as tokens, e.g. { listkey: { Projektek: '<guid>' } }. */
  tokens?: Record<string, Record<string, string>>;
}

/** Shared state of content steps (items, files) within one install run. */
export interface IContentContext {
  /** The loaded package: item and file entries are read from it. */
  reader: ITemplateReader;
  /** Source item ID → target item ID per list key; fills as item steps run, used to resolve lookups. */
  idMaps: { [listKey: string]: { [sourceId: number]: number } };
  /** Source IDs of the item chunk being written per list key, saved before the write (resume, spike 14). */
  pendingItems?: { [listKey: string]: number[] };
  /** Template principals → target users ({principal:key} tokens). */
  principals: PrincipalMapper;
  /** Template terms → target terms (Managed Metadata values); without it those values are not written. */
  terms?: TermMapper;
}

export interface IInstallContext {
  targetSiteUrl: string;
  /** Target-site values; providers register the identifiers they create. */
  tokens: TokenContext;
  log: Logger;
  signal?: AbortSignal;
  /** Present when the template carries content (items, files). */
  content?: IContentContext;
  /** Saves the run state in the middle of a long step (the install sets it; content providers call it per chunk). */
  checkpoint?: () => Promise<void>;
}

export interface IProvider<TDef> {
  kind: ArtifactKind;
  diff(sp: SPFI, def: TDef, ctx: IInstallContext): Promise<IDiffResult>;
  apply(sp: SPFI, def: TDef, mode: ConflictMode, ctx: IInstallContext): Promise<IApplyResult>;
}
