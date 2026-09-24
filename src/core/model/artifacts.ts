export type ArtifactKind =
  | 'group'
  | 'siteField'
  | 'contentType'
  | 'list'
  | 'listField'
  | 'view'
  | 'listSecurity'
  | 'items'
  | 'files'
  | 'page'
  | 'navigation';

/** Common identifier shared by the planner, the engine and the log. */
export interface IArtifactRef {
  kind: ArtifactKind;
  /** e.g. 'list:Projektek', 'field:CJ_Status' */
  key: string;
}

/** An artifact found on the source site by an extractor's discover() call (Setup tree view). */
export interface IDiscoveredArtifact {
  ref: IArtifactRef;
  title: string;
  /** Parent artifact key for tree grouping (e.g. a view under its list). */
  parentKey?: string;
  itemCount?: number;
  sizeBytes?: number;
}
