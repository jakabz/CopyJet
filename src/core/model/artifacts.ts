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
  /** Display grouping on the source site (e.g. site column group). */
  group?: string;
  itemCount?: number;
  sizeBytes?: number;
  /** Lists: the SharePoint list template (100 = list, 101 = library …), for grouping in the Setup. */
  listTemplate?: number;
  /** Set when CopyJet cannot copy the artifact (reason code, e.g. 'LIST_TEMPLATE_UNSUPPORTED'); shown disabled. */
  unsupported?: string;
}
