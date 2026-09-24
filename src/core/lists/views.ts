import type { ICopyJetTemplate, IView } from '../model';
import { tokenize, type TokenContext } from '../tokenizer';

/** A view in the context of its list (the provider's unit of work). */
export interface IListViewDef {
  listKey: string;
  /** Site-relative list URL from the template; a renamed copy is found through the {listurl:K} token. */
  listUrl: string;
  view: IView;
}

export interface IViewInfoLike {
  Id: string;
  Title: string;
  ServerRelativeUrl?: string;
  DefaultView?: boolean;
  Hidden?: boolean;
  PersonalView?: boolean;
  ViewType?: string;
  RowLimit?: number;
  Paged?: boolean;
  Scope?: number;
  ViewQuery?: string;
  CustomFormatter?: string;
}

export const VIEW_SELECT = ['Id', 'Title', 'ServerRelativeUrl', 'DefaultView', 'Hidden', 'PersonalView', 'ViewType', 'RowLimit', 'Paged', 'Scope', 'ViewQuery', 'CustomFormatter'];

/** Public, visible views (spike 06: personal views and library system views are excluded). */
export const COPIED_VIEWS_FILTER = 'Hidden eq false and PersonalView eq false';

export const viewKey = (listKey: string, title: string): string => `view:${listKey}/${title}`;

const SCOPES: Array<NonNullable<IView['scope']>> = ['Default', 'Recursive', 'RecursiveAll', 'FilesOnly'];

export function scopeName(scope: number | undefined): IView['scope'] {
  return SCOPES[scope || 0] || 'Default';
}

export function scopeNumber(scope: IView['scope']): number {
  const i = SCOPES.indexOf(scope || 'Default');
  return i < 0 ? 0 : i;
}

/** View types the provider creates: REST can only create HTML views (ViewTypeKind is not on SP.View, spike 06). */
export function isCreatableViewType(viewType: IView['viewType']): boolean {
  return !viewType || viewType === 'HTML';
}

export function toViewDef(info: IViewInfoLike, fields: string[], tokens: TokenContext): IView {
  const view: IView = { title: info.Title, fields: fields.slice() };
  if (info.DefaultView) view.default = true;
  if (info.ViewQuery) view.query = tokenize(info.ViewQuery, tokens);
  if (info.RowLimit) view.rowLimit = Math.max(1, Math.min(5000, info.RowLimit));
  view.paged = !!info.Paged;
  view.scope = scopeName(info.Scope);
  if (info.ViewType === 'HTML' || info.ViewType === 'GRID' || info.ViewType === 'CALENDAR') view.viewType = info.ViewType;
  if (info.CustomFormatter) view.customFormatter = tokenize(info.CustomFormatter, tokens);
  return view;
}

/** The target view a template view corresponds to: the default view by flag (its title is language-dependent), others by title. */
export function matchView(view: IView, targetViews: IViewInfoLike[]): IViewInfoLike | undefined {
  return view.default ? targetViews.filter((v) => v.DefaultView)[0] : targetViews.filter((v) => v.Title === view.title)[0];
}

/** Differences between a (resolved) template view and the target view. */
export function compareViews(view: IView, target: IViewInfoLike, targetFields: string[]): string[] {
  const changes: string[] = [];
  if (view.title !== target.Title) changes.push('title');
  if ((view.query || '') !== (target.ViewQuery || '')) changes.push('query');
  if (view.rowLimit !== undefined && view.rowLimit !== target.RowLimit) changes.push('rowLimit');
  if (view.paged !== undefined && view.paged !== !!target.Paged) changes.push('paged');
  if (scopeNumber(view.scope) !== (target.Scope || 0)) changes.push('scope');
  if ((view.customFormatter || '') !== (target.CustomFormatter || '')) changes.push('customFormatter');
  if (view.fields.join('|') !== targetFields.join('|')) changes.push('fields');
  return changes;
}

/** All views of a template as provider inputs, in list order. */
export function listViewDefs(template: ICopyJetTemplate): IListViewDef[] {
  const out: IListViewDef[] = [];
  template.lists.forEach((l) => (l.views || []).forEach((view) => out.push({ listKey: l.key, listUrl: l.url, view })));
  return out;
}
