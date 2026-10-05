import * as strings from 'CopyJetInstallWebPartStrings';
import type { IListFilesDef } from '../../../core/files';
import type { INavStepDef } from '../../../core/navigation';
import type { IPageDef } from '../../../core/pages';
import type { IListItemsDef } from '../../../core/items';
import type { IListFieldDef, IListViewDef } from '../../../core/lists';
import type { ArtifactKind, IContentType, ICopyJetTemplate, IField, IGroup, IList } from '../../../core/model';
import type { IPlanStep } from '../../../core/planner';

const DOCUMENT_LIBRARY = 101;

/** Type column of the preview: lists and libraries are told apart as in the mockups. */
export function kindLabel(kind: ArtifactKind, def?: unknown): string {
  const labels: Partial<Record<ArtifactKind, string>> = {
    group: strings.KindGroup,
    siteField: strings.KindSiteField,
    contentType: strings.KindContentType,
    list: strings.KindList,
    listField: strings.KindListField,
    view: strings.KindView,
    items: strings.KindItems,
    files: strings.KindFiles,
    page: strings.KindPage,
    itemLookups: strings.KindItemLookups,
    navigation: strings.KindNavigation
  };
  if (kind === 'list' && def && (def as IList).template === DOCUMENT_LIBRARY) return strings.KindLibrary;
  return labels[kind] || kind;
}

/** Display name of a plan step, as the user will see it on the target site. */
export function stepTitle(step: IPlanStep, template: ICopyJetTemplate, targetSiteTitle: string): string {
  const listTitle = (key: string): string => (template.lists.filter((l) => l.key === key)[0] || { title: key }).title;
  switch (step.ref.kind) {
    case 'group':
      return (step.def as IGroup).title.replace('{sitename}', targetSiteTitle);
    case 'siteField':
      return (step.def as IField).title;
    case 'contentType':
      return (step.def as IContentType).name;
    case 'list':
      return (step.def as IList).title;
    case 'listField': {
      const d = step.def as IListFieldDef;
      return `${listTitle(d.listKey)} / ${d.field.title}`;
    }
    case 'view': {
      const d = step.def as IListViewDef;
      return `${listTitle(d.listKey)} / ${d.view.title}`;
    }
    case 'items':
    case 'itemLookups':
      return listTitle((step.def as IListItemsDef).listKey);
    case 'files':
      return listTitle((step.def as IListFilesDef).listKey);
    case 'page':
      return (step.def as IPageDef).title.replace('{sitename}', targetSiteTitle);
    case 'navigation': {
      const d = step.def as INavStepDef;
      if (d.part === 'homePage') return format(strings.NavHomePage, d.page);
      return d.part === 'quickLaunch' ? strings.NavQuickLaunch : strings.NavTopNavigation;
    }
    default:
      return step.ref.key;
  }
}

/** "{0} / {1}" style formatting for loc strings. */
export function format(template: string, ...args: Array<string | number>): string {
  return template.replace(/\{(\d+)\}/g, (m, i) => (args[Number(i)] !== undefined ? String(args[Number(i)]) : m));
}

/** Human-readable reason for a diff change code (see providers); "code:arg" fills {0} with arg. */
export function reasonText(code: string): string {
  const i = code.indexOf(':');
  const name = i < 0 ? code : code.slice(0, i);
  const text = (strings as unknown as Record<string, string>)[`Reason${name}`];
  return text ? format(text, i < 0 ? '' : code.slice(i + 1)) : code;
}

export const logLabels = {
  title: strings.LogTitle,
  allLevels: strings.LogAll,
  levels: { info: strings.LogInfo, warn: strings.LogWarn, error: strings.LogError },
  empty: strings.LogEmpty,
  exportCsv: strings.DownloadCsv,
  exportJson: strings.DownloadJson
};
