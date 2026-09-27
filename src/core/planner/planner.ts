import { listItemsDefs, type IListItemsDef } from '../items/itemsModel';
import { listFieldDefs, type IListFieldDef } from '../lists/listFields';
import { listViewDefs, type IListViewDef } from '../lists/views';
import type { ArtifactKind, IArtifactRef, IContentType, ICopyJetTemplate, IField, IGroup, IList } from '../model';
import {
  artifactKeys,
  contentTypeDependencies,
  groupDependencies,
  itemLookupsDependencies,
  itemsDependencies,
  listDependencies,
  listFieldDependencies,
  siteFieldDependencies,
  viewDependencies
} from './dependencies';

/** Provider input per kind. */
export type StepDef = IGroup | IField | IContentType | IList | IListFieldDef | IListViewDef | IListItemsDef;

export interface IPlanStep {
  ref: IArtifactRef;
  def: StepDef;
  /** Keys of steps (in this plan) that must succeed first. */
  dependsOn: string[];
  /** 0-based; steps on the same level do not depend on each other and may run in parallel. */
  level: number;
  /**
   * What the step changes. Steps with the same lock run one after another: SharePoint fails concurrent schema
   * changes of one list or of the web's columns/content types with random 500s (0x8007047E, 0x80131904).
   */
  lock: string;
}

export type ExclusionReason = 'disabled' | 'dependencyDisabled' | 'cycle';

export interface IExcludedStep {
  ref: IArtifactRef;
  reason: ExclusionReason;
  /** For 'dependencyDisabled': the disabled step that caused it. */
  cause?: string;
}

export interface IPlan {
  /** Runnable steps, level by level. */
  steps: IPlanStep[];
  levels: IPlanStep[][];
  excluded: IExcludedStep[];
}

export interface IPlanOptions {
  /** Artifact keys the user switched off; their dependents are switched off too. */
  disabled?: string[];
}

/** Install order within a level (rendszerterv §6): groups, site columns, content types, lists, list columns, views, content. */
const KIND_ORDER: ArtifactKind[] = ['group', 'siteField', 'contentType', 'list', 'listField', 'view', 'items', 'itemLookups'];

function node(kind: ArtifactKind, key: string, def: StepDef, deps: IArtifactRef[]): { ref: IArtifactRef; def: StepDef; deps: string[] } {
  return { ref: { kind, key }, def, deps: deps.map((d) => d.key) };
}

/**
 * Every artifact of the template with its unfiltered dependencies (they may point outside the template –
 * the Setup uses them to offer missing dependencies, the planner keeps only in-template edges).
 */
export function templateNodes(template: ICopyJetTemplate): Array<{ ref: IArtifactRef; def: StepDef; deps: string[] }> {
  return [
    ...template.groups.map((g) => node('group', artifactKeys.group(g.key), g, groupDependencies(g))),
    ...template.siteFields.map((f) => node('siteField', artifactKeys.siteField(f.internalName), f, siteFieldDependencies(f))),
    ...template.contentTypes.map((c) => node('contentType', artifactKeys.contentType(c.id), c, contentTypeDependencies(c))),
    ...template.lists.map((l) => node('list', artifactKeys.list(l.key), l, listDependencies(l))),
    ...listFieldDefs(template).map((d) => node('listField', artifactKeys.listField(d.listKey, d.field.internalName), d, listFieldDependencies(d))),
    ...listViewDefs(template).map((d) => node('view', artifactKeys.view(d.listKey, d.view.title), d, viewDependencies(d))),
    ...listItemsDefs(template).map((d) => node('items', artifactKeys.items(d.listKey), d, itemsDependencies(d, template.lists.filter((l) => l.key === d.listKey)[0]))),
    ...listItemsDefs(template)
      .filter((d) => d.lookupTargets.length > 0)
      .map((d) => node('itemLookups', artifactKeys.itemLookups(d.listKey), d, itemLookupsDependencies(d)))
  ];
}

/** The SharePoint schema a step changes (see IPlanStep.lock). */
function lockOf(ref: IArtifactRef, def: StepDef): string {
  switch (ref.kind) {
    case 'listField':
      return `list:${(def as IListFieldDef).listKey}`;
    case 'view':
      return `list:${(def as IListViewDef).listKey}`;
    case 'siteField':
      return 'web:fields';
    case 'contentType':
      return 'web:contentTypes';
    default:
      return ref.key;
  }
}

const byKindThenKey = (a: { ref: IArtifactRef }, b: { ref: IArtifactRef }): number =>
  KIND_ORDER.indexOf(a.ref.kind) - KIND_ORDER.indexOf(b.ref.kind) || (a.ref.key < b.ref.key ? -1 : a.ref.key > b.ref.key ? 1 : 0);

/**
 * Dependency graph of the template → install plan. Only edges between artifacts of the template are kept
 * (built-in content types and columns exist on every target). Levels come from Kahn's algorithm; a step
 * is on the level after its latest dependency. Disabled steps take their dependents with them; steps left
 * in a cycle are excluded rather than guessed at.
 */
export function buildPlan(template: ICopyJetTemplate, options: IPlanOptions = {}): IPlan {
  const all = templateNodes(template);
  const known: { [key: string]: boolean } = {};
  all.forEach((n) => (known[n.ref.key] = true));
  all.forEach((n) => (n.deps = n.deps.filter((d, i, arr) => known[d] && d !== n.ref.key && arr.indexOf(d) === i)));

  // Disabled steps and everything that (transitively) depends on them.
  const excluded: IExcludedStep[] = [];
  const out: { [key: string]: boolean } = {};
  (options.disabled || []).filter((k) => known[k]).forEach((k) => (out[k] = true));
  all.filter((n) => out[n.ref.key]).forEach((n) => excluded.push({ ref: n.ref, reason: 'disabled' }));
  const causeOf: { [key: string]: string } = {};
  Object.keys(out).forEach((k) => (causeOf[k] = k));
  for (let changed = true; changed; ) {
    changed = false;
    all.forEach((n) => {
      if (out[n.ref.key]) return;
      const blocker = n.deps.filter((d) => out[d])[0];
      if (blocker) {
        out[n.ref.key] = true;
        causeOf[n.ref.key] = causeOf[blocker];
        excluded.push({ ref: n.ref, reason: 'dependencyDisabled', cause: causeOf[blocker] });
        changed = true;
      }
    });
  }
  const active = all.filter((n) => !out[n.ref.key]);

  // Kahn's algorithm, level by level.
  const pending: { [key: string]: number } = {};
  const dependents: { [key: string]: string[] } = {};
  active.forEach((n) => {
    pending[n.ref.key] = n.deps.length;
    n.deps.forEach((d) => (dependents[d] = dependents[d] || []).push(n.ref.key));
  });
  const byKey: { [key: string]: (typeof active)[number] } = {};
  active.forEach((n) => (byKey[n.ref.key] = n));

  const levels: IPlanStep[][] = [];
  let ready = active.filter((n) => pending[n.ref.key] === 0);
  while (ready.length) {
    ready.sort(byKindThenKey);
    const level = ready.map((n) => ({ ref: n.ref, def: n.def, dependsOn: n.deps, level: levels.length, lock: lockOf(n.ref, n.def) }));
    levels.push(level);
    const next: typeof active = [];
    ready.forEach((n) =>
      (dependents[n.ref.key] || []).forEach((k) => {
        if (--pending[k] === 0) next.push(byKey[k]);
      })
    );
    ready = next;
  }

  const placed: { [key: string]: boolean } = {};
  levels.forEach((l) => l.forEach((s) => (placed[s.ref.key] = true)));
  active.filter((n) => !placed[n.ref.key]).forEach((n) => excluded.push({ ref: n.ref, reason: 'cycle' }));

  return { steps: ([] as IPlanStep[]).concat(...levels), levels, excluded };
}
