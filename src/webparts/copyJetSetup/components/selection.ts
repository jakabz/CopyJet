import type { ArtifactKind, IArtifactRef, IDiscoveredArtifact } from '../../../core/model';

/** Sections of the selection tree as in docs/ui (setup-1): lists, libraries, groups, columns and types. */
export type Category = 'lists' | 'libraries' | 'groups' | 'columns';
export const CATEGORIES: Category[] = ['lists', 'libraries', 'groups', 'columns'];

const DOCUMENT_LIBRARY = 101;

export interface ITreeNode {
  item: IDiscoveredArtifact;
  /** List columns and views under a list. */
  children: IDiscoveredArtifact[];
}

const CHILD_KINDS: ArtifactKind[] = ['listField', 'view'];

export function categoryOf(a: IDiscoveredArtifact): Category | undefined {
  switch (a.ref.kind) {
    case 'list':
      return a.listTemplate === DOCUMENT_LIBRARY ? 'libraries' : 'lists';
    case 'group':
      return 'groups';
    case 'siteField':
    case 'contentType':
      return 'columns';
    default:
      return undefined;
  }
}

/** Groups the discovery into sections; list columns and views go under their list (parentKey). */
export function buildTree(artifacts: IDiscoveredArtifact[]): Record<Category, ITreeNode[]> {
  const tree: Record<Category, ITreeNode[]> = { lists: [], libraries: [], groups: [], columns: [] };
  artifacts.forEach((a) => {
    const c = categoryOf(a);
    if (c) tree[c].push({ item: a, children: artifacts.filter((x) => CHILD_KINDS.indexOf(x.ref.kind) >= 0 && x.parentKey === a.ref.key) });
  });
  const order = (n: ITreeNode): string => `${n.item.ref.kind === 'contentType' ? 1 : 0}${n.item.title.toLowerCase()}`;
  CATEGORIES.forEach((c) => tree[c].sort((a, b) => order(a).localeCompare(order(b))));
  return tree;
}

/** Nodes whose title (or a child's) contains the search text; children are filtered too. */
export function filterTree(nodes: ITreeNode[], search: string): ITreeNode[] {
  const q = search.trim().toLowerCase();
  if (!q) return nodes;
  const hit = (a: IDiscoveredArtifact): boolean => a.title.toLowerCase().indexOf(q) >= 0;
  return nodes
    .map((n) => (hit(n.item) ? n : { item: n.item, children: n.children.filter(hit) }))
    .filter((n) => hit(n.item) || n.children.length > 0);
}

const selectable = (a: IDiscoveredArtifact): boolean => !a.unsupported;

const keysOf = (node: ITreeNode): string[] => [node.item].concat(node.children).filter(selectable).map((a) => a.ref.key);

/** Switches an item (and, for a list, its supported columns and views) on or off. */
export function toggle(selected: string[], node: ITreeNode, on: boolean): string[] {
  const keys = keysOf(node);
  const rest = selected.filter((k) => keys.indexOf(k) < 0);
  return on ? rest.concat(keys) : rest;
}

/** Switches every item of a section. */
export function toggleAll(selected: string[], nodes: ITreeNode[], on: boolean): string[] {
  return nodes.reduce((acc, n) => toggle(acc, n, on), selected);
}

/** Switches a single list column or view; switching one on also switches its list on. */
export function toggleChild(selected: string[], node: ITreeNode, child: IDiscoveredArtifact, on: boolean): string[] {
  const rest = selected.filter((k) => k !== child.ref.key);
  if (!on) return rest;
  return rest.concat(child.ref.key, selected.indexOf(node.item.ref.key) < 0 ? [node.item.ref.key] : []);
}

/** "Whole site": every supported artifact. */
export function selectAll(artifacts: IDiscoveredArtifact[]): string[] {
  return artifacts.filter(selectable).map((a) => a.ref.key);
}

export function selectedRefs(artifacts: IDiscoveredArtifact[], selected: string[]): IArtifactRef[] {
  return artifacts.filter((a) => selected.indexOf(a.ref.key) >= 0).map((a) => a.ref);
}

export function selectedArtifacts(artifacts: IDiscoveredArtifact[], selected: string[]): IDiscoveredArtifact[] {
  return artifacts.filter((a) => selected.indexOf(a.ref.key) >= 0);
}

/** "{0} / {1} kiválasztva" style formatting for loc strings. */
export function format(template: string, ...args: Array<string | number>): string {
  return template.replace(/\{(\d+)\}/g, (m, i) => (args[Number(i)] !== undefined ? String(args[Number(i)]) : m));
}
