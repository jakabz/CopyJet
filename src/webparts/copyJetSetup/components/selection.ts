import type { ArtifactKind, IArtifactRef, IDiscoveredArtifact } from '../../../core/model';

/** Top-level sections of the selection tree, in display order. */
export type Category = 'list' | 'group' | 'siteField' | 'contentType';
export const CATEGORIES: Category[] = ['list', 'group', 'siteField', 'contentType'];

export interface ITreeNode {
  item: IDiscoveredArtifact;
  /** List columns and views under a list. */
  children: IDiscoveredArtifact[];
}

const CHILD_KINDS: ArtifactKind[] = ['listField', 'view'];

/** Groups the discovery into sections; list columns and views go under their list (parentKey). */
export function buildTree(artifacts: IDiscoveredArtifact[]): Record<Category, ITreeNode[]> {
  const tree: Record<Category, ITreeNode[]> = { list: [], group: [], siteField: [], contentType: [] };
  const byTitle = (a: ITreeNode, b: ITreeNode): number => a.item.title.localeCompare(b.item.title);
  artifacts
    .filter((a) => CATEGORIES.indexOf(a.ref.kind as Category) >= 0)
    .forEach((a) => tree[a.ref.kind as Category].push({ item: a, children: artifacts.filter((c) => CHILD_KINDS.indexOf(c.ref.kind) >= 0 && c.parentKey === a.ref.key) }));
  CATEGORIES.forEach((c) => tree[c].sort(byTitle));
  return tree;
}

const selectable = (a: IDiscoveredArtifact): boolean => !a.unsupported;

/** Switches an item (and, for a list, its supported columns and views) on or off. */
export function toggle(selected: string[], node: ITreeNode, on: boolean): string[] {
  const keys = [node.item].concat(node.children).filter(selectable).map((a) => a.ref.key);
  const rest = selected.filter((k) => keys.indexOf(k) < 0);
  return on ? rest.concat(keys) : rest;
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

export function countByKind(refs: IArtifactRef[]): Partial<Record<ArtifactKind, number>> {
  const out: Partial<Record<ArtifactKind, number>> = {};
  refs.forEach((r) => (out[r.kind] = (out[r.kind] || 0) + 1));
  return out;
}

/** "{0} / {1} kiválasztva" style formatting for loc strings. */
export function format(template: string, ...args: Array<string | number>): string {
  return template.replace(/\{(\d+)\}/g, (m, i) => (args[Number(i)] !== undefined ? String(args[Number(i)]) : m));
}
