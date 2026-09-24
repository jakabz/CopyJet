import * as React from 'react';
import { Icon } from '@fluentui/react';
import * as strings from 'CopyJetSetupWebPartStrings';
import type { IDiscoveredArtifact } from '../../../core/model';
import { ui } from '../../../shared/components/ui';
import { CATEGORIES, buildTree, categoryOf, filterTree, format, selectAll, selectedArtifacts, toggle, toggleAll, toggleChild, type Category, type ITreeNode } from './selection';
import styles from './CopyJetSetup.module.scss';

export interface ISelectStepProps {
  artifacts: IDiscoveredArtifact[];
  selected: string[];
  onChange: (selected: string[]) => void;
}

const CATEGORY_LABEL: Record<Category, () => string> = {
  lists: () => strings.CategoryLists,
  libraries: () => strings.CategoryLibraries,
  groups: () => strings.CategoryGroups,
  columns: () => strings.CategoryColumnsAndTypes
};

function unsupportedText(code: string): string {
  return (strings as unknown as Record<string, string>)[`Unsupported${code}`] || strings.UnsupportedOther;
}

function meta(a: IDiscoveredArtifact): string {
  if (a.unsupported) return unsupportedText(a.unsupported);
  if (a.itemCount !== undefined) return format(strings.ItemCount, a.itemCount);
  if (a.ref.kind === 'contentType') return strings.TypeContentType;
  return a.group || '';
}

const Chevron: React.FC<{ open: boolean; onClick: () => void; label: string }> = ({ open, onClick, label }) => (
  <button type="button" className={styles.chevron} onClick={onClick} aria-expanded={open} aria-label={label}>
    <Icon iconName={open ? 'ChevronDown' : 'ChevronRight'} />
  </button>
);

/** Step 1 (docs/ui setup-1): search, whole-site switch, tree with sections, "Selected" side panel. */
export const SelectStep: React.FC<ISelectStepProps> = ({ artifacts, selected, onChange }) => {
  const tree = React.useMemo(() => buildTree(artifacts), [artifacts]);
  const [search, setSearch] = React.useState('');
  const [closed, setClosed] = React.useState<{ [key: string]: boolean }>({});
  const [openLists, setOpenLists] = React.useState<{ [key: string]: boolean }>({});
  const all = React.useMemo(() => selectAll(artifacts), [artifacts]);
  const isOn = (a: IDiscoveredArtifact): boolean => selected.indexOf(a.ref.key) >= 0;
  const wholeSite = all.length > 0 && all.every((k) => selected.indexOf(k) >= 0);

  const picked = selectedArtifacts(artifacts, selected);
  const count = (pred: (a: IDiscoveredArtifact) => boolean): number => picked.filter(pred).length;
  const side: Array<[string, number]> = [
    [strings.SideLists, count((a) => categoryOf(a) === 'lists')],
    [strings.SideLibraries, count((a) => categoryOf(a) === 'libraries')],
    [strings.SideGroups, count((a) => a.ref.kind === 'group')],
    [strings.SideSiteColumns, count((a) => a.ref.kind === 'siteField')],
    [strings.SideContentTypes, count((a) => a.ref.kind === 'contentType')],
    [strings.SideColumns, count((a) => a.ref.kind === 'listField')],
    [strings.SideViews, count((a) => a.ref.kind === 'view')]
  ];

  const row = (key: string, level: 0 | 1 | 2, content: React.ReactNode): React.ReactNode => (
    <div key={key} className={`${styles.row} ${styles[`level${level}`]}`}>
      {content}
    </div>
  );

  const renderNode = (n: ITreeNode): React.ReactNode[] => {
    const id = `cj-${n.item.ref.key}`;
    const hasChildren = n.children.length > 0;
    const open = !!openLists[n.item.ref.key] || (!!search.trim() && n.children.length > 0);
    const out: React.ReactNode[] = [
      row(
        n.item.ref.key,
        1,
        <>
          {hasChildren ? (
            <Chevron open={open} label={n.item.title} onClick={() => setOpenLists({ ...openLists, [n.item.ref.key]: !openLists[n.item.ref.key] })} />
          ) : (
            <span className={styles.chevronSpace} />
          )}
          <input className={ui.check} type="checkbox" id={id} checked={isOn(n.item)} disabled={!!n.item.unsupported} onChange={(e) => onChange(toggle(selected, n, e.target.checked))} />
          <label htmlFor={id} className={styles.rowLabel}>
            {n.item.title}
          </label>
          <span className={`${ui.muted} ${styles.rowMeta}`}>{meta(n.item)}</span>
        </>
      )
    ];
    if (open) {
      (['listField', 'view'] as const).forEach((kind) => {
        const items = n.children.filter((c) => c.ref.kind === kind);
        if (!items.length) return;
        out.push(
          row(`${n.item.ref.key}:${kind}`, 2, <span className={`${ui.muted} ${styles.subhead}`}>{kind === 'listField' ? strings.Columns : strings.Views}</span>)
        );
        items.forEach((c) => {
          const cid = `cj-${c.ref.key}`;
          out.push(
            row(
              c.ref.key,
              2,
              <>
                <input className={ui.check} type="checkbox" id={cid} checked={isOn(c)} disabled={!!c.unsupported} onChange={(e) => onChange(toggleChild(selected, n, c, e.target.checked))} />
                <label htmlFor={cid} className={styles.rowLabel}>
                  {c.title}
                </label>
                <span className={`${ui.muted} ${styles.rowMeta}`}>{c.unsupported ? unsupportedText(c.unsupported) : ''}</span>
              </>
            )
          );
        });
      });
    }
    return out;
  };

  return (
    <>
      <div className={styles.toolbar}>
        <input
          className={ui.input}
          type="text"
          value={search}
          placeholder={strings.SearchPlaceholder}
          aria-label={strings.SearchPlaceholder}
          onChange={(e) => setSearch(e.target.value)}
          style={{ width: 360, maxWidth: '100%' }}
        />
        <div className={ui.spacer} />
        <input className={ui.check} type="checkbox" id="cj-whole-site" checked={wholeSite} onChange={(e) => onChange(e.target.checked ? all : [])} />
        <label htmlFor="cj-whole-site" className={ui.strong}>
          {strings.WholeSite}
        </label>
      </div>
      <div className={ui.split}>
        <div className={`${ui.grow} ${ui.box} ${styles.tree}`}>
          {CATEGORIES.filter((c) => tree[c].length > 0).map((c) => {
            const nodes = filterTree(tree[c], search);
            const selectableNodes = tree[c].filter((n) => !n.item.unsupported);
            const on = selectableNodes.filter((n) => isOn(n.item)).length;
            const id = `cj-cat-${c}`;
            return (
              <React.Fragment key={c}>
                {row(
                  id,
                  0,
                  <>
                    <Chevron open={!closed[c]} label={CATEGORY_LABEL[c]()} onClick={() => setClosed({ ...closed, [c]: !closed[c] })} />
                    <input
                      className={ui.check}
                      type="checkbox"
                      id={id}
                      checked={selectableNodes.length > 0 && on === selectableNodes.length}
                      ref={(el) => {
                        if (el) el.indeterminate = on > 0 && on < selectableNodes.length;
                      }}
                      disabled={selectableNodes.length === 0}
                      onChange={(e) => onChange(toggleAll(selected, selectableNodes, e.target.checked))}
                    />
                    <label htmlFor={id} className={`${styles.rowLabel} ${ui.strong}`}>
                      {CATEGORY_LABEL[c]()}
                    </label>
                    <span className={`${ui.muted} ${styles.rowMeta}`}>{format(strings.SelectedOf, on, tree[c].length)}</span>
                  </>
                )}
                {!closed[c] && nodes.map(renderNode)}
              </React.Fragment>
            );
          })}
        </div>
        <div className={`${ui.side} ${ui.panel}`}>
          <div className={ui.panelTitle}>{strings.SelectionSummary}</div>
          <div className={ui.facts}>
            {side.map(([label, n]) => (
              <React.Fragment key={label}>
                <span className={ui.muted}>{label}</span>
                <span>{n}</span>
              </React.Fragment>
            ))}
          </div>
          <div className={ui.divider} />
          <div className={ui.muted} style={{ fontSize: 13 }}>
            {picked.length ? strings.SideNote : strings.NothingSelected}
          </div>
        </div>
      </div>
    </>
  );
};
