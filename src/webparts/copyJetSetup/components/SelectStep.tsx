import * as React from 'react';
import { Checkbox, IconButton, Stack, Text, Toggle } from '@fluentui/react';
import * as strings from 'CopyJetSetupWebPartStrings';
import type { IDiscoveredArtifact } from '../../../core/model';
import { CATEGORIES, buildTree, format, selectAll, toggle, toggleChild, type Category, type ITreeNode } from './selection';
import styles from './CopyJetSetup.module.scss';

export interface ISelectStepProps {
  artifacts: IDiscoveredArtifact[];
  selected: string[];
  onChange: (selected: string[]) => void;
}

const CATEGORY_LABEL: Record<Category, () => string> = {
  list: () => strings.CategoryList,
  group: () => strings.CategoryGroup,
  siteField: () => strings.CategorySiteField,
  contentType: () => strings.CategoryContentType
};

function unsupportedText(code: string): string {
  const key = `Unsupported${code}` as keyof ICopyJetSetupWebPartStrings;
  return (strings as unknown as Record<string, string>)[key] || strings.UnsupportedOther;
}

const Label: React.FC<{ item: IDiscoveredArtifact }> = ({ item }) => (
  <span>
    {item.title}
    {item.itemCount !== undefined && <span className={styles.meta}> · {format(strings.ItemCount, item.itemCount)}</span>}
    {item.group && <span className={styles.meta}> · {item.group}</span>}
    {item.unsupported && <span className={styles.meta}> · {unsupportedText(item.unsupported)}</span>}
  </span>
);

/** Step 1: tree of the site's artifacts with checkboxes and a "whole site" switch. */
export const SelectStep: React.FC<ISelectStepProps> = ({ artifacts, selected, onChange }) => {
  const tree = React.useMemo(() => buildTree(artifacts), [artifacts]);
  const [open, setOpen] = React.useState<{ [key: string]: boolean }>({});
  const all = React.useMemo(() => selectAll(artifacts), [artifacts]);
  const isOn = (a: IDiscoveredArtifact): boolean => selected.indexOf(a.ref.key) >= 0;
  const wholeSite = all.length > 0 && all.every((k) => selected.indexOf(k) >= 0);

  const renderChildren = (node: ITreeNode, label: string, kind: 'listField' | 'view'): React.ReactNode => {
    const items = node.children.filter((c) => c.ref.kind === kind);
    if (items.length === 0) return null;
    return (
      <div className={styles.children}>
        <Text variant="small" className={styles.meta}>
          {label}
        </Text>
        {items.map((c) => (
          <Checkbox
            key={c.ref.key}
            className={styles.child}
            label={c.title}
            onRenderLabel={() => <Label item={c} />}
            checked={isOn(c)}
            disabled={!!c.unsupported}
            onChange={(_, on) => onChange(toggleChild(selected, node, c, !!on))}
          />
        ))}
      </div>
    );
  };

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Toggle label={strings.WholeSite} inlineLabel checked={wholeSite} onChange={(_, on) => onChange(on ? all : [])} />
      {CATEGORIES.filter((c) => tree[c].length > 0).map((c) => {
        const nodes = tree[c];
        const selectedCount = nodes.filter((n) => isOn(n.item)).length;
        return (
          <div key={c} className={styles.category}>
            <Text variant="mediumPlus" className={styles.categoryTitle}>
              {CATEGORY_LABEL[c]()} <span className={styles.meta}>{format(strings.SelectedOf, selectedCount, nodes.length)}</span>
            </Text>
            {nodes.map((n) => (
              <div key={n.item.ref.key}>
                <Stack horizontal verticalAlign="center">
                  {n.children.length > 0 ? (
                    <IconButton
                      iconProps={{ iconName: open[n.item.ref.key] ? 'ChevronDown' : 'ChevronRight' }}
                      ariaLabel={n.item.title}
                      onClick={() => setOpen({ ...open, [n.item.ref.key]: !open[n.item.ref.key] })}
                    />
                  ) : (
                    <span className={styles.chevronSpace} />
                  )}
                  <Checkbox
                    label={n.item.title}
                    onRenderLabel={() => <Label item={n.item} />}
                    checked={isOn(n.item)}
                    disabled={!!n.item.unsupported}
                    onChange={(_, on) => onChange(toggle(selected, n, !!on))}
                  />
                </Stack>
                {open[n.item.ref.key] && (
                  <>
                    {renderChildren(n, strings.Columns, 'listField')}
                    {renderChildren(n, strings.Views, 'view')}
                  </>
                )}
              </div>
            ))}
          </div>
        );
      })}
    </Stack>
  );
};
