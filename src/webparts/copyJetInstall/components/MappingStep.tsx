import * as React from 'react';
import * as strings from 'CopyJetInstallWebPartStrings';
import type { ICopyJetTemplate } from '../../../core/model';
import { Message, Tag, ui } from '../../../shared/components/ui';

/**
 * Step 2 (docs/ui install-2): user and term mapping. Mapping itself arrives in phase 2; this step lists what
 * the template references so nothing is skipped silently.
 */
export const MappingStep: React.FC<{ template: ICopyJetTemplate }> = ({ template }) => {
  const users = template.principals;
  const terms = template.terms || [];
  if (users.length === 0 && terms.length === 0) {
    return <Message kind="success">{strings.NothingToMap}</Message>;
  }
  return (
    <>
      <Message kind="warning">{strings.MappingLater}</Message>
      {users.length > 0 && (
        <>
          <div className={ui.strong}>{strings.Users}</div>
          <table className={ui.table}>
            <thead>
              <tr>
                <th>{strings.ColName}</th>
                <th>{strings.ColSourceUser}</th>
                <th>{strings.ColTargetUser}</th>
                <th>{strings.ColStatus}</th>
              </tr>
            </thead>
            <tbody>
              {users.map((u) => (
                <tr key={u.key}>
                  <td className={ui.strong}>{u.displayName || u.key}</td>
                  <td className={ui.muted}>{u.email || u.loginName || '—'}</td>
                  <td className={ui.muted}>—</td>
                  <td>
                    <Tag kind="diff">{strings.NotMapped}</Tag>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}
      {terms.length > 0 && (
        <>
          <div className={ui.strong}>{strings.ManagedMetadata}</div>
          <table className={ui.table}>
            <thead>
              <tr>
                <th>{strings.ColTermPath}</th>
                <th>{strings.ColStatus}</th>
              </tr>
            </thead>
            <tbody>
              {terms.map((t) => (
                <tr key={t.key}>
                  <td>{t.path}</td>
                  <td>
                    <Tag kind="diff">{strings.NotMapped}</Tag>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}
    </>
  );
};
