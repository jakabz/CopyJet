import * as React from 'react';
import { Spinner } from '@fluentui/react';
import * as strings from 'CopyJetInstallWebPartStrings';
import { Logger } from '../../../core/logger';
import type { IPrincipalMapping, PrincipalMapper } from '../../../core/mapping';
import type { ICopyJetTemplate } from '../../../core/model';
import { TokenContext } from '../../../core/tokenizer';
import { Message, Tag, ui } from '../../../shared/components/ui';
import { format } from './labels';

export interface IMappingStepProps {
  template: ICopyJetTemplate;
  /** Looks the people up on the target site; the install reuses its results. */
  principals: PrincipalMapper;
}

const loginText = (login: string | undefined): string => (login ? login.split('|').pop() || login : '—');

/**
 * Step 2 (docs/ui install-2): the template's people looked up on the target site (same account, then same
 * e-mail). Domain replacement, CSV and a fallback user, and term mapping come later in phase 2.
 */
export const MappingStep: React.FC<IMappingStepProps> = ({ template, principals }) => {
  const users = template.principals;
  const terms = template.terms || [];
  const [mappings, setMappings] = React.useState<{ [key: string]: IPrincipalMapping } | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    if (users.length === 0) return;
    const ac = new AbortController();
    principals.map(users.map((u) => u.key), new TokenContext(), new Logger(), ac.signal).then(
      (list) => {
        if (ac.signal.aborted) return;
        const byKey: { [key: string]: IPrincipalMapping } = {};
        list.forEach((m) => (byKey[m.key] = m));
        setMappings(byKey);
      },
      (e: unknown) => !ac.signal.aborted && setError(e instanceof Error ? e.message : String(e))
    );
    return () => ac.abort();
  }, [template, principals]);

  if (users.length === 0 && terms.length === 0) {
    return <Message kind="success">{strings.NothingToMap}</Message>;
  }
  const notFound = mappings ? users.filter((u) => !mappings[u.key] || !mappings[u.key].login).length : 0;
  return (
    <>
      {users.length > 0 && (
        <>
          <Message kind="note">{strings.MappingAuto}</Message>
          {error && <Message kind="error">{error}</Message>}
          {mappings && notFound > 0 && <Message kind="warning">{format(strings.MappingNotFoundNote, notFound)}</Message>}
          <div className={ui.strong}>{strings.Users}</div>
          {!mappings && !error ? (
            <Spinner label={strings.MappingChecking} />
          ) : (
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
                {users.map((u) => {
                  const m = mappings && mappings[u.key];
                  return (
                    <tr key={u.key}>
                      <td className={ui.strong}>{u.displayName || u.key}</td>
                      <td className={ui.muted}>{u.email || loginText(u.loginName)}</td>
                      <td className={ui.muted}>{loginText(m && m.login)}</td>
                      <td>
                        {m && m.strategy === 'sameLogin' && <Tag kind="same">{strings.MapSameLogin}</Tag>}
                        {m && m.strategy === 'sameEmail' && <Tag kind="new">{strings.MapSameEmail}</Tag>}
                        {(!m || !m.login) && <Tag kind="diff">{strings.MapNotFound}</Tag>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </>
      )}
      {terms.length > 0 && (
        <>
          <Message kind="warning">{strings.TermsLater}</Message>
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
