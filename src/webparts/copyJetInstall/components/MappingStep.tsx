import * as React from 'react';
import { Spinner } from '@fluentui/react';
import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import * as strings from 'CopyJetInstallWebPartStrings';
import { Logger } from '../../../core/logger';
import { parseMappingCsv, type IDomainRule, type IPrincipalMapping, type IPrincipalRules, type PrincipalMapper, type PrincipalStrategy } from '../../../core/mapping';
import type { ICopyJetTemplate } from '../../../core/model';
import { TokenContext } from '../../../core/tokenizer';
import { Button, Message, Tag, ui } from '../../../shared/components/ui';
import { format } from './labels';

export interface IMappingStepProps {
  sp: SPFI;
  template: ICopyJetTemplate;
  /** Looks the people up on the target site; its rules and results are reused by the install. */
  principals: PrincipalMapper;
}

const loginText = (login: string | undefined): string => (login ? login.split('|').pop() || login : '—');

const domainOf = (email: string | undefined): string | undefined => (email && email.indexOf('@') > 0 ? email.split('@')[1].toLowerCase() : undefined);

const METHOD: Record<PrincipalStrategy, () => string> = {
  manual: () => strings.MethodManual,
  csv: () => strings.MethodCsv,
  sameLogin: () => strings.MapSameLogin,
  sameEmail: () => strings.MapSameEmail,
  domain: () => strings.MethodDomain,
  fallback: () => strings.MethodFallback
};

/** The most frequent e-mail domain of the template's people (the domain replacement's suggested source). */
function commonDomain(template: ICopyJetTemplate): string | undefined {
  const counts: { [d: string]: number } = {};
  template.principals.forEach((p) => {
    const d = domainOf(p.email);
    if (d) counts[d] = (counts[d] || 0) + 1;
  });
  return Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
}

/**
 * Step 2 (docs/ui install-2): the template's people looked up on the target site. Order: a target typed in the
 * table, the CSV table, the same account, the same e-mail, domain replacement, the fallback user. Changing a
 * rule maps everyone again. Managed Metadata terms follow in a later step.
 */
export const MappingStep: React.FC<IMappingStepProps> = ({ sp, template, principals }) => {
  const users = template.principals;
  const terms = template.terms || [];
  const [rules, setRules] = React.useState<IPrincipalRules>(principals.rules);
  const [mappings, setMappings] = React.useState<{ [key: string]: IPrincipalMapping } | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [domainOpen, setDomainOpen] = React.useState(!!(rules.domains && rules.domains.length));
  const [domain, setDomain] = React.useState<IDomainRule>((rules.domains && rules.domains[0]) || { from: commonDomain(template) || '', to: '' });
  const [csvNote, setCsvNote] = React.useState<string | undefined>(undefined);
  const [manual, setManual] = React.useState<{ [key: string]: string }>(rules.manual || {});
  const [fallback, setFallback] = React.useState<string>(rules.fallback || '');
  const [useFallback, setUseFallback] = React.useState(!!rules.fallback);
  const csvInput = React.useRef<HTMLInputElement>(null);

  // Suggest the installing user's domain as the target of the domain replacement.
  React.useEffect(() => {
    if (domain.to) return;
    sp.web.currentUser.select('Email')<{ Email?: string }>().then(
      (me) => {
        const d = domainOf(me.Email);
        if (d && d !== domain.from) setDomain((x) => (x.to ? x : { ...x, to: d }));
      },
      () => undefined
    );
  }, [sp]);

  React.useEffect(() => {
    if (users.length === 0) return;
    const ac = new AbortController();
    setMappings(undefined);
    setError(undefined);
    principals.setRules(rules);
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
  }, [template, principals, rules]);

  // A new rules object maps everyone again, so only a real change is applied.
  const apply = (patch: Partial<IPrincipalRules>): void =>
    setRules((r) => {
      const next = { ...r, ...patch };
      return JSON.stringify(next) === JSON.stringify(r) ? r : next;
    });

  const loadCsv = (file: File | undefined): void => {
    if (!file) return;
    file.text().then(
      (text) => {
        const csv = parseMappingCsv(text);
        const count = Object.keys(csv.rows).length;
        setCsvNote(format(strings.CsvLoaded, count) + (csv.invalid.length ? format(strings.CsvInvalid, csv.invalid.join(', ')) : ''));
        apply({ csv: csv.rows });
      },
      (e: unknown) => setCsvNote(e instanceof Error ? e.message : String(e))
    );
  };

  if (users.length === 0 && terms.length === 0) {
    return <Message kind="success">{strings.NothingToMap}</Message>;
  }
  const found = mappings ? users.filter((u) => mappings[u.key] && mappings[u.key].login).length : 0;
  const notFound = mappings ? users.length - found : 0;
  return (
    <>
      {users.length > 0 && (
        <>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
            <span className={ui.strong}>{strings.Users}</span>
            {mappings && <span className={ui.muted}>{format(strings.UsersMapped, found, users.length)}</span>}
            <span style={{ flexGrow: 1 }} />
            <Button text={strings.DomainReplace} onClick={() => setDomainOpen(!domainOpen)} />
            <Button text={strings.CsvImport} onClick={() => csvInput.current && csvInput.current.click()} />
            <input
              ref={csvInput}
              type="file"
              accept=".csv,.txt,text/csv"
              style={{ display: 'none' }}
              onChange={(e) => {
                loadCsv(e.target.files ? e.target.files[0] : undefined);
                e.target.value = '';
              }}
            />
          </div>
          {domainOpen && (
            <div className={ui.box} style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', padding: 12 }}>
              <label htmlFor="cj-dom-from">{strings.DomainFrom}</label>
              <input id="cj-dom-from" className={ui.input} style={{ maxWidth: 200 }} value={domain.from} onChange={(e) => setDomain({ ...domain, from: e.target.value.trim() })} />
              <span aria-hidden="true">→</span>
              <label htmlFor="cj-dom-to">{strings.DomainTo}</label>
              <input id="cj-dom-to" className={ui.input} style={{ maxWidth: 200 }} value={domain.to} onChange={(e) => setDomain({ ...domain, to: e.target.value.trim() })} />
              <Button text={strings.Apply} kind="primary" disabled={!domain.from || !domain.to} onClick={() => apply({ domains: [domain] })} />
              {rules.domains && rules.domains.length > 0 && <Button text={strings.Clear} onClick={() => apply({ domains: [] })} />}
            </div>
          )}
          {csvNote && <Message kind="note">{csvNote}</Message>}
          {error && <Message kind="error">{error}</Message>}
          {mappings && notFound > 0 && <Message kind="warning">{format(strings.MappingNotFoundNote, notFound)}</Message>}
          {!mappings && !error ? (
            <Spinner label={strings.MappingChecking} />
          ) : (
            <table className={ui.table}>
              <thead>
                <tr>
                  <th>{strings.ColName}</th>
                  <th>{strings.ColSourceUser}</th>
                  <th>{strings.ColTargetUser}</th>
                  <th>{strings.ColMethod}</th>
                  <th>{strings.ColStatus}</th>
                </tr>
              </thead>
              <tbody>
                {users.map((u) => {
                  const m = mappings && mappings[u.key];
                  const typed = manual[u.key] || '';
                  const showInput = !m || !m.login || m.strategy === 'fallback' || m.strategy === 'manual';
                  return (
                    <tr key={u.key}>
                      <td className={ui.strong}>{u.displayName || u.key}</td>
                      <td className={ui.muted}>{u.email || loginText(u.loginName)}</td>
                      <td className={ui.muted}>
                        {showInput ? (
                          <input
                            className={ui.input}
                            aria-label={format(strings.ManualTargetFor, u.displayName || u.key)}
                            placeholder={m && m.login ? loginText(m.login) : strings.ManualPlaceholder}
                            value={typed}
                            onChange={(e) => setManual({ ...manual, [u.key]: e.target.value })}
                            onBlur={() => apply({ manual: { ...manual } })}
                            onKeyDown={(e) => e.key === 'Enter' && apply({ manual: { ...manual } })}
                          />
                        ) : (
                          loginText(m && m.login)
                        )}
                      </td>
                      <td className={ui.muted}>{m && m.strategy ? METHOD[m.strategy]() : '—'}</td>
                      <td>{m && m.login ? <Tag kind={m.strategy === 'fallback' ? 'info' : 'new'}>{strings.StatusFound}</Tag> : <Tag kind="diff">{strings.MapNotFound}</Tag>}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
          <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap', fontSize: 14 }}>
            <span>{strings.FallbackLabel}</span>
            <span>
              <input
                type="radio"
                className={ui.check}
                name="cj-fallback"
                id="cj-fb-empty"
                checked={!useFallback}
                onChange={() => {
                  setUseFallback(false);
                  apply({ fallback: undefined });
                }}
              />{' '}
              <label htmlFor="cj-fb-empty">{strings.FallbackEmpty}</label>
            </span>
            <span>
              <input type="radio" className={ui.check} name="cj-fallback" id="cj-fb-user" checked={useFallback} onChange={() => setUseFallback(true)} />{' '}
              <label htmlFor="cj-fb-user">{strings.FallbackUser}</label>
            </span>
            {useFallback && (
              <>
                <input
                  className={ui.input}
                  style={{ maxWidth: 260 }}
                  aria-label={strings.FallbackUser}
                  placeholder={strings.ManualPlaceholder}
                  value={fallback}
                  onChange={(e) => setFallback(e.target.value.trim())}
                  onKeyDown={(e) => e.key === 'Enter' && apply({ fallback: fallback || undefined })}
                />
                <Button text={strings.Apply} disabled={!fallback} onClick={() => apply({ fallback })} />
              </>
            )}
          </div>
          <Message kind="note">{strings.MappingAuto}</Message>
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
