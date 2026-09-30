/**
 * The user mapping table of the MappingStep: two columns, source and target (e-mail, UPN or claims login).
 * Separator ";", "," or tab (whatever the first data line uses); values may be quoted; a first line without
 * any "@" or "|" is taken as a header. Keys are lower case.
 */

export interface IMappingCsv {
  /** Source (lower case) → target. */
  rows: { [source: string]: string };
  /** 1-based line numbers that were not "source;target". */
  invalid: number[];
}

const looksLikeAccount = (value: string): boolean => /[@|]/.test(value);

function splitLine(line: string, separator: string): string[] {
  const out: string[] = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < line.length; i++) {
    const c = line.charAt(i);
    if (c === '"') {
      if (quoted && line.charAt(i + 1) === '"') {
        cell += '"';
        i++;
      } else {
        quoted = !quoted;
      }
    } else if (c === separator && !quoted) {
      out.push(cell);
      cell = '';
    } else {
      cell += c;
    }
  }
  out.push(cell);
  return out.map((v) => v.trim());
}

export function parseMappingCsv(text: string): IMappingCsv {
  const lines = text.replace(/^﻿/, '').split(/\r?\n/);
  const firstData = lines.filter((l) => l.trim())[0] || '';
  const separator = [';', '\t', ','].filter((s) => firstData.indexOf(s) >= 0)[0] || ';';
  const out: IMappingCsv = { rows: {}, invalid: [] };
  let headerChecked = false;
  lines.forEach((line, i) => {
    if (!line.trim()) return;
    const cells = splitLine(line, separator);
    if (!headerChecked) {
      headerChecked = true;
      if (!cells.some(looksLikeAccount)) return; // header
    }
    if (cells.length < 2 || !looksLikeAccount(cells[0]) || !looksLikeAccount(cells[1])) {
      out.invalid.push(i + 1);
      return;
    }
    out.rows[cells[0].toLowerCase()] = cells[1];
  });
  return out;
}

export interface IDomainRule {
  /** Source domain without "@", e.g. forras.hu. */
  from: string;
  to: string;
}

/** The address with its domain replaced by the first matching rule, or undefined. Works on e-mails and claims logins. */
export function replaceDomain(account: string, rules: IDomainRule[]): string | undefined {
  const lower = account.toLowerCase();
  for (const rule of rules) {
    const from = `@${rule.from.replace(/^@/, '').toLowerCase()}`;
    if (from.length > 1 && lower.slice(-from.length) === from) {
      return `${account.slice(0, account.length - from.length)}@${rule.to.replace(/^@/, '')}`;
    }
  }
  return undefined;
}
