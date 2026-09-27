/**
 * Values for AddValidateUpdateItemUsingPath / ValidateUpdateListItem. That API parses numbers and dates in the
 * TARGET web's regional settings and time zone (spike 08): en-US wants "12.5" and "3/2/2026 9:15 AM", hu-HU
 * wants "12,5" and "2026. 03. 02. 9:15"; ISO strings are rejected everywhere. These pure functions build
 * those strings; the local time itself comes from SharePoint (RegionalSettings/TimeZone/utcToLocalTime).
 */

export interface IWebLocale {
  /** SharePoint LocaleId (LCID). */
  localeId: number;
  /** BCP 47 tag for Intl, derived from the LCID. */
  tag: string;
  decimalSeparator: string;
  time24: boolean;
}

/** LCIDs SharePoint Online offers as regional settings, mapped to BCP 47 tags (most used ones). */
const LCID_TAGS: { [lcid: number]: string } = {
  1025: 'ar-SA', 1026: 'bg-BG', 1027: 'ca-ES', 1028: 'zh-TW', 1029: 'cs-CZ', 1030: 'da-DK', 1031: 'de-DE',
  1032: 'el-GR', 1033: 'en-US', 1035: 'fi-FI', 1036: 'fr-FR', 1037: 'he-IL', 1038: 'hu-HU', 1040: 'it-IT',
  1041: 'ja-JP', 1042: 'ko-KR', 1043: 'nl-NL', 1044: 'nb-NO', 1045: 'pl-PL', 1046: 'pt-BR', 1048: 'ro-RO',
  1049: 'ru-RU', 1050: 'hr-HR', 1051: 'sk-SK', 1053: 'sv-SE', 1054: 'th-TH', 1055: 'tr-TR', 1058: 'uk-UA',
  1060: 'sl-SI', 1061: 'et-EE', 1062: 'lv-LV', 1063: 'lt-LT', 1066: 'vi-VN', 2052: 'zh-CN', 2057: 'en-GB',
  2070: 'pt-PT', 2074: 'sr-Latn-RS', 3079: 'de-AT', 3081: 'en-AU', 3082: 'es-ES', 3084: 'fr-CA', 4105: 'en-CA',
  2055: 'de-CH', 4108: 'fr-CH', 1081: 'hi-IN', 16393: 'en-IN'
};

export function lcidToTag(lcid: number): string | undefined {
  return LCID_TAGS[lcid];
}

/** "2026-03-02T09:15:00" (local wall-clock time as returned by utcToLocalTime) → parts. */
function localParts(localIso: string): { y: number; mo: number; d: number; h: number; mi: number } {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(localIso);
  if (!m) throw new Error(`Not a local ISO date-time: ${localIso}`);
  return { y: Number(m[1]), mo: Number(m[2]), d: Number(m[3]), h: Number(m[4]), mi: Number(m[5]) };
}

/**
 * Local wall-clock time → the web's short date + short time, e.g. "3/2/2026 9:15 AM" (en-US),
 * "2026. 03. 02. 9:15" (hu-HU). Intl formats in UTC so the browser's own time zone never shifts the value.
 * Intl puts ", " between date and time in some locales; SharePoint expects a single space.
 */
export function formatSpDateTime(localIso: string, locale: IWebLocale): string {
  const p = localParts(localIso);
  const date = new Date(Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi));
  const text = new Intl.DateTimeFormat(locale.tag, {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    hour12: !locale.time24,
    timeZone: 'UTC'
  }).format(date);
  return text.replace(/ | /g, ' ').replace(/,\s*/, ' ').replace(/\s+/g, ' ').trim();
}

/** Plain number in the web's notation, no grouping: 12.5 → "12,5" when the decimal separator is ",". */
export function formatSpNumber(value: number, locale: IWebLocale): string {
  return String(value).replace('.', locale.decimalSeparator);
}
