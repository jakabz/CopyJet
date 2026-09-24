/**
 * Template key from a name, matching the schema's key pattern ([A-Za-z0-9_.-]{1,128}):
 * accents dropped (not the letters), whitespace → '_', other characters removed.
 * 'Shared Documents' → 'Shared_Documents', 'Ügyfél projektek' → 'Ugyfel_projektek'.
 */
export function toTemplateKey(name: string, fallback: string): string {
  const key = name
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_.-]/g, '')
    .slice(0, 120);
  return key || fallback;
}

/**
 * Keys for [id, name] pairs, unique case-insensitively (a clash gets _2, _3 …), returned by id.
 * Order-independent: pairs are processed sorted by id.
 */
export function uniqueKeys(entries: Array<[string, string]>, fallback: string): { [id: string]: string } {
  const out: { [id: string]: string } = {};
  const used: { [key: string]: boolean } = {};
  entries
    .slice()
    .sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0))
    .forEach(([id, name]) => {
      const base = toTemplateKey(name, fallback);
      let key = base;
      for (let n = 2; used[key.toLowerCase()]; n++) key = `${base}_${n}`;
      used[key.toLowerCase()] = true;
      out[id] = key;
    });
  return out;
}
