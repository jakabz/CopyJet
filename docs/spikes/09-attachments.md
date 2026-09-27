# Spike 09 – Listaelem-mellékletek

Állapot: **lezárva** · Érintett kód (2. fázis): `ItemExtractor`, `ItemProvider`, `packager`

## Kérdések

1. **Kiolvasás:** hogyan találhatók meg a mellékletes elemek a nézetküszöb alatt? Működik-e az `Attachments eq 1` szűrés? Milyen a mellékletek URL-je, és letölthetők-e a `$value`-val?
2. **Írás:** működik-e az `items(id)/AttachmentFiles/add(FileName='…')` ékezetes, szóközös, `#`, `&` és `%` karaktert tartalmazó névvel, és bináris tartalommal?
3. **Mellékhatás:** a melléklet hozzáadása átírja-e a Módosította és a Módosítva mezőt, és keletkezik-e új verzió? Visszaállítható-e mindez egy `ValidateUpdateListItem`-mel (`bNewDocumentUpdate: true`), új verzió nélkül?
4. **Ütközés:** mit ad a SharePoint, ha ugyanazt a nevet másodszor is feltöltjük?

## A. kísérlet – Forrás site (csak olvas)

**Előkészítés:** a Forrás „Teszt lista” egy-két eleméhez csatolj mellékletet, lehetőleg egy ékezetes nevű fájlt is.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (url) => { const r = await fetch(url, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status }; };
  const webInfo = await get(guess + '/_api/web?$select=Url,ServerRelativeUrl');
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const hide = (s) => (typeof s === 'string' ? s.split(location.host).join('contoso.sharepoint.com') : s);
  const LIST = `${web}/_api/web/getList('${enc(rel + '/Lists/Teszt lista')}')`;
  const info = await get(`${LIST}?$select=Title,EnableAttachments,ItemCount`);
  const out = { list: info, items: [] };
  // Items that have attachments: the Attachments flag, then their files.
  const items = (await get(`${LIST}/items?$select=Id,Title,Attachments&$filter=Attachments eq 1&$top=50`)).value || [];
  out.filterWorked = Array.isArray(items);
  for (const i of items) {
    const files = (await get(`${LIST}/items(${i.Id})/AttachmentFiles?$select=FileName,ServerRelativeUrl`)).value || [];
    const detail = [];
    for (const f of files) {
      const meta = await get(`${web}/_api/web/GetFileByServerRelativePath(decodedurl='${enc(f.ServerRelativeUrl)}')?$select=Length,TimeLastModified`);
      const blob = await fetch(`${web}/_api/web/GetFileByServerRelativePath(decodedurl='${enc(f.ServerRelativeUrl)}')/$value`, { headers: { Accept: '*/*' } });
      detail.push({ fileName: f.FileName, serverRelativeUrl: hide(f.ServerRelativeUrl), length: meta.Length, downloaded: blob.ok ? (await blob.arrayBuffer()).byteLength : `HTTP ${blob.status}` });
    }
    out.items.push({ id: i.Id, title: i.Title, attachments: detail });
  }
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## B. kísérlet – Cél site (**ír**, önálló)

Létrehozza a `CopyJetSpike09` listát (verziózott, mellékletek bekapcsolva), ha nincs meg. Utána egy elemet ír ugyanúgy, ahogy a CopyJet: a site egy másik felhasználója a szerző, régi dátumokkal. Ehhez négy mellékletet ad (egyszerű, ékezetes, különleges karakteres, bináris), feltölti az egyik nevet még egyszer, végül visszaállítja a Módosítottát és a Módosítva értéket.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const listUrl = `${rel}/Lists/CopyJetSpike09`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  const out = { setup: {}, tests: {} };
  // Self-contained: a versioned test list with attachments on (created when missing).
  if ((await fetch(`${LIST}?$select=Id`, { headers: H })).status === 404) {
    out.setup.listCreated = (await fetch(`${web}/_api/web/lists`, { method: 'POST', headers: W, body: JSON.stringify({ Title: 'CopyJetSpike09', BaseTemplate: 100 }) })).status;
  }
  out.setup.listMerge = (await fetch(LIST, { method: 'POST', headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ EnableAttachments: true, EnableVersioning: true, MajorVersionLimit: 50 }) })).status;
  const me = await (await fetch(`${web}/_api/web/currentuser?$select=Id,LoginName`, { headers: H })).json();
  const users = ((await (await fetch(`${web}/_api/web/siteusers?$select=Id,LoginName,PrincipalType&$filter=PrincipalType eq 1`, { headers: H })).json()).value || [])
    .filter((u) => /^i:0#\.f\|membership\|/.test(u.LoginName) && !/urn%3aspo%3a|app@sharepoint/i.test(u.LoginName) && u.Id !== me.Id);
  const other = users[0] || me;
  out.setup.author = users[0] ? 'other' : 'me (no other user on the site)';
  const alias = (id) => (id === me.Id ? 'me' : id === other.Id ? 'other' : id);
  const rs = await (await fetch(`${web}/_api/web/RegionalSettings?$select=LocaleId,Time24`, { headers: H })).json();
  const local = async (iso) => (await (await fetch(`${web}/_api/web/RegionalSettings/TimeZone/utcToLocalTime(@d)?@d='${iso}'`, { headers: H })).json()).value;
  const fmt = (l) => {
    const [d, t] = l.split('T');
    const [y, mo, da] = d.split('-').map(Number);
    const [h, mi] = t.split(':').map(Number);
    return new Intl.DateTimeFormat(rs.LocaleId === 1038 ? 'hu-HU' : 'en-US', { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit', hour12: !rs.Time24, timeZone: 'UTC' })
      .format(new Date(Date.UTC(y, mo - 1, da, h, mi))).replace(/[  ]/g, ' ').replace(/,\s*/, ' ');
  };
  const person = JSON.stringify([{ Key: other.LoginName }]);
  const created = fmt(await local('2026-01-15T10:00:00Z')), modified = fmt(await local('2026-01-16T11:30:00Z'));
  const read = async (id) => {
    const i = await (await fetch(`${LIST}/items(${id})?$select=AuthorId,EditorId,Created,Modified,Attachments,OData__UIVersionString`, { headers: H })).json();
    const files = ((await (await fetch(`${LIST}/items(${id})/AttachmentFiles?$select=FileName`, { headers: H })).json()).value || []).map((f) => f.FileName);
    return { AuthorId: alias(i.AuthorId), EditorId: alias(i.EditorId), Created: i.Created, Modified: i.Modified, version: i.OData__UIVersionString, files };
  };
  // 1. The item as CopyJet writes it (spike 08: author and dates kept).
  const r = await fetch(`${LIST}/AddValidateUpdateItemUsingPath`, { method: 'POST', headers: W, body: JSON.stringify({
    listItemCreateInfo: { FolderPath: { DecodedUrl: listUrl }, UnderlyingObjectType: 0 },
    formValues: [{ FieldName: 'Title', FieldValue: 'Spike 09' }, { FieldName: 'Author', FieldValue: person }, { FieldName: 'Editor', FieldValue: person }, { FieldName: 'Created', FieldValue: created }, { FieldName: 'Modified', FieldValue: modified }],
    bNewDocumentUpdate: true }) });
  const id = Number((((await r.json()).value || []).filter((v) => v.FieldName === 'Id')[0] || {}).FieldValue);
  out.tests.B0_created = await read(id);
  // 2. Attachments with plain, accented and special names; binary content (a few PNG header bytes).
  const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4, 5]);
  const add = async (name, body) => {
    const res = await fetch(`${LIST}/items(${id})/AttachmentFiles/add(FileName='${enc(name)}')`, { method: 'POST', headers: { Accept: H.Accept, 'X-RequestDigest': digest }, body });
    return res.ok ? 'ok' : `HTTP ${res.status}: ${(await res.text()).slice(0, 160)}`;
  };
  out.tests.B1_add = {
    plain: await add('egyszeru.txt', 'hello'),
    accented: await add('Ékezetes név (1).txt', 'árvíztűrő'),
    hash: await add('hash#jel & 50%.txt', 'x'),
    binary: await add('kep.png', bytes)
  };
  out.tests.B1_after = await read(id);
  // Binary round trip.
  const bin = await fetch(`${web}/_api/web/GetFileByServerRelativePath(decodedurl='${enc(`${listUrl}/Attachments/${id}/kep.png`)}')/$value`, { headers: { Accept: '*/*' } });
  out.tests.B1_binaryBack = bin.ok ? Array.from(new Uint8Array(await bin.arrayBuffer())).join(',') === Array.from(bytes).join(',') : `HTTP ${bin.status}`;
  // 3. The same name again.
  out.tests.B2_duplicate = await add('egyszeru.txt', 'again');
  // 4. Restore Editor and Modified afterwards (bNewDocumentUpdate true: no new version expected).
  const u = await fetch(`${LIST}/items(${id})/ValidateUpdateListItem`, { method: 'POST', headers: W, body: JSON.stringify({
    formValues: [{ FieldName: 'Editor', FieldValue: person }, { FieldName: 'Modified', FieldValue: modified }], bNewDocumentUpdate: true }) });
  out.tests.B3_restore = { status: u.status, errors: u.ok ? ((await u.json()).value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`) : [], after: await read(id) };
  out.expected = { AuthorId: out.setup.author.slice(0, 5) === 'other' ? 'other' : 'me', Modified: '2026-01-16T11:30:00Z' };
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- `B1_add` mind `ok`, és `B1_binaryBack: true`.
- A `B1_after` alapján látszik, hogy a mellékletek elrontják-e a Módosítottát, a dátumot és a verziót.
- A `B3_restore.after` legyen újra `EditorId: other` és `2026-01-16T11:30:00Z`, lehetőleg változatlan verziószámmal.

## Eredmény

### A – kiolvasás (2026-09-25, Forrás site)

- ✅ `Attachments eq 1` szűrés működik (`filterWorked: true`); a mellékletes elem: 7 („Adele teszt”).
- ✅ `items(id)/AttachmentFiles` → `FileName`, `ServerRelativeUrl` = `/sites/Forras/Lists/Teszt lista/Attachments/7/logo.png`.
- ✅ Letöltés `GetFileByServerRelativePath(decodedurl=…)/$value`-val bájtra pontos (1712 = 1712).
- A `Length` **szövegként** jön (`"1712"`), számmá kell alakítani.

→ Kiolvasás: az elemlapok `Attachments` jelzője alapján csak a mellékletes elemeknél kell lekérni az `AttachmentFiles`-t; a fájlok `attachments/<listkey>/<forrás elem-ID>/<fájlnév>` útvonalon kerülnek a csomagba.

### B – írás (2026-09-27, Cél site, verziózott lista, a szerző egy másik felhasználó)

- **B0:** az elem a spike 08 szerint jön létre (`AuthorId`/`EditorId` = `other`, régi dátumok, 1.0).
- **B1:** ✅ mind a négy melléklet feltöltődött `AttachmentFiles/add(FileName='…')`-szel: `egyszeru.txt`, `Ékezetes név (1).txt`, `hash#jel & 50%.txt` és a bináris `kep.png` (visszaolvasva bájtra egyezik).
  - ⚠️ **Mellékhatás:** minden melléklet új verziót hoz létre (1.0 → 5.0), a Módosította a telepítő lesz (`me`), a Módosítva az aktuális idő. A Létrehozta és a Létrehozva változatlan.
- **B2:** ugyanaz a név másodszor: **HTTP 400**, `-2130575257`, „A megadott név már használatban van.”
- **B3:** ✅ `ValidateUpdateListItem` (Editor, Modified, `bNewDocumentUpdate: true`) visszaállítja a Módosítottát és a Módosítva értéket, de ez is új verzió (6.0).

→ **Döntés:**
- A mellékletek az elemek után, elemenként sorban kerülnek fel.
- Utána egy `ValidateUpdateListItem` visszaállítja a Módosítottát és a Módosítva értéket.
- A már meglévő nevet (HTTP 400) kihagyja.
- **Ismert korlát:** egy N mellékletes elem verzióelőzményében N+1 plusz verzió jelenik meg a telepítő nevén. A lista nézetében a Módosította és a Módosítva a forrás szerinti. Verzió nélküli felülírás CSOM `UpdateOverwriteVersion`-nel lehetséges, ez egy későbbi fejlesztés.

Állapot: **lezárva**.
