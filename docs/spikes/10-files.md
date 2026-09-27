# Spike 10 – Dokumentumtárak fájljai

Állapot: **lezárva** · Érintett kód (2. fázis): `FileExtractor`, `FileProvider`

## Kérdések

1. **Kiolvasás:**
   - Egy lekérdezésben megkapható-e a fájlok mérete, kivételezése és verziója (`$expand=File`)?
   - Milyen a verzióelőzmény, és letölthető-e egy régi verzió?
2. **Feltöltés:**
   - Működik-e a `Files/AddUsingPath` különleges nevekkel?
   - Mit ad a SharePoint névütközéskor, ha nincs felülírás?
   - Működik-e a darabolt feltöltés (`StartUpload` / `ContinueUpload` / `FinishUpload`) egy 12 MB-os fájllal?
3. **Metaadatok:** a feltöltés után a `ValidateUpdateListItem` beállítja-e egy másik felhasználóra a Létrehozta, Módosította, Létrehozva és Módosítva mezőt, valamint a címet? Keletkezik-e új verzió?
4. **Verzióelőzmény újraépítése:** ha egy fájlt kétszer töltünk fel, és mindkét feltöltés után beállítjuk a metaadatokat, milyen szerző és dátum marad meg a verzióelőzményben?

## A. kísérlet – Forrás site (csak olvas)

**Előkészítés** a Forrás „Dokumentumok” tárában, ha még nincs:
- néhány fájl, köztük egy ékezetes nevű;
- egy mappa, benne egy fájllal;
- egy fájl, amely legalább kétszer volt mentve (hogy legyen verzióelőzménye).

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (url) => { const r = await fetch(url, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status, error: (await r.text()).slice(0, 200) }; };
  const webInfo = await get(guess + '/_api/web?$select=Url,ServerRelativeUrl');
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const hide = (s) => (typeof s === 'string' ? s.split(location.host).join('contoso.sharepoint.com') : s);
  const libUrl = `${rel}/Shared Documents`;
  const LIB = `${web}/_api/web/getList('${enc(libUrl)}')`;
  const out = { library: await get(`${LIB}?$select=Title,ItemCount,EnableVersioning,EnableMinorVersions,MajorVersionLimit,ForceCheckout,EnableModeration`) };
  // Files and folders with the file facts CopyJet needs (size, checkout, version) in one query.
  const rows = await get(`${LIB}/items?$select=Id,FSObjType,FileRef,FileLeafRef,ContentTypeId,AuthorId,EditorId,Created,Modified,File/Length,File/CheckOutType,File/UIVersionLabel,File/MajorVersion,File/MinorVersion&$expand=File&$orderby=ID&$top=200`);
  out.query = rows.httpStatus ? rows : 'ok';
  const items = rows.value || [];
  out.folders = items.filter((i) => i.FSObjType === 1).map((i) => hide(i.FileRef).slice(libUrl.length + 1));
  out.files = items.filter((i) => i.FSObjType === 0).map((i) => ({ id: i.Id, path: hide(i.FileRef).slice(libUrl.length + 1), length: i.File && i.File.Length, checkOut: i.File && i.File.CheckOutType, version: i.File && i.File.UIVersionLabel }));
  // Version history of the file with the most versions: labels, sizes, authors, and whether an old version downloads.
  const multi = items.filter((i) => i.FSObjType === 0 && i.File && i.File.MajorVersion + i.File.MinorVersion > 1).sort((a, b) => (b.File.MajorVersion - a.File.MajorVersion))[0];
  if (multi) {
    const FILE = `${web}/_api/web/GetFileByServerRelativePath(decodedurl='${enc(multi.FileRef)}')`;
    const versions = await get(`${FILE}/Versions?$select=ID,VersionLabel,Size,Created,CheckInComment,Url,IsCurrentVersion,CreatedBy/Email&$expand=CreatedBy`);
    out.versionsOf = hide(multi.FileRef).slice(libUrl.length + 1);
    out.versions = versions.httpStatus ? versions : (versions.value || []).map((v) => ({ id: v.ID, label: v.VersionLabel, size: v.Size, created: v.Created, comment: v.CheckInComment, url: hide(v.Url), byEmailDomain: v.CreatedBy && v.CreatedBy.Email ? v.CreatedBy.Email.split('@')[1] : null }));
    const first = (versions.value || [])[0];
    if (first) {
      const byId = await fetch(`${FILE}/Versions(${first.ID})/$value`, { headers: { Accept: '*/*' } });
      const byUrl = await fetch(`${web}/${first.Url}`, { headers: { Accept: '*/*' } });
      out.oldVersionDownload = { byIdValue: byId.ok ? (await byId.arrayBuffer()).byteLength : `HTTP ${byId.status}`, byUrl: byUrl.ok ? (await byUrl.arrayBuffer()).byteLength : `HTTP ${byUrl.status}`, expectedSize: first.Size };
    }
  }
  // Writable metadata columns of the library (what the item serializer would copy).
  out.columns = ((await get(`${LIB}/fields?$select=InternalName,TypeAsString&$filter=Hidden eq false and ReadOnlyField eq false`)).value || []).map((f) => `${f.InternalName}: ${f.TypeAsString}`);
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## B. kísérlet – Cél site (**ír**, önálló)

Létrehozza a `CopyJetSpike10` dokumentumtárat (verziózva), ha nincs meg. Minden fájlnév időbélyeget kap, így a kísérlet többször is futtatható. A site egy másik felhasználóját használja (`other`), és csak álnevet ír ki.

- **B1:** kis fájl, utána metaadatok (szerző, dátumok, cím).
- **B2:** ugyanaz a név még egyszer, felülírás nélkül.
- **B3:** ékezetes, `#`, `&` és `%` karaktert tartalmazó név.
- **B4:** 12 MB darabolva (5 MB-os darabok), visszaolvasás és összehasonlítás.
- **B5:** két verzió feltöltése, mindkettő után metaadat-beállítással, végül a verzióelőzmény.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const B = { Accept: H.Accept, 'X-RequestDigest': digest };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const libUrl = `${rel}/CopyJetSpike10`;
  const LIB = `${web}/_api/web/getList('${enc(libUrl)}')`;
  const FOLDER = `${web}/_api/web/GetFolderByServerRelativePath(decodedurl='${enc(libUrl)}')`;
  const fileApi = (name) => `${web}/_api/web/GetFileByServerRelativePath(decodedurl='${enc(`${libUrl}/${name}`)}')`;
  const out = { setup: {}, tests: {} };
  // Self-contained: a versioned document library (created when missing).
  if ((await fetch(`${LIB}?$select=Id`, { headers: H })).status === 404) {
    out.setup.libraryCreated = (await fetch(`${web}/_api/web/lists`, { method: 'POST', headers: W, body: JSON.stringify({ Title: 'CopyJetSpike10', BaseTemplate: 101 }) })).status;
  }
  out.setup.versioning = (await fetch(LIB, { method: 'POST', headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ EnableVersioning: true, MajorVersionLimit: 50 }) })).status;
  const me = await (await fetch(`${web}/_api/web/currentuser?$select=Id`, { headers: H })).json();
  const users = ((await (await fetch(`${web}/_api/web/siteusers?$select=Id,LoginName,PrincipalType&$filter=PrincipalType eq 1`, { headers: H })).json()).value || [])
    .filter((u) => /^i:0#\.f\|membership\|/.test(u.LoginName) && !/urn%3aspo%3a|app@sharepoint/i.test(u.LoginName) && u.Id !== me.Id);
  const other = users[0];
  if (!other) { console.log('No other user on the site.'); return; }
  const alias = (id) => (id === me.Id ? 'me' : id === other.Id ? 'other' : id);
  const person = JSON.stringify([{ Key: other.LoginName }]);
  const rs = await (await fetch(`${web}/_api/web/RegionalSettings?$select=LocaleId,Time24`, { headers: H })).json();
  const local = async (iso) => (await (await fetch(`${web}/_api/web/RegionalSettings/TimeZone/utcToLocalTime(@d)?@d='${iso}'`, { headers: H })).json()).value;
  const fmt = (l) => {
    const [d, t] = l.split('T');
    const [y, mo, da] = d.split('-').map(Number);
    const [h, mi] = t.split(':').map(Number);
    return new Intl.DateTimeFormat(rs.LocaleId === 1038 ? 'hu-HU' : 'en-US', { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit', hour12: !rs.Time24, timeZone: 'UTC' })
      .format(new Date(Date.UTC(y, mo - 1, da, h, mi))).replace(/[  ]/g, ' ').replace(/,\s*/, ' ');
  };
  const stamp = Date.now();
  const upload = async (name, body, overwrite) => {
    const r = await fetch(`${FOLDER}/Files/AddUsingPath(DecodedUrl='${enc(name)}',Overwrite=${overwrite ? 'true' : 'false'})?$expand=ListItemAllFields&$select=Name,Length,UIVersionLabel,ListItemAllFields/Id`, { method: 'POST', headers: B, body });
    const text = await r.text();
    let j = {}; try { j = JSON.parse(text); } catch { /* not JSON */ }
    return { status: r.status, id: j.ListItemAllFields && j.ListItemAllFields.Id, name: j.Name, length: j.Length, version: j.UIVersionLabel, error: r.ok ? undefined : text.slice(0, 200) };
  };
  const meta = async (id, created, modified, which) => {
    const values = [{ FieldName: 'Editor', FieldValue: person }, { FieldName: 'Modified', FieldValue: fmt(await local(modified)) }];
    if (which === 'all') values.unshift({ FieldName: 'Author', FieldValue: person }, { FieldName: 'Created', FieldValue: fmt(await local(created)) }, { FieldName: 'Title', FieldValue: 'CopyJet cím' });
    const r = await fetch(`${LIB}/items(${id})/ValidateUpdateListItem`, { method: 'POST', headers: W, body: JSON.stringify({ formValues: values, bNewDocumentUpdate: true }) });
    return r.ok ? ((await r.json()).value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`) : [`HTTP ${r.status}`];
  };
  const read = async (id) => {
    const i = await (await fetch(`${LIB}/items(${id})?$select=AuthorId,EditorId,Created,Modified,Title,OData__UIVersionString`, { headers: H })).json();
    return { AuthorId: alias(i.AuthorId), EditorId: alias(i.EditorId), Created: i.Created, Modified: i.Modified, Title: i.Title, version: i.OData__UIVersionString };
  };
  // B1: small file, then metadata (author, editor, dates, title) with ValidateUpdateListItem.
  const n1 = `kicsi-${stamp}.txt`;
  const u1 = await upload(n1, 'hello CopyJet', false);
  out.tests.B1_upload = u1;
  if (u1.id) {
    out.tests.B1_metaErrors = await meta(u1.id, '2026-01-15T10:00:00Z', '2026-01-16T11:30:00Z', 'all');
    out.tests.B1_after = await read(u1.id);
  }
  // B2: the same name again without overwrite.
  out.tests.B2_conflict = await upload(n1, 'again', false);
  // B3: accented and special characters.
  const n3 = `Ékezetes #1 & 50% ${stamp}.txt`;
  const u3 = await upload(n3, 'árvíztűrő', false);
  out.tests.B3_special = { status: u3.status, nameBack: u3.name === n3, error: u3.error };
  // B4: chunked upload, 12 MB in 5 MB chunks (StartUpload / ContinueUpload / FinishUpload).
  const size = 12 * 1024 * 1024, chunk = 5 * 1024 * 1024;
  const data = new Uint8Array(size);
  for (let i = 0; i < size; i += 4096) data[i] = (i / 4096) % 251;
  data[size - 1] = 7;
  const n4 = `nagy-${stamp}.bin`;
  const empty = await upload(n4, '', false);
  const uploadId = crypto.randomUUID();
  const steps = [];
  let offset = 0;
  for (let start = 0; start < size; start += chunk) {
    const end = Math.min(start + chunk, size);
    const op = start === 0 ? `StartUpload(uploadId=guid'${uploadId}')` : end === size ? `FinishUpload(uploadId=guid'${uploadId}',fileOffset=${offset})` : `ContinueUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`;
    const r = await fetch(`${fileApi(n4)}/${op}`, { method: 'POST', headers: B, body: data.slice(start, end) });
    steps.push(`${op.split('(')[0]}: ${r.status}`);
    offset = end;
  }
  const back = await fetch(`${fileApi(n4)}/$value`, { headers: { Accept: '*/*' } });
  const got = back.ok ? new Uint8Array(await back.arrayBuffer()) : new Uint8Array(0);
  out.tests.B4_chunked = { emptyCreate: empty.status, steps, lengthBack: got.length, same: got.length === size && got[4096 * 3] === 3 && got[size - 1] === 7 };
  // B5: version history rebuilt: v1 by "other" in January, v2 (overwrite) by "other" in February.
  const n5 = `verziok-${stamp}.txt`;
  const v1 = await upload(n5, 'első verzió', false);
  const errs = [];
  if (v1.id) errs.push(...(await meta(v1.id, '2026-01-15T10:00:00Z', '2026-01-15T10:00:00Z', 'all')));
  const v2 = await upload(n5, 'második verzió', true);
  if (v2.id) errs.push(...(await meta(v2.id, '2026-01-15T10:00:00Z', '2026-02-20T09:00:00Z', 'modified')));
  const versions = await (await fetch(`${fileApi(n5)}/Versions?$select=VersionLabel,Created,IsCurrentVersion,CreatedBy/Id&$expand=CreatedBy`, { headers: H })).json();
  out.tests.B5_versions = {
    uploads: [v1.version, v2.version],
    errors: errs,
    current: v2.id ? await read(v2.id) : null,
    history: (versions.value || []).map((v) => ({ label: v.VersionLabel, created: v.Created, by: alias(v.CreatedBy && v.CreatedBy.Id) }))
  };
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- a B1, B3 és B4 sikeres, a `B4_chunked.same` értéke `true`;
- a `B1_after` szerint a szerző `other`, a dátumok pedig a kért értékek;
- a B2 mutatja a névütközés hibakódját;
- a B5-ből kiderül, hogy a verzióelőzmény megőrizheti-e a forrás szerzőit és dátumait.

## Eredmény

### A – kiolvasás (2026-09-27, Forrás site, „Dokumentumok”)

- A tár verziózott (`MajorVersionLimit` 500), mellékverzió, jóváhagyás és kötelező kivétel nélkül.
- ⚠️ `ItemCount: 1`, pedig 3 fájl és 1 mappa van benne: az `ItemCount` **nem megbízható** (spike 08 F-hez hasonlóan).
- ✅ Egy lekérdezés elég: `items?$select=…,File/Length,File/CheckOutType,File/UIVersionLabel,File/MajorVersion,File/MinorVersion&$expand=File`.
  - A `File/Length` **szöveg** (`"422833"`).
  - `CheckOutType: 2` = nincs kivéve (0 = online, 1 = offline kivétel).
- A fájlok: `Kép.png` (413 KB, 1.0), `Ékezetes mappa/README.md` (1.0) és `Ékezetes mappa/Üzleti követelmény … (1).docx` (**10,5 MB**, 3.0). Ékezetes mappa- és fájlnevek.
- ✅ **Verzióelőzmény:** `File/Versions` csak a korábbi verziókat adja (1.0, 2.0; a 3.0 az aktuális fájl). Mezők: `ID` (512, 1024 = major × 512), `VersionLabel`, `Size`, `Created`, `CheckInComment`, `Url` (`_vti_history/512/Shared Documents/…`), `CreatedBy` (a szerző a saját tenantból, `contoso.onmicrosoft.com`).
- ✅ **Régi verzió letöltése** bájtra pontos, mindkét módon: `Versions(512)/$value` és `{web}/_vti_history/512/…`.
- Írható oszlopok: `FileLeafRef` (a név), `Title`, `_ExtendedDescription` (Note), `MediaServiceImageTags` (TaxonomyFieldTypeMulti; a SharePoint automatikus képcímkéi), `ContentType`.

→ **Döntések a kiolvasáshoz:**
- A fájlok és a mappák a lista elemeiből jönnek (`FSObjType`), egy lekérdezéssel.
- A kivett fájl (`CheckOutType` ≠ 2) kimarad, figyelmeztetéssel (rendszerterv §7).
- A metaadat az elemekével azonos szerializálóval megy. A `FileLeafRef` nem metaadat, az a fájl neve.
- A `MediaServiceImageTags` rendszeroszlop: nem másolódik, és figyelmeztetést sem kap.
- Verzióknál a `Versions(ID)/$value` a letöltési mód.

### B – írás (2026-09-27, Cél site, új verziózott tár, a szerző egy másik felhasználó)

- **B1:** ✅ `Files/AddUsingPath(DecodedUrl=…,Overwrite=false)` → 1.0. Utána a `ValidateUpdateListItem` (Author, Editor, Created, Modified, Title, `bNewDocumentUpdate: true`) mindent beállít, és **nem hoz létre új verziót** (1.0 marad). A dokumentumtár itt másképp viselkedik, mint a listaelem mellékletnél (spike 09 B3: 6.0).
- **B2:** névütközés felülírás nélkül: **HTTP 400**, `-2130575257`, „… nevű fájl már létezik”.
- **B3:** ✅ `Ékezetes #1 & 50% ….txt`, a név pontosan visszajön.
- **B4:** ✅ darabolt feltöltés: üres fájl, utána `StartUpload` → `ContinueUpload` → `FinishUpload` (5 MB-os darabok). 12 582 912 bájt, bájtra egyezik.
- **B5:** két feltöltés (a második felülírással), mindkettő után metaadat-beállítással:
  - az aktuális 2.0: `other` a szerző és a módosító, a dátumok pontosak;
  - az előzményben az 1.0 **dátuma** a beállított (2026-01-15), de a **szerzője a telepítő** (`me`).

→ **Döntések a `FileProvider`-hez:**
- 10 MB-ig egyszerű feltöltés, felette darabolt (PnPjs `addChunked`, 10 MB-os darabok).
- **Soha nem ír felül:** a már meglévő útvonal (HTTP 400, `-2130575257`) kimarad, és számolva naplózódik. Így az újrafuttatás a hiányzókat pótolja, duplikáció és felülírás nélkül.
- A metaadat (értékek, `ContentTypeId`, szerző, dátumok) a feltöltés után egy `ValidateUpdateListItem`-mel megy, `bNewDocumentUpdate: true` beállítással, új verzió nélkül.
- **Verziók (opcionális):** a legrégebbitől az aktuálisig sorban feltölt, és minden verzió után beállítja a Módosítottát és a Módosítva értéket.
  - **Ismert korlát:** a régi verziók dátuma megmarad, a szerzőjük a telepítő lesz; csak az aktuális verzióé a forrás szerinti.

Állapot: **lezárva**.

### Valódi telepítés (2026-09-27, 1.4.1.0 és 1.4.2.0)

- Forrás „Dokumentumok” (3 fájl, köztük a 10,5 MB-os `.docx`, és egy ékezetes mappa) és „TesztDoktr” (1 fájl) → Cél: 3 / 3 és 1 / 1, az ékezetes mappával együtt.
- Hiba az 1.4.0.0-ban: az aláhúzással kezdődő mezők (`_ExtendedDescription`) a REST-ben `OData__ExtendedDescription` néven érhetők el, így a nyers név miatt a teljes lekérdezés HTTP 400-at adott. Javítva: `restPropertyOf`.
- A darabolt feltöltés (üres kezdőfájl + `setContentChunked`) **egyetlen** verziót hoz létre, a kezdőfájl nem ad plusz verziót.
- Verzióelőzménnyel (1.4.2.0): a Setup „earlier versions: 2” (31 MB csomag), a Célon a `.docx` 3 verzióval.
- Újrafuttatás: „Files already present and kept”, semmi nem íródott felül.
- A `MediaServiceImageTags` rendszeroszlop az 1.4.2.0-tól a lista-oszlopok közül is kimarad.

