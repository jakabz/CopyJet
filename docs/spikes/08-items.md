# Spike 08 – Listaelemek: kiolvasott értékek, írás `AddValidateUpdateItemUsingPath`-szal

Állapot: **lezárva** · Érintett kód (2. fázis): `FieldValueSerializer`, `ItemExtractor`, `ItemProvider`

## Kérdések

1. **Kiolvasás:** milyen alakban adja vissza a REST az egyes mezőtípusok értékét? Érintett típusok: Person, Choice, MultiChoice, Lookup, Number, DateTime, URL, Boolean és a rendszermezők (Author, Editor, Created, Modified). A mentett érték legyen nyelvfüggetlen.
2. **Írás:** az `AddValidateUpdateItemUsingPath` szöveges `FieldValue`-kat vár. Ezeket a **cél web területi beállításai** szerint értelmezi (tizedesjel, dátumformátum, időzóna).
   - Elfogadja-e a számot ponttal (`12.5`) és vesszővel (`12,5`)?
   - Elfogadja-e a dátumot ISO-alakban (`2026-03-02T08:15:00Z`)? Ha nem, milyen helyi alakban?
3. **Rendszermezők:** beállítható-e ugyanebben a hívásban a Létrehozta, a Módosította, a Létrehozva és a Módosítva mező?
   - A Létrehozta és a Módosította felhasználó: `[{"Key":"i:0#.f|membership|…"}]`.
   - A `bNewDocumentUpdate: true` kell-e hozzá?
4. **Person, Lookup, MultiChoice, URL, Boolean:** melyik írási alakot fogadja el?

## A. kísérlet – Forrás site (csak olvas)

**Előkészítés:** a Forrás site „Teszt lista” listájában legyen legalább 2 elem, lehetőleg minden oszlop kitöltésével: Valaki, Válassz, Lookup, SiteColumn1, SiteColumn2 (tizedes számmal, pl. 12,5), Szöveg, Keresés.

F12 → Console → beillesztés → Enter, a kiírt JSON-t küldd vissza. A kód a felhasználókat csak álnévvel (`user1`, `user2` …) mutatja, e-mail-címet nem ír ki.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const LIST = `${web}/_api/web/getList('${enc(rel + '/Lists/Teszt lista')}')`;
  const rs = await (await fetch(`${web}/_api/web/RegionalSettings?$select=LocaleId,DecimalSeparator,DateSeparator,Time24&$expand=TimeZone`, { headers: H })).json();
  const fields = ((await (await fetch(`${LIST}/fields?$select=InternalName,TypeAsString&$filter=Hidden eq false and ReadOnlyField eq false`, { headers: H })).json()).value || [])
    .filter((f) => !/^(ContentType|Attachments|_ColorTag|ComplianceAssetId)$/.test(f.InternalName));
  const names = fields.map((f) => f.InternalName);
  const userFields = fields.filter((f) => /^User/.test(f.TypeAsString)).map((f) => f.InternalName).concat(['Author', 'Editor']);
  const lookupFields = fields.filter((f) => /^Lookup/.test(f.TypeAsString)).map((f) => f.InternalName);
  const select = ['ID', 'Created', 'Modified'].concat(names.filter((n) => userFields.indexOf(n) < 0 && lookupFields.indexOf(n) < 0))
    .concat(userFields.map((n) => `${n}/Title,${n}/EMail,${n}/Name,${n}Id`)).concat(lookupFields.map((n) => `${n}/Title,${n}Id`));
  const items = ((await (await fetch(`${LIST}/items?$top=3&$orderby=ID&$select=${select.join(',')}&$expand=${userFields.concat(lookupFields).join(',')}`, { headers: H })).json()).value || []);
  const alias = {}; let n = 0;
  const anon = (v) => JSON.parse(JSON.stringify(v, (k, x) => {
    if ((k === 'EMail' || k === 'Name' || k === 'Title') && typeof x === 'string' && /@|\|/.test(x)) return alias[x] || (alias[x] = 'user' + (++n));
    return x;
  }));
  // The same items as the list form shows them (formatted in the web's locale).
  const asText = ((await (await fetch(`${LIST}/items?$top=3&$orderby=ID&$select=ID,FieldValuesAsText&$expand=FieldValuesAsText`, { headers: H })).json()).value || [])
    .map((i) => { const t = i.FieldValuesAsText || {}; const o = {}; names.concat(['Created', 'Modified']).forEach((k) => { if (t[k] !== undefined) o[k] = t[k]; }); return o; });
  const out = { web, regional: { LocaleId: rs.LocaleId, DecimalSeparator: rs.DecimalSeparator, DateSeparator: rs.DateSeparator, Time24: rs.Time24, TimeZone: rs.TimeZone && rs.TimeZone.Description }, fields: fields.map((f) => `${f.InternalName}: ${f.TypeAsString}`), items: anon(items), asText: anon(asText) };
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## B. kísérlet – Cél site (**ír**: új tesztlista, oszlopok, 3 elem)

Csak a Cél **teszt** site-on futtasd. Semmit nem töröl, és nem függ korábbi tesztelemektől. Létrehozza a `CopyJetSpike08` listát 9 oszloppal: Szám, Dátum, Választás, Többes választás, Személy, Lookup (önmagára), Igen/Nem, Hivatkozás, Többsoros szöveg. Utána 3 elemet ír `AddValidateUpdateItemUsingPath`-szal:

- **1. elem, ISO-értékek:** a szám ponttal (`12.5`), a dátum `2026-03-02T08:15:00Z`, a Létrehozva és a Módosítva ISO-alakban.
- **2. elem, helyi értékek:** a szám a web tizedesjelével, a dátum a web formátumában, a Módosította a jelenlegi felhasználó.
- **3. elem, lookup:** az 1. elemre mutató lookup, és a Létrehozta, Módosította, Létrehozva, Módosítva együtt, `bNewDocumentUpdate: true`-val.

Végül visszaolvassa mindhárom elemet. A felhasználót itt is csak álnévvel mutatja. A teszt után a lista kézzel törölhető.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const short = async (r) => { try { return (await r.text()).slice(0, 400); } catch { return ''; } };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const listUrl = `${rel}/Lists/CopyJetSpike08`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  const out = { web };

  const rs = await (await fetch(`${web}/_api/web/RegionalSettings?$select=LocaleId,DecimalSeparator,DateSeparator,Time24`, { headers: H })).json();
  out.regional = rs;
  const me = await (await fetch(`${web}/_api/web/currentuser?$select=Id,LoginName`, { headers: H })).json();

  let list = await fetch(`${LIST}?$select=Id`, { headers: H });
  if (list.status === 404) {
    await fetch(`${web}/_api/web/lists`, { method: 'POST', headers: W, body: JSON.stringify({ Title: 'CopyJetSpike08', BaseTemplate: 100 }) });
    list = await fetch(`${LIST}?$select=Id`, { headers: H });
    const listId = (await list.json()).Id;
    const xml = [
      '<Field Name="CJNum" DisplayName="CJ szám" Type="Number" Decimals="2" />',
      '<Field Name="CJDate" DisplayName="CJ dátum" Type="DateTime" Format="DateTime" />',
      '<Field Name="CJChoice" DisplayName="CJ választás" Type="Choice"><CHOICES><CHOICE>Egy</CHOICE><CHOICE>Kettő</CHOICE></CHOICES></Field>',
      '<Field Name="CJMulti" DisplayName="CJ többes" Type="MultiChoice"><CHOICES><CHOICE>A</CHOICE><CHOICE>B</CHOICE><CHOICE>C</CHOICE></CHOICES></Field>',
      '<Field Name="CJUser" DisplayName="CJ személy" Type="User" UserSelectionMode="PeopleOnly" />',
      `<Field Name="CJLookup" DisplayName="CJ lookup" Type="Lookup" List="{${listId}}" ShowField="Title" />`,
      '<Field Name="CJBool" DisplayName="CJ igen/nem" Type="Boolean"><Default>0</Default></Field>',
      '<Field Name="CJUrl" DisplayName="CJ hivatkozás" Type="URL" Format="Hyperlink" />',
      '<Field Name="CJNote" DisplayName="CJ jegyzet" Type="Note" NumLines="4" RichText="FALSE" />'
    ];
    for (const x of xml) {
      const r = await fetch(`${LIST}/fields/createfieldasxml`, { method: 'POST', headers: W, body: JSON.stringify({ parameters: { SchemaXml: x, Options: 12 } }) });
      if (!r.ok) out['fieldError_' + x.slice(13, 22)] = await short(r);
    }
  }

  const add = async (key, values, newDocument) => {
    const r = await fetch(`${LIST}/AddValidateUpdateItemUsingPath`, { method: 'POST', headers: W, body: JSON.stringify({
      listItemCreateInfo: { FolderPath: { DecodedUrl: listUrl }, UnderlyingObjectType: 0 },
      formValues: Object.keys(values).map((k) => ({ FieldName: k, FieldValue: values[k] })),
      bNewDocumentUpdate: !!newDocument }) });
    const body = r.ok ? await r.json() : await short(r);
    out[key] = { status: r.status, fieldErrors: r.ok ? (body.value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`) : body,
      id: r.ok ? ((body.value || []).filter((v) => v.FieldName === 'Id')[0] || {}).FieldValue : undefined };
    return out[key].id;
  };
  const dec = rs.DecimalSeparator || ',';
  const user = JSON.stringify([{ Key: me.LoginName }]);

  const id1 = await add('item1_iso', { Title: 'ISO', CJNum: '12.5', CJDate: '2026-03-02T08:15:00Z', CJChoice: 'Egy', CJMulti: ';#A;#C;#', CJUser: user, CJBool: '1', CJUrl: 'https://example.com, Példa', CJNote: 'Első sor\nMásodik sor', Created: '2026-01-15T10:00:00Z', Modified: '2026-01-16T11:30:00Z' });
  await add('item2_local', { Title: 'HELYI', CJNum: `12${dec}5`, CJDate: new Date(Date.UTC(2026, 2, 2, 8, 15)).toLocaleString(), CJChoice: 'Kettő', CJBool: 'true', Editor: user });
  await add('item3_lookup_system', { Title: 'LOOKUP', CJLookup: String(id1 || 1), Author: user, Editor: user, Created: '2026-01-15T10:00:00Z', Modified: '2026-01-16T11:30:00Z' }, true);

  const sel = 'ID,Title,CJNum,CJDate,CJChoice,CJMulti,CJBool,CJUrl,CJNote,CJLookupId,CJUserId,AuthorId,EditorId,Created,Modified';
  const items = ((await (await fetch(`${LIST}/items?$select=${sel},FieldValuesAsText/CJNum,FieldValuesAsText/CJDate&$expand=FieldValuesAsText&$orderby=ID`, { headers: H })).json()).value || []);
  out.readBack = items.map((i) => ({ ...i, CJUserId: i.CJUserId === me.Id ? 'me' : i.CJUserId, AuthorId: i.AuthorId === me.Id ? 'me' : i.AuthorId, EditorId: i.EditorId === me.Id ? 'me' : i.EditorId,
    asText: i.FieldValuesAsText ? { CJNum: i.FieldValuesAsText.CJNum, CJDate: i.FieldValuesAsText.CJDate } : undefined, FieldValuesAsText: undefined }));
  out.localDateSent = new Date(Date.UTC(2026, 2, 2, 8, 15)).toLocaleString();

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## Mit várunk?

- **A:** az értékek nyelvfüggetlen alakja.
  - A dátumok UTC ISO-alakúak.
  - A szám a JSON-ban szám.
  - A Person mezőnél ID van, plusz a kibontott e-mail és login.
  - A Lookupnál ID van.
  - A MultiChoice tömb.
  - A `FieldValuesAsText` a helyi formázás, csak összehasonlításhoz.
- **B:**
  - Melyik írási alak adta a helyes értéket: `CJNum` = 12.5, `CJDate` = 2026-03-02T08:15:00Z.
  - A `Created`/`Modified` a kért érték lett-e, és az `AuthorId`/`EditorId` „me”-e.
  - A `fieldErrors` mező üres-e.

## Eredmény

### A – kiolvasás (2026-09-24, Forrás site)

- **Területi beállítás:** en-US (1033), tizedesjel `.`, dátumelválasztó `/`, 12 órás idő, időzóna „(UTC+01:00) Budapest”. A magyar felhasználói felület a site területi beállításait nem változtatja meg.
- **A REST-értékek nyelvfüggetlenek** (`odata=nometadata`):
  - **Number:** JSON-szám (`SiteColumn2: 11.5`).
  - **DateTime** (Created, Modified): UTC ISO (`2026-09-24T15:06:38Z`). A felület ugyanezt „9/24/2026 5:06 PM” alakban mutatja (`FieldValuesAsText`).
  - **Choice:** szöveg (`"Kettő"`).
  - **Lookup:** `<mező>Id` = a cél elem ID-ja (`LookupId: 3`); a kibontott `Title` csak tájékoztató.
  - **User** (Valaki, Author, Editor): `<mező>Id` = site-szintű felhasználó-ID. Ez site-onként más, ezért a sablonba a kibontott `EMail` és `Name` (login) kerül, `{principal:key}` tokenként.
- A `FieldValuesAsText` a kódolt belső nevű mezőknél (`V_x00e1_lassz`) más kulcsot használ, összehasonlításra nem kell.

→ **Kiolvasás:** a sablon a REST-alakot tárolja: számot, ISO dátumot (UTC), szöveget, `{ lookup: [ID] }`, `{ principals: ["{principal:key}"] }`, `{ url, description }`.

### B – írás (2026-09-24, Cél site, en-US 1033, tizedesjel `.`)

- **ISO dátum: ❌.** A `CJDate`, a `Created` és a `Modified` mezőre is ezt a hibát adta: „Adjon meg egy érvényes dátumot a(z) 1/1/1900 és 12/31/8900 közötti tartományban”. Ez a `2026-03-02T08:15:00Z` alakra jött.
- **Böngészős helyi alak** (`3/2/2026, 9:15:00 AM`): ❌. A hibaüzenet megadja az elvárt mintát: `2/23/2012 2:25 PM`.
- **Egyetlen hibás mező miatt az egész elem nem jön létre** (`id: 0`, HTTP 200, `HasException`). A provider ezért a mezőhibákat is hibának veszi, nem csak a HTTP-státuszt.
- A többi mezőre nem jött hiba: a szám ponttal (`12.5`), a Choice, a MultiChoice (`;#A;#C;#`), a User (`[{"Key":login}]`), a Boolean (`1` / `true`), az URL (`url, leírás`) és a Note. Mivel egyik elem sem jött létre, ezeket visszaolvasással még nem igazoltuk.

## C. kísérlet – Cél site (**ír**): dátumalakok, időzóna, típusok visszaolvasása

Csak a Cél **teszt** site-on futtasd, a B-ben létrehozott `CopyJetSpike08` listán. Új elemeket ír bele, semmit nem töröl.

- **C1:** egy elem minden dátum nélküli típussal, visszaolvasásra.
- **C2:** ugyanaz az időpont négy alakban: `2026-03-02 09:15`, `2026-03-02T09:15:00`, `3/2/2026 9:15 AM`, ISO `Z`-vel.
- **C3:** a rendszermezők (Létrehozta, Módosította, Létrehozva, Módosítva) az első elfogadott alakkal, `bNewDocumentUpdate` `false` és `true` értékkel.
- **Időzóna:** a `utcToLocalTime` REST-függvény téli és nyári időpontra.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const listUrl = `${rel}/Lists/CopyJetSpike08`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  const me = await (await fetch(`${web}/_api/web/currentuser?$select=Id,LoginName`, { headers: H })).json();
  const rs = await (await fetch(`${web}/_api/web/RegionalSettings?$select=LocaleId,DecimalSeparator&$expand=TimeZone`, { headers: H })).json();
  const out = { web, locale: rs.LocaleId, decimal: rs.DecimalSeparator, timeZone: rs.TimeZone && rs.TimeZone.Description, tzInfo: rs.TimeZone && rs.TimeZone.Information };

  // UTC -> web local time, computed by SharePoint (handles daylight saving).
  const toLocal = async (iso) => {
    const r = await fetch(`${web}/_api/web/RegionalSettings/TimeZone/utcToLocalTime(@date)?@date='${iso}'`, { headers: H });
    const b = r.ok ? await r.json() : await r.text();
    return r.ok ? b.value : `HTTP ${r.status}: ${String(b).slice(0, 200)}`;
  };
  out.utcToLocal_winter = await toLocal('2026-03-02T08:15:00Z');
  out.utcToLocal_summer = await toLocal('2026-07-02T08:15:00Z');

  const add = async (title, values, newDocument) => {
    const r = await fetch(`${LIST}/AddValidateUpdateItemUsingPath`, { method: 'POST', headers: W, body: JSON.stringify({
      listItemCreateInfo: { FolderPath: { DecodedUrl: listUrl }, UnderlyingObjectType: 0 },
      formValues: [{ FieldName: 'Title', FieldValue: title }].concat(Object.keys(values).map((k) => ({ FieldName: k, FieldValue: values[k] }))),
      bNewDocumentUpdate: !!newDocument }) });
    const body = r.ok ? await r.json() : await r.text();
    const errors = r.ok ? (body.value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`) : [String(body).slice(0, 300)];
    return { title, errors, id: r.ok ? Number(((body.value || []).filter((v) => v.FieldName === 'Id')[0] || {}).FieldValue) : 0 };
  };
  const user = JSON.stringify([{ Key: me.LoginName }]);

  // C1: every non-date type, to read back
  const c1 = await add('C1-tipusok', { CJNum: '12.5', CJChoice: 'Egy', CJMulti: ';#A;#C;#', CJUser: user, CJBool: '1', CJUrl: 'https://example.com, Példa', CJNote: 'Első sor\nMásodik sor' });
  out.c1 = c1;
  // C2: date formats (the intended instant is 2026-03-02 08:15 UTC = 09:15 in Budapest)
  out.c2 = [];
  for (const f of ['2026-03-02 09:15', '2026-03-02T09:15:00', '3/2/2026 9:15 AM', '2026-03-02T08:15:00Z']) {
    out.c2.push({ sent: f, ...(await add('C2 ' + f, { CJDate: f })) });
  }
  // C3: system fields with the first accepted format, bNewDocumentUpdate false and true
  const ok = out.c2.filter((x) => x.errors.length === 0)[0];
  out.c3 = [];
  if (ok) {
    const created = ok.sent, modified = ok.sent.replace('09:15', '10:30').replace('9:15', '10:30');
    for (const nd of [false, true]) {
      out.c3.push({ bNewDocumentUpdate: nd, ...(await add('C3 nd=' + nd, { CJLookup: String(c1.id || 1), Author: user, Editor: user, Created: created, Modified: modified }, nd)) });
    }
  }

  const sel = 'ID,Title,CJNum,CJDate,CJChoice,CJMulti,CJBool,CJUrl,CJNote,CJLookupId,CJUserId,AuthorId,EditorId,Created,Modified';
  const items = ((await (await fetch(`${LIST}/items?$select=${sel}&$orderby=ID`, { headers: H })).json()).value || []);
  const mark = (v) => (v === me.Id ? 'me' : v);
  out.readBack = items.map((i) => ({ ...i, CJUserId: mark(i.CJUserId), AuthorId: mark(i.AuthorId), EditorId: mark(i.EditorId) }));

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();

```

**Ha van rá időd, második futás magyar területi beállítással:** a Cél site-on *Webhelybeállítások → Területi beállítások → Területi beállítás: Magyar*, majd futtasd újra ugyanezt a kódot. Utána visszaállíthatod angolra. Ebből derül ki, hogy a nyelvfüggetlen alakok (C2 első kettő) más területi beállításnál is működnek-e.

**Mit várunk:**
- **C1:** a visszaolvasott `CJNum` = 12.5, a `CJMulti` = `["A","C"]`, a `CJUserId` = „me”, a `CJBool` = `true`, a `CJUrl` = `{ Url, Description }`.
- **C2:** melyik alak ad hibátlan elemet, és a `CJDate` a visszaolvasáskor `2026-03-02T08:15:00Z` lesz-e. Ez azt jelenti, hogy a SharePoint a site időzónája (Budapest, UTC+1) szerint értelmezte az időt.
- **C3:** a `Created`/`Modified` a kért érték lett-e, és melyik `bNewDocumentUpdate` beállítással.
- **`utcToLocalTime`:** télen `09:15`, nyáron `10:15` (nyári időszámítás).

### C – eredmény, 1. futás (2026-09-24, Cél site, en-US 1033, időzóna Budapest)

- **C1 – dátum nélküli típusok: ✅**, írva és visszaolvasva:

  | Mező | Küldött `FieldValue` | Visszaolvasva |
  | --- | --- | --- |
  | Number | `12.5` | `12.5` |
  | Choice | `Egy` | `"Egy"` |
  | MultiChoice | `;#A;#C;#` | `["A","C"]` |
  | User | `[{"Key":"<login>"}]` | `CJUserId` = a felhasználó |
  | Boolean | `1` | `true` |
  | URL | `https://example.com, Példa` | `{ Url, Description }` |
  | Note | `Első sor\nMásodik sor` | ugyanaz |
  | Lookup (C3) | `1` | `CJLookupId: 1` |

- **C2 – dátumalakok:**
  - ✅ **csak a site mintája működik:** `3/2/2026 9:15 AM` → `2026-03-02T08:15:00Z`, tehát a site időzónája szerint értelmezte (Budapest, UTC+1).
  - ❌ `2026-03-02 09:15` (a hibaüzenet a `2/23/2012 2:25 PM` mintát kéri), ❌ `2026-03-02T09:15:00`, ❌ `2026-03-02T08:15:00Z`.
- **Időzóna:** a `RegionalSettings/TimeZone/utcToLocalTime(@date)` REST-függvény a nyári időszámítást is kezeli: `08:15Z` → télen `09:15`, nyáron `10:15`. A `TimeZone.Information` csak eltolásokat ad (Bias −60, DaylightBias −60), átállási dátumokat nem, ezért saját számításra nem alkalmas.
- **C3 – rendszermezők: ✅.** A `Created` és a `Modified` pontosan a kért érték lett (`08:15Z`, `09:30Z`), `bNewDocumentUpdate` `false` és `true` mellett is. Az `Author`/`Editor` beállítása hiba nélkül lefutott, de a saját felhasználóval, így egy *másik* felhasználóra ez még nincs igazolva (felhasználó-leképezés).

→ **Részeredmény:** a dátumot `utcToLocalTime`-mal a site idejére kell váltani, majd a site területi mintája szerint szövegként átadni. A minta nyelvfüggő, ezt a magyar futás pontosítja.

### C – eredmény, 2. futás (2026-09-24, Cél site, hu-HU 1038, tizedesjel `,`)

- **Szám:** a `12.5` ❌ („Itt csak számok szerepelhetnek”), az elem nem jött létre. Magyar site-on tizedesvessző kell.
- **Dátum:** mind a négy alak ❌. Az elvárt minta `2012. 02. 23. 14:25` (év. hónap. nap. óra:perc, 24 órás).
- Az `utcToLocalTime` ugyanazt adta, mint angolul (09:15 / 10:15).

→ **Döntés:** a számot és a dátumot a **cél web** területi beállítása szerint kell szövegként átadni (`src/core/items/locale.ts`):
- `RegionalSettings`: `LocaleId` → BCP 47 nyelvkód, `DecimalSeparator`, `Time24`;
- a helyi idő a SharePoint `utcToLocalTime` függvényéből jön;
- a formázás `Intl.DateTimeFormat` a site nyelvén, UTC-ben (hogy a böngésző időzónája ne tolja el), a dátum és az idő közti vesszőt szóközre cserélve.

Helyben kipróbálva: en-US → `2/23/2012 2:25 PM`, hu-HU → `2012. 02. 23. 14:25`. Mindkettő pontosan a hibaüzenetben kért minta.

## D. kísérlet – Cél site (**ír**): a CopyJet-formázó valódi site-on

Csak a Cél **teszt** site-on futtasd, a `CopyJetSpike08` listán. **Pontosan azt a formázót** futtatja (a `locale.ts` lefordítva), amelyet a provider használ, a cél web aktuális területi beállításával. Két elemet ír: dátum, Létrehozva, Módosítva, valamint egy tizedes és egy negatív szám. A második elemben év- és nyári időszámítás-határ is van. Végül visszaolvassa és összeveti őket (`matches`).

Érdemes kétszer futtatni: magyar és angol területi beállítással.

```js
(async () => {
  // Exactly the formatter CopyJet uses (compiled from src/core/items/locale.ts):
  const L = (function () {
    const exports = {};
    /**
     * Values for AddValidateUpdateItemUsingPath / ValidateUpdateListItem. That API parses numbers and dates in the
     * TARGET web's regional settings and time zone (spike 08): en-US wants "12.5" and "3/2/2026 9:15 AM", hu-HU
     * wants "12,5" and "2026. 03. 02. 9:15"; ISO strings are rejected everywhere. These pure functions build
     * those strings; the local time itself comes from SharePoint (RegionalSettings/TimeZone/utcToLocalTime).
     */
    exports.lcidToTag = lcidToTag;
    exports.formatSpDateTime = formatSpDateTime;
    exports.formatSpNumber = formatSpNumber;
    /** LCIDs SharePoint Online offers as regional settings, mapped to BCP 47 tags (most used ones). */
    const LCID_TAGS = {
        1025: 'ar-SA', 1026: 'bg-BG', 1027: 'ca-ES', 1028: 'zh-TW', 1029: 'cs-CZ', 1030: 'da-DK', 1031: 'de-DE',
        1032: 'el-GR', 1033: 'en-US', 1035: 'fi-FI', 1036: 'fr-FR', 1037: 'he-IL', 1038: 'hu-HU', 1040: 'it-IT',
        1041: 'ja-JP', 1042: 'ko-KR', 1043: 'nl-NL', 1044: 'nb-NO', 1045: 'pl-PL', 1046: 'pt-BR', 1048: 'ro-RO',
        1049: 'ru-RU', 1050: 'hr-HR', 1051: 'sk-SK', 1053: 'sv-SE', 1054: 'th-TH', 1055: 'tr-TR', 1058: 'uk-UA',
        1060: 'sl-SI', 1061: 'et-EE', 1062: 'lv-LV', 1063: 'lt-LT', 1066: 'vi-VN', 2052: 'zh-CN', 2057: 'en-GB',
        2070: 'pt-PT', 2074: 'sr-Latn-RS', 3079: 'de-AT', 3081: 'en-AU', 3082: 'es-ES', 3084: 'fr-CA', 4105: 'en-CA',
        2055: 'de-CH', 4108: 'fr-CH', 1081: 'hi-IN', 16393: 'en-IN'
    };
    function lcidToTag(lcid) {
        return LCID_TAGS[lcid];
    }
    /** "2026-03-02T09:15:00" (local wall-clock time as returned by utcToLocalTime) → parts. */
    function localParts(localIso) {
        const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(localIso);
        if (!m)
            throw new Error(`Not a local ISO date-time: ${localIso}`);
        return { y: Number(m[1]), mo: Number(m[2]), d: Number(m[3]), h: Number(m[4]), mi: Number(m[5]) };
    }
    /**
     * Local wall-clock time → the web's short date + short time, e.g. "3/2/2026 9:15 AM" (en-US),
     * "2026. 03. 02. 9:15" (hu-HU). Intl formats in UTC so the browser's own time zone never shifts the value.
     * Intl puts ", " between date and time in some locales; SharePoint expects a single space.
     */
    function formatSpDateTime(localIso, locale) {
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
    function formatSpNumber(value, locale) {
        return String(value).replace('.', locale.decimalSeparator);
    }
    return exports;
  })();
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const listUrl = `${rel}/Lists/CopyJetSpike08`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  const rs = await (await fetch(`${web}/_api/web/RegionalSettings?$select=LocaleId,DecimalSeparator,Time24`, { headers: H })).json();
  const locale = { localeId: rs.LocaleId, tag: L.lcidToTag(rs.LocaleId), decimalSeparator: rs.DecimalSeparator, time24: rs.Time24 };
  const toLocal = async (iso) => (await (await fetch(`${web}/_api/web/RegionalSettings/TimeZone/utcToLocalTime(@date)?@date='${iso}'`, { headers: H })).json()).value;
  const cases = [
    { utc: '2026-03-02T08:15:00Z', created: '2026-01-15T10:00:00Z', modified: '2026-07-16T11:30:00Z', num: 12.5 },
    { utc: '2026-12-31T23:05:00Z', created: '2025-06-01T06:07:00Z', modified: '2025-06-01T16:07:00Z', num: -1234.75 }
  ];
  const out = { web, locale, results: [] };
  for (const c of cases) {
    const values = {
      CJDate: L.formatSpDateTime(await toLocal(c.utc), locale),
      Created: L.formatSpDateTime(await toLocal(c.created), locale),
      Modified: L.formatSpDateTime(await toLocal(c.modified), locale),
      CJNum: L.formatSpNumber(c.num, locale)
    };
    const r = await fetch(`${LIST}/AddValidateUpdateItemUsingPath`, { method: 'POST', headers: W, body: JSON.stringify({
      listItemCreateInfo: { FolderPath: { DecodedUrl: listUrl }, UnderlyingObjectType: 0 },
      formValues: [{ FieldName: 'Title', FieldValue: 'D ' + locale.tag }].concat(Object.keys(values).map((k) => ({ FieldName: k, FieldValue: values[k] }))),
      bNewDocumentUpdate: true }) });
    const body = r.ok ? await r.json() : await r.text();
    const id = r.ok ? Number(((body.value || []).filter((v) => v.FieldName === 'Id')[0] || {}).FieldValue) : 0;
    const errors = r.ok ? (body.value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`) : [String(body).slice(0, 300)];
    let back = null;
    if (id) {
      const i = await (await fetch(`${LIST}/items(${id})?$select=CJDate,CJNum,Created,Modified`, { headers: H })).json();
      back = { CJDate: i.CJDate, CJNum: i.CJNum, Created: i.Created, Modified: i.Modified };
    }
    const expected = { CJDate: c.utc, CJNum: c.num, Created: c.created, Modified: c.modified };
    out.results.push({ sent: values, errors, back, matches: back ? Object.keys(expected).every((k) => back[k] === expected[k]) : false });
  }
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:** mindkét elemnél `errors: []` és `matches: true`, mindkét területi beállítással.

### D – eredmény, 1. futás (2026-09-25, Cél site, hu-HU 1038, tizedesjel `,`, 24 órás)

**✅ Mindkét elem létrejött, `errors: []`, `matches: true`.**

| Mező | Küldött (hu-HU) | Visszaolvasva (UTC) |
| --- | --- | --- |
| CJDate | `2026. 03. 02. 9:15` | `2026-03-02T08:15:00Z` |
| Created | `2026. 01. 15. 11:00` | `2026-01-15T10:00:00Z` |
| Modified (nyári idő) | `2026. 07. 16. 13:30` | `2026-07-16T11:30:00Z` |
| CJNum | `12,5` | `12.5` |
| CJDate (évhatár) | `2027. 01. 01. 0:05` | `2026-12-31T23:05:00Z` |
| Created (nyári idő) | `2025. 06. 01. 8:07` | `2025-06-01T06:07:00Z` |
| Modified | `2025. 06. 01. 18:07` | `2025-06-01T16:07:00Z` |
| CJNum (negatív) | `-1234,75` | `-1234.75` |

- Az óra vezető nulla nélkül (`9:15`, `0:05`) is elfogadott.
- Az évhatár és a nyári időszámítás is helyes: a helyi idő a SharePointtól jön, nem a böngészőtől.

### D – eredmény, 2. futás (2026-09-25, Cél site, en-US 1033, tizedesjel `.`, 12 órás)

**✅ Mindkét elem létrejött, `errors: []`, `matches: true`.** Küldött értékek: `3/2/2026 9:15 AM`, `1/15/2026 11:00 AM`, `7/16/2026 1:30 PM`, `12.5`; illetve `1/1/2027 12:05 AM` (évhatár), `6/1/2025 8:07 AM`, `6/1/2025 6:07 PM`, `-1234.75`. A visszaolvasott UTC-értékek pontosan egyeznek a magyar futáséval.

→ **Lezárva:** a `locale.ts` formázó mindkét területi beállítással helyes. A helyi időt a SharePoint `utcToLocalTime` adja (nyári időszámítás, évhatár), a formázás a cél web `LocaleId`/`DecimalSeparator`/`Time24` alapján készül.

## E. kísérlet – Cél site (**ír**): többértékű mezők, tartalomtípus, más szerző, mappa

Csak a Cél **teszt** site-on futtasd, a `CopyJetSpike08` listán. Semmit nem töröl.

A lista előkészítése:
- két új oszlop: `CJLookupMulti` (több lookup) és `CJUserMulti` (több személy);
- bekapcsolja a tartalomtípusokat és a mappákat;
- hozzáadja a beépített Hirdetmény (Announcement, `0x0104`) tartalomtípust.

Utána elemeket ír, és visszaolvassa őket:
- **E1a–c:** a LookupMulti három írási alakja (`a;#;#b`, `a;#b`, `a;#;#b;#;#`).
- **E2:** UserMulti, `[{"Key":…},{"Key":…}]`.
- **E3a–b:** URL, benne vesszővel (a vessző duplázva és anélkül), a leírásban is vesszővel.
- **E4a–b:** a tartalomtípus név (`ContentType`) vagy ID (`ContentTypeId`) szerint.
- **E5:** a Létrehozta és a Módosította egy *másik* felhasználó (ha van ilyen a site-on).
- **E6:** mappa létrehozása ugyanezzel a hívással, elem a mappában, valamint egy nem létező almappa.

A felhasználókat itt is csak álnévvel (`me`, `other`) mutatja.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const short = async (r) => { try { return (await r.text()).slice(0, 300); } catch { return ''; } };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const listUrl = `${rel}/Lists/CopyJetSpike08`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  const out = { setup: {} };

  // Setup: two multi-value columns, content types on, the built-in Announcement type (0x0104) added.
  const listId = (await (await fetch(`${LIST}?$select=Id`, { headers: H })).json()).Id;
  const addField = async (name, xml) => {
    const exists = await fetch(`${LIST}/fields/getbyinternalnameortitle('${name}')?$select=Id`, { headers: H });
    if (exists.ok) return 'exists';
    const r = await fetch(`${LIST}/fields/createfieldasxml`, { method: 'POST', headers: W, body: JSON.stringify({ parameters: { SchemaXml: xml, Options: 12 } }) });
    return r.ok ? 'created' : await short(r);
  };
  out.setup.CJLookupMulti = await addField('CJLookupMulti', `<Field Name="CJLookupMulti" DisplayName="CJ lookup multi" Type="LookupMulti" Mult="TRUE" List="{${listId}}" ShowField="Title" />`);
  out.setup.CJUserMulti = await addField('CJUserMulti', '<Field Name="CJUserMulti" DisplayName="CJ személy multi" Type="UserMulti" Mult="TRUE" UserSelectionMode="PeopleOnly" />');
  const merge = await fetch(LIST, { method: 'POST', headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ ContentTypesEnabled: true, EnableFolderCreation: true }) });
  out.setup.listMerge = merge.status;
  const cts = (await (await fetch(`${LIST}/contenttypes?$select=StringId,Name`, { headers: H })).json()).value || [];
  if (!cts.some((c) => c.StringId.indexOf('0x0104') === 0)) {
    const r = await fetch(`${LIST}/contenttypes/addAvailableContentType('0x0104')`, { method: 'POST', headers: W });
    out.setup.addCt = r.ok ? 'added' : await short(r);
  }
  const listCts = ((await (await fetch(`${LIST}/contenttypes?$select=StringId,Name`, { headers: H })).json()).value || []);
  out.setup.listContentTypes = listCts.map((c) => `${c.Name}: ${c.StringId.slice(0, 12)}…`);
  const ann = listCts.filter((c) => c.StringId.indexOf('0x0104') === 0)[0];

  // People: me and one other real user of the site (only aliases are printed).
  const me = await (await fetch(`${web}/_api/web/currentuser?$select=Id,LoginName`, { headers: H })).json();
  const users = ((await (await fetch(`${web}/_api/web/siteusers?$select=Id,LoginName,PrincipalType&$filter=PrincipalType eq 1`, { headers: H })).json()).value || [])
    .filter((u) => /^i:0#\.f\|membership\|/.test(u.LoginName) && !/urn%3aspo%3a|app@sharepoint/i.test(u.LoginName) && u.Id !== me.Id);
  const other = users[0];
  out.setup.otherUser = other ? 'found' : 'none (Author/UserMulti tests use only you)';
  const alias = (id) => (id === me.Id ? 'me' : other && id === other.Id ? 'other' : id);

  // Two lookup targets.
  const ids = ((await (await fetch(`${LIST}/items?$select=ID&$top=2&$orderby=ID`, { headers: H })).json()).value || []).map((i) => i.ID);
  const [a, b] = ids;
  out.setup.lookupTargets = ids.length;

  const add = async (title, values, opts) => {
    const o = opts || {};
    const r = await fetch(`${LIST}/AddValidateUpdateItemUsingPath`, { method: 'POST', headers: W, body: JSON.stringify({
      listItemCreateInfo: { FolderPath: { DecodedUrl: o.folder || listUrl }, UnderlyingObjectType: o.folderItem ? 1 : 0, LeafName: o.leaf ? { DecodedUrl: o.leaf } : undefined },
      formValues: [{ FieldName: 'Title', FieldValue: title }].concat(Object.keys(values).map((k) => ({ FieldName: k, FieldValue: values[k] }))),
      bNewDocumentUpdate: !!o.newDoc }) });
    const body = r.ok ? await r.json() : await short(r);
    const errors = r.ok ? (body.value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`) : [body];
    const id = r.ok ? Number(((body.value || []).filter((v) => v.FieldName === 'Id')[0] || {}).FieldValue) : 0;
    return { title, sent: values, status: r.status, errors, id };
  };
  const users2 = JSON.stringify([{ Key: me.LoginName }].concat(other ? [{ Key: other.LoginName }] : []));
  const tests = [];
  tests.push(await add('E1a lookupmulti', { CJLookupMulti: `${a};#;#${b}` }));
  tests.push(await add('E1b lookupmulti', { CJLookupMulti: `${a};#${b}` }));
  tests.push(await add('E1c lookupmulti', { CJLookupMulti: `${a};#;#${b};#;#` }));
  tests.push(await add('E2 usermulti', { CJUserMulti: users2 }));
  tests.push(await add('E3a url comma', { CJUrl: 'https://example.com/a,,b?x=1,,2, Leírás, vesszővel' }));
  tests.push(await add('E3b url plain', { CJUrl: 'https://example.com/a,b, Leírás' }));
  if (ann) {
    tests.push(await add('E4a ct by name', { ContentType: ann.Name }));
    tests.push(await add('E4b ct by id', { ContentTypeId: ann.StringId }));
  }
  if (other) tests.push(await add('E5 author other', { Author: JSON.stringify([{ Key: other.LoginName }]), Editor: JSON.stringify([{ Key: other.LoginName }]) }, { newDoc: true }));
  const folder = await add('E6folder', {}, { folderItem: true, leaf: 'E6folder' });
  tests.push(folder);
  tests.push(await add('E6 item in folder', {}, { folder: `${listUrl}/E6folder` }));
  tests.push(await add('E6b nested missing', {}, { folder: `${listUrl}/E6folder/Missing` }));

  const sel = 'ID,Title,CJLookupMultiId,CJUserMultiId,CJUrl,ContentTypeId,AuthorId,EditorId,FileDirRef,FSObjType';
  for (const t of tests) {
    if (!t.id) continue;
    const i = await (await fetch(`${LIST}/items(${t.id})?$select=${sel}`, { headers: H })).json();
    t.back = { CJLookupMultiId: i.CJLookupMultiId, CJUserMultiId: (i.CJUserMultiId || []).map(alias), CJUrl: i.CJUrl,
      ct: ann && i.ContentTypeId === ann.StringId ? 'Announcement' : (i.ContentTypeId || '').slice(0, 8), AuthorId: alias(i.AuthorId), EditorId: alias(i.EditorId),
      folder: (i.FileDirRef || '').slice(listUrl.length) || '/', FSObjType: i.FSObjType };
  }
  out.expectedLookup = [a, b];
  out.tests = tests.map((t) => ({ title: t.title, status: t.status, errors: t.errors, back: t.back }));
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- a LookupMulti legalább egyik alakja `[a, b]`-t ad vissza;
- a UserMulti `["me","other"]`;
- az URL-nél kiderül a vessző szabálya;
- a tartalomtípus legalább egyik módon „Announcement” lesz;
- az E5 `AuthorId`/`EditorId` értéke `other`;
- az E6 elem a `/E6folder` mappába kerül, a nem létező almappa pedig hibát ad.

### E – eredmény (2026-09-25, Cél site, hu-HU)

Előkészítés rendben: mindkét oszlop létrejött, a tartalomtípusok és a mappák bekapcsolva, a Hirdetmény típus hozzáadva. Mellékes megfigyelés: nem létező mezőnél a `fields/getbyinternalnameortitle` **400**-at ad, nem 404-et. A site-on nem volt másik felhasználó, ezért az E5 kimaradt, az E2 pedig csak egy személlyel futott.

| Teszt | Küldött | Eredmény |
| --- | --- | --- |
| E1a LookupMulti | `1;#;#2` | ✅ `[1, 2]` |
| E1b LookupMulti | `1;#2` | ❌ csak `[1]`, hibaüzenet nélkül |
| E1c LookupMulti | `1;#;#2;#;#` | ✅ `[1, 2]` |
| E2 UserMulti | `[{"Key":…}]` | ✅ (csak egy felhasználóval igazolva) |
| E3a URL | `https://example.com/a,,b?x=1,,2, Leírás, vesszővel` | ✅ Url `…/a,b?x=1,2`, Description `Leírás, vesszővel` |
| E3b URL | `https://example.com/a,b, Leírás` | Url `…/a,b`, Description `Leírás`: a szóköz nélküli vessző is átment |
| E4a tartalomtípus név szerint | `ContentType` = `Hirdetmény` (a lokalizált név) | ✅ |
| E4b tartalomtípus ID szerint | `ContentTypeId` = a lista-tartalomtípus ID-ja | ✅ |
| E6 mappa | `UnderlyingObjectType: 1`, `LeafName` | ✅ mappa (`FSObjType: 1`) |
| E6 elem mappában | `FolderPath` = `…/E6folder` | ✅ `/E6folder` |
| E6b nem létező almappa | `FolderPath` = `…/E6folder/Missing` | ❌ **HTTP 500** (SPException „Nincs … URL-című fájl”), nem mezőszintű hiba |

→ **Döntések az `ItemProvider`-hez:**
- LookupMulti: `id;#;#id` (az `id;#id` csendben elveszíti az elemeket).
- UserMulti: `[{"Key":login}, …]`.
- URL: `<url, a vesszők duplázva>, <leírás>`. A leírásban a vessző szabadon maradhat.
- Tartalomtípus: `ContentTypeId`-t küldünk. A cél lista tartalomtípusát a sablonbeli szülő site-tartalomtípus ID alapján keressük meg (a lista-tartalomtípus ID = szülő + `00` + 32 hexa), mert a név nyelvfüggő.
- Mappák: a hiányzó mappákat (a szülőt előbb) az elemek előtt létre kell hozni, különben a kérés HTTP 500-zal elbukik.
- A Létrehozta/Módosította **másik** felhasználóra nincs igazolva. Ezt az első valódi telepítés napló ellenőrzi (figyelmeztetés, ha nem állítható).

Állapot: **lezárva** (a más szerzős eset a telepítési teszten ellenőrizendő).
