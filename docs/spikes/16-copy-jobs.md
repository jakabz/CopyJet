# Spike 16 – Hivatkozásos fájlmásolás (`CreateCopyJobs`)

Állapot: **lezárva** · Érintett kód (4. fázis): `FileProvider` hivatkozásos módja, séma `list.content`

## Háttér

Azonos tenanton belül a fájlokat nem kell letölteni és újra feltölteni: a `Site.CreateCopyJobs` szerveroldalon másol, aszinkron feladatként. A sablon ilyenkor csak a forrás URL-eket hordozza. A böngésző a feladat állapotát a `Site.GetCopyJobProgress` hívással kérdezi le, amely az új naplóbejegyzéseket adja vissza (`Logs`, JSON-szövegek tömbje).

## Kérdések

1. **Hívás helye:** a Cél site-on fut az Install. Elfogadja-e a Cél `_api/site/CreateCopyJobs` a Forrás URL-jeit, vagy a forrás site-on kell hívni?
2. **Mappa egyben:** egy mappa URL-je az almappákkal és fájlokkal együtt átmegy-e (`ExcludeChildren: false`)? Mennyi ideig tart, milyen `JobState` értékeket látunk?
3. **Metaadatok:**
   - megmarad-e a Létrehozta, a Módosította, a Létrehozva és a Módosítva;
   - átmegy-e a verzióelőzmény (`IgnoreVersionHistory: false`);
   - átmegy-e egy egyéni oszlop értéke, ha a céltárban ugyanaz az oszlop megvan (a CopyJet a tárat a mezőivel együtt előbb hozza létre).
4. **Azonosítók:** a cél `UniqueId` azonos-e a forráséval? Kiolvasható-e a naplóból a forrás → cél párosítás (`TargetObjectUniqueId`)? Ebből építjük a lookupokhoz szükséges elem-ID térképet.
5. **Névütközés:** mit ad ugyanannak a mappának a második másolása `NameConflictBehavior: 0` (Fail) esetén – hibát az egész feladatra, vagy fájlonként? Mit csinál az `1` (Replace) és a `2` (KeepBoth)? A CopyJet soha nem írhat felül, ezért a Fail viselkedése a lényeges.

## Előkészítés (Forrás)

A Forrás „Dokumentumok” tárában hozz létre egy **„CopyJetSpike16”** mappát:
- benne 2–3 fájl, köztük egy ékezetes nevű;
- egy almappa egy fájllal;
- legalább egy fájl, amelyet kétszer mentettél (legyen verzióelőzménye);
- ha a tárban van egysoros szöveges egyéni oszlop, egy fájlnál töltsd ki, és írd be a belső nevét a `SRC_COLUMN` értékébe.

## A kísérlet – Cél site (ír a saját tárába, a Forrást csak olvassa)

A kód a Célon minden futáskor új **„CopyJetSpike16-<időpont>”** dokumentumtárat hoz létre (verziókövetéssel, és ha a `SRC_COLUMN` meg van adva, ugyanazzal az oszloppal). Ezután négy másolást indít:

| Feladat | Mit másol | Hová | Verziók | Ütközés |
| --- | --- | --- | --- | --- |
| A | a forrásmappát | a céltár gyökerébe | igen | Fail (0) |
| B | ugyanazt a mappát még egyszer | a céltár gyökerébe | igen | Fail (0) |
| C | az első két fájlt külön URL-lel, egy feladatban | a már átmásolt mappába | nem | Replace (1) |
| D | az első fájlt | a már átmásolt mappába | nem | KeepBoth (2) |

A végén kiírja a forrás és a cél fájljait (szerző, dátumok, verzió, `UniqueId`, oszlopérték) és a feladatok naplóját. A naplóban a site-címek helyén `{forras}` és `{cel}` áll, a tenant nevét nem írja ki.

A tárat a végén **nem** törli, hogy megnézhesd; utána a „CopyJetSpike16…” tárakat nyugodtan töröld.

```js
(async () => {
  // ---- set these ----
  const SRC = 'https://…/sites/…/Shared Documents/CopyJetSpike16'; // the Forrás folder (absolute URL, copy it from the address bar)
  const SRC_COLUMN = ''; // optional: internal name of a single line text column of the Forrás library
  // -------------------
  const J = 'application/json;odata=nometadata';
  const siteOf = (u) => { const x = new URL(u); return x.origin + ((decodeURIComponent(x.pathname).match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]); };
  const dst = siteOf(location.href);
  const srcFolder = decodeURIComponent(new URL(SRC).pathname);
  const srcSite = siteOf(SRC);
  const host = location.origin;
  const mask = (v) => (typeof v === 'string' ? v.split(srcSite).join('{forras}').split(dst).join('{cel}').split(host).join('{host}') : v);
  const enc = (s) => encodeURIComponent(`'${s.replace(/'/g, "''")}'`);
  const digest = async (site) => (await (await fetch(site + '/_api/contextinfo', { method: 'POST', headers: { Accept: J } })).json()).FormDigestValue;
  const [DD, DS] = [await digest(dst), await digest(srcSite)];
  const call = async (site, dig, p, body) => {
    const r = await fetch(site + '/_api/' + p, { method: 'POST', headers: { Accept: J, 'Content-Type': J, 'X-RequestDigest': dig }, body: JSON.stringify(body) });
    const t = await r.text();
    return r.ok ? (t ? JSON.parse(t) : {}) : { httpStatus: r.status, error: mask(t.slice(0, 400)) };
  };
  const get = async (site, p) => { const r = await fetch(site + '/_api/' + p, { headers: { Accept: J } }); return r.ok ? r.json() : { httpStatus: r.status, error: mask((await r.text()).slice(0, 400)) }; };
  const out = { sameHost: new URL(SRC).origin === host };

  // Source library and (optional) column.
  const srcList = await get(srcSite, `web/GetFolderByServerRelativePath(decodedurl=@p)/ListItemAllFields/ParentList?$select=Id,Title&@p=${enc(srcFolder)}`);
  if (srcList.httpStatus) { console.log(JSON.stringify({ step: 'source library', ...srcList }, null, 2)); return; }
  const srcField = SRC_COLUMN ? await get(srcSite, `web/lists(guid'${srcList.Id}')/fields/getbyinternalnameortitle('${SRC_COLUMN}')?$select=SchemaXml,InternalName`) : undefined;

  // Target library.
  // A new library on every run (CopyJetSpike16-<time>), so the target is always empty.
  const LIB = `CopyJetSpike16-${new Date().toISOString().slice(11, 19).replace(/:/g, '')}`;
  const made = await call(dst, DD, 'web/lists', { Title: LIB, BaseTemplate: 101, EnableVersioning: true });
  if (made.httpStatus) { console.log(JSON.stringify({ step: 'create library', ...made }, null, 2)); return; }
  const lib = await get(dst, `web/lists/getbytitle('${LIB}')?$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`);
  out.targetLibrary = LIB;
  if (srcField && !srcField.httpStatus) {
    const has = await get(dst, `web/lists(guid'${lib.Id}')/fields/getbyinternalnameortitle('${SRC_COLUMN}')?$select=Id`);
    if (has.httpStatus) {
      const xml = srcField.SchemaXml.replace(/\s(ID|SourceID|Version|ColName|RowOrdinal|StaticName)="[^"]*"/g, '');
      out.targetColumn = await call(dst, DD, `web/lists(guid'${lib.Id}')/fields/createfieldasxml`, { parameters: { SchemaXml: xml, Options: 8 } }).then((r) => (r.httpStatus ? r : 'created'));
    } else out.targetColumn = 'already there';
  } else if (srcField) out.sourceColumn = srcField;
  const libRel = lib.RootFolder.ServerRelativeUrl;
  const folderName = srcFolder.replace(/\/$/, '').split('/').pop();

  // Files and folders under a folder, with what the copy should keep. GetItems refuses $expand=Author, so the
  // people come from the site users by ID.
  const people = async (site) => {
    const u = await get(site, 'web/siteusers?$select=Id,Title');
    const m = {};
    (u.value || []).forEach((x) => (m[x.Id] = x.Title));
    return m;
  };
  const [srcPeople, dstPeople] = [await people(srcSite), await people(dst)];
  const snap = async (site, dig, listId, folder) => {
    const col = srcField && !srcField.httpStatus ? `,${SRC_COLUMN}` : '';
    const who = site === srcSite ? srcPeople : dstPeople;
    const r = await call(site, dig, `web/lists(guid'${listId}')/GetItems?$select=Id,FSObjType,FileRef,UniqueId,Created,Modified,AuthorId,EditorId,OData__UIVersionString${col}`, {
      query: { ViewXml: '<View Scope="RecursiveAll"><RowLimit>200</RowLimit></View>', FolderServerRelativeUrl: folder }
    });
    if (r.httpStatus) return r;
    return (r.value || []).map((i) => ({
      path: i.FileRef.slice(folder.length + 1),
      type: i.FSObjType === 1 ? 'folder' : 'file',
      id: i.Id,
      uniqueId: i.UniqueId,
      created: i.Created,
      modified: i.Modified,
      author: who[i.AuthorId] || i.AuthorId,
      editor: who[i.EditorId] || i.EditorId,
      version: i.OData__UIVersionString,
      ...(col ? { column: i[SRC_COLUMN] } : {})
    }));
  };
  out.source = await snap(srcSite, DS, srcList.Id, srcFolder);

  // One copy job: create on Cél (or Forrás if Cél refuses), then poll until it ends.
  const job = async (label, uris, destination, options) => {
    const body = {
      exportObjectUris: uris,
      destinationUri: host + destination,
      options: { AllowSchemaMismatch: true, IgnoreVersionHistory: false, IsMoveMode: false, NameConflictBehavior: 0, IncludeItemPermissions: false, ExcludeChildren: false, SameWebCopyMoveOptimization: true, ...options }
    };
    const res = { label, calledOn: 'cel' };
    let site = dst, dig = DD;
    let r = await call(dst, DD, 'site/CreateCopyJobs', body);
    if (r.httpStatus) {
      res.celError = r;
      res.calledOn = 'forras';
      site = srcSite; dig = DS;
      r = await call(srcSite, DS, 'site/CreateCopyJobs', body);
    }
    if (r.httpStatus) return { ...res, error: r };
    const info = (r.value || [])[0];
    res.jobs = (r.value || []).length;
    res.sourceItemUniqueIds = info && (info.SourceListItemUniqueIds || []).length;
    const states = [], logs = [], t0 = Date.now();
    let idle = 0;
    while (info && Date.now() - t0 < 300000) {
      await new Promise((ok) => setTimeout(ok, 2000));
      const p = await call(site, dig, 'site/GetCopyJobProgress', { copyJobInfo: info });
      if (p.httpStatus) { logs.push(p); break; }
      if (states[states.length - 1] !== p.JobState) states.push(p.JobState);
      (p.Logs || []).forEach((l) => { try { logs.push(JSON.parse(l)); } catch (e) { logs.push(l); } });
      // A copy is an export and then an import, each with its own JobEnd: only the import's means done.
      const ended = logs.some((l) => l && (/JobFatalError|JobCancel/i.test(l.Event || '') || (l.Event === 'JobEnd' && l.MigrationDirection === 'Import')));
      idle = p.JobState === 0 ? idle + 1 : 0;
      if ((p.JobState === 0 && ended) || idle >= 15) break;
    }
    res.seconds = Math.round((Date.now() - t0) / 1000);
    res.states = states;
    res.logs = logs.map((l) => {
      if (typeof l !== 'object' || !l) return mask(l);
      const c = {};
      Object.keys(l).filter((k) => !/^(JobId|CorrelationId|SiteId|WebId|ListId|EncryptionKey)$/.test(k)).forEach((k) => (c[k] = mask(l[k])));
      return c;
    });
    return res;
  };

  out.jobs = [];
  out.jobs.push(await job('A folder, versions, Fail', [host + srcFolder], libRel, {}));
  out.targetAfterA = await snap(dst, DD, lib.Id, `${libRel}/${folderName}`);
  out.jobs.push(await job('B same folder again, Fail', [host + srcFolder], libRel, {}));
  const files = (Array.isArray(out.source) ? out.source : []).filter((f) => f.type === 'file' && f.path.indexOf('/') < 0).slice(0, 2);
  if (files.length) {
    out.jobs.push(await job('C two files, no versions, Replace', files.map((f) => `${host}${srcFolder}/${f.path}`), `${libRel}/${folderName}`, { IgnoreVersionHistory: true, NameConflictBehavior: 1 }));
    out.jobs.push(await job('D one file, no versions, KeepBoth', [`${host}${srcFolder}/${files[0].path}`], `${libRel}/${folderName}`, { IgnoreVersionHistory: true, NameConflictBehavior: 2 }));
  }
  out.targetFinal = await snap(dst, DD, lib.Id, `${libRel}/${folderName}`);
  console.log(JSON.stringify(out, null, 2));
})();
```

**Futtatás:**
1. A Cél site egy lapján nyisd meg a konzolt (F12 → Console).
2. A `SRC` értékébe másold be a Forrás „CopyJetSpike16” mappájának teljes URL-jét. Ezt a böngésző címsorából kapod meg, ha a tárban megnyitod a mappát; ha a címsorban `?id=…` van, azt a részt alakítsd útvonallá, vagy add meg kézzel: `https://<tenant>.sharepoint.com/sites/<Forrás>/Shared Documents/CopyJetSpike16`.
3. Ha van egyéni oszlop, írd be a belső nevét a `SRC_COLUMN` értékébe.
4. Futtasd. A négy feladat együtt pár percig is eltarthat; a kimenet egyetlen JSON a végén.

A kimenő JSON-t kérem vissza, és ha szemmel is megnézed a Cél „CopyJetSpike16” tárát (fájlok, verzióelőzmény, Létrehozta/Módosította), azt is írd meg.

## Eredmények

### 1. futás (2026-10-05, a snippet hibáival)

- **Hívás helye:** a Cél `_api/site/CreateCopyJobs` elfogadta a Forrás URL-jét (`calledOn: cel`); a `GetCopyJobProgress` is a Célon fut.
- **Mappa egyben:** az A feladat a teljes mappát átvitte (`FilesCreated: 11`, `ObjectsProcessed: 14`, kb. 850 KB) 2–3 másodperc alatt. A feladat két részből áll: export (Forrás), majd import (Cél), mindkettő saját `JobStart` / `JobEnd` bejegyzéssel. A `JobState` már az első lekérdezéskor 0 volt, ezért a vége csak az **import** `JobEnd`-jéből olvasható ki.
- **Forrás → cél párosítás:** `JobFinishedObjectInfo` csak a kijelölt objektumra (a mappára) jött, `TargetObjectUniqueId`-vel; a mappán belüli fájlokra nem.
- **Névütközés (Fail):** a B feladat a már létező mappán **egészében** elbukott (`JobError`, `ErrorCode -2147024713`, „Már létezik … nevű fájl vagy mappa a célhelyen”). Meglévő mappába tehát nem „egyesít” – a hiányzó fájlokat külön kell másolni.
- A snippet hibái: a `GetItems` nem fogadja az `$expand=Author`-t (400), a cél oszlop hiányára a `getbyinternalnameortitle` 400-at ad, nem 404-et, és a ciklus az export `JobEnd`-jénél kilépett. Javítva; a 2. futás minden alkalommal új céltárat hoz létre.

### 2. futás (2026-10-05)

Forrás: „CopyJetSpike16” mappa, 4 kép (egy ékezetes nevű, a `002.jpeg` 2.0-s verzióval és „Demo” oszlopértékkel), egy almappa egy 4.0-s verziójú szövegfájllal. A céltár üres, verziókövetéssel, a „Demo” oszloppal.

| Kérdés | Eredmény |
| --- | --- |
| Hívás helye | A Célon működik (`calledOn: cel` mind a négy feladatnál). |
| Mappa egyben | Az A feladat a mappát az almappával és mind az 5 fájllal átvitte, kb. 7 s alatt (az export 0,3 s, az import 5,7 s). Közben `JobProgress` bejegyzések jönnek; a `JobState` végig 0. |
| Verzióelőzmény | **Átmegy**: `002.jpeg` 2.0, `LICENSE.txt` 4.0 a célon is. |
| Egyéni oszlop | **Átmegy**, ha a céltárban ugyanaz a belső nevű oszlop megvan („Demo értéke”). |
| Létrehozva / Módosítva | **Nem marad meg**: mindkettő a másolás ideje lett (13:51 → 14:04). |
| Létrehozta / Módosította | Nem dönthető el – a forrásban minden fájl a futtatóé volt. |
| Azonosítók | Az elem-ID és a `UniqueId` is új a célon. `JobFinishedObjectInfo` csak a kijelölt objektumokra jön (mappánál a mappára), a benne lévő fájlokra nem. |
| Több URL egy hívásban | A C feladat két URL-jére **két külön feladat** jött vissza (`jobs: 2`) – URL-enként egy `copyJobInfo`, mindegyiket külön kell lekérdezni. (A snippet csak az elsőt figyelte.) |
| Fail (0) | Létező mappára vagy fájlra `JobError`, `ErrorCode -2147024713`; a mappát nem egyesíti, a feladat ezzel véget ér. |
| Replace (1) | Felülírja a fájlt: az elem-ID és a `UniqueId` marad, de az előzmény elvész (2.0 → 1.0). |
| KeepBoth (2) | Új név: `002.jpeg` → `0021.jpeg`. |

## Döntések

1. **Mikor:** hivatkozásos mód csak akkor, ha a Forrás és a Cél hostja azonos (`sourceUrl` originje = a Cél originje); különben a beágyazott tartalom kell, és ha nincs, a lépés `unsupported`. A séma ezt már tudja: `list.content.sourceMode: "reference"` + `sourceUrl`, a `_meta.json` fájlonként `sourceUrl`-lel, `blob` nélkül.
2. **Hívás:** a Cél `_api/site/CreateCopyJobs`-ja; a válasz URL-enként egy `copyJobInfo`, mindegyik `GetCopyJobProgress`-szel figyelve. Kész: az **import** `JobEnd`-je (vagy `JobFatalError`); fájlonkénti hiba: `JobError`.
3. **Ütközés:** mindig `NameConflictBehavior: 0` (Fail) – a CopyJet nem ír felül, a Replace ráadásul az előzményt is eldobja. Az „already exists” hiba (`-2147024713`) = már ott van, kihagyva.
4. **Szemcsézettség:** fájlonként egy URL, mappánként csoportosítva (egy hívás ≤ 50 URL, közös célmappa). A mappákat a CopyJet maga hozza létre előtte (a `ListProvider` a sablon mappáit, a `FileProvider` a hiányzókat), ezért egy egész mappa másolása a már létező célmappán úgyis elbukna. Így az újrafuttatás és a folytatás egyszerű: a már meglévő fájlok listázáskor kimaradnak, a másolás közben odakerültek pedig `exists` eredménnyel jönnek vissza. A mappaszintű másolás (kevesebb feladat) későbbi optimalizálás lehet üres céltárra.
5. **Metaadatok:** a másolás után ugyanaz a `ValidateUpdateListItem` fut, mint beágyazott módban (Létrehozta, Módosította, dátumok, és a leképezett mezők – lookup, személy, kifejezés –, mert ezek forrás-ID-ket hordoznának). A fájl útvonala alapján olvassuk ki a cél elem-ID-t, ebből lesz az ID-térkép.
6. **Verziók:** `IgnoreVersionHistory = !includeVersions`. Az előzmény szerzőit és dátumait a szerver viszi; ezt a valódi telepítésnél ellenőrizzük egy másik felhasználó verziójával.

## Megvalósítás (1.10.0.0)

- `core/files/copyJobs.ts`: `copyFiles` (feladatindítás, követés legfeljebb 4 párhuzamos lekérdezéssel, 2 s-onként, 10 perces időkorláttal) és `copyJobOutcome` (a napló értelmezése).
- `FileExtractor`: a Setup „Fájlok hivatkozással (azonos tenanton)” beállításával csak a `_meta.json` készül, fájlonként `sourceUrl`-lel, letöltés és méretkorlát nélkül; `content.sourceMode = "reference"`, `content.sourceUrl` = a forrástár abszolút URL-je.
- `FileProvider`: hivatkozásos módban más hostra `unsupported` (`otherTenant`); egyébként mappánként másol, majd ugyanúgy beállítja a metaadatokat és az ID-térképet, mint beágyazott módban.

## Ellenőrzés (1.10.0.0, azonos tenant)

Forrás „Dokumentumok” → Cél, verziókkal: 9 fájl hivatkozással átmásolva (kb. 26 s), 3 mappa létrehozva; az újrafuttatás mind a 9 fájlt meglévőként kihagyta (`FILES_KEPT`), másolás nélkül. A más felhasználó (Adele) által létrehozott fájlnál a Létrehozta helyes, a **korábbi verziók szerzője viszont mindenhol a telepítő** – a szerveroldali másolás is így hozza létre őket, és a verziók metaadata utólag nem írható. Ugyanez a beágyazott mód ismert korlátja (spike 10).

Másik tenant (fabrikam): a struktúra települt (a tár a meglévő „Megosztott dokumentumok”-ba, `LIST_MATCHED`), a fájllépés `FILES_UNSUPPORTED` / `otherTenant` figyelmeztetéssel kimaradt.
