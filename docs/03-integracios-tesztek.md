# CopyJet – integrációs és regressziós tesztek

Ez a dokumentum a kiadásonkénti kézi ellenőrzés forgatókönyve. Az automatikus részt a `npx jest` futtatja (egység-, szerződés- és regressziós tesztek, `tests/regression/`); ami csak valódi SharePointon ellenőrizhető, az itt van lépésenként, kitölthető eredménnyel.

## 1. Környezet

| Név | Mi ez | Megjegyzés |
| --- | --- | --- |
| **Forrás** | Magyar nyelvű csoportwebhely az első tenantban | Itt fut a Setup; a 2. pont generátora ide ír |
| **Cél** | Magyar nyelvű site ugyanabban a tenantban | Azonos tenantos telepítés, hivatkozásos fájlmásolás |
| **Cél EN** | Angol nyelvű site ugyanabban a tenantban | Nyelvi eltérés (`Shared Documents` ↔ `Megosztott dokumentumok`, beépített szintnevek) |
| **Cél2** | Site a második tenantban | Más tenant: felhasználó-leképezés, hiányzó felhasználó, hivatkozásos mód kizárása |

Mindegyiken telepítve a vizsgált `.sppkg` verzió, a lapokon a Setup (Forrás) és az Install (a többi) webpart. A futtató a Forráson olvasó, a célokon site-gyűjtemény-adminisztrátor (a 3.9 eset kivételével).

A tesztek előtt a célokon a korábbi futás listái, tárai, csoportjai törölhetők, de nem kötelező: a 3.10 eset éppen a meglévőkre telepítést vizsgálja.

## 2. Tesztadat a Forráson (generátor)

A konzolkód a Forráson a saját, **„CJ Teszt”** kezdetű listáit és mappáit hozza létre és tölti fel; mást nem módosít. Ha egy lista már megvan, nem hozza létre újra, és az elemeit nem duplikálja (az elemszámot a meglévőkhöz igazítja).

| Rész | Mit hoz létre |
| --- | --- |
| `LOOKUP` | „CJ Teszt A” és „CJ Teszt B” lista: B → A lookup (lánc), A → B lookup (ciklus), 5–5 elem kölcsönös hivatkozással |
| `BIG_LIST` | „CJ Teszt nagy lista” 6500 elemmel (Cím + Szám oszlop) |
| `BIG_FILE` | „Dokumentumok/CJ Teszt” mappában egy 100 MB-os `nagy-100MB.bin`, darabolt feltöltéssel |
| `VERSIONS` | „Dokumentumok/CJ Teszt” mappában `verziozott.txt`, háromszor mentve (1.0, 2.0, 3.0) |

A futás a végén összesítést ír ki. A `BIG_LIST` 3–6 percig, a `BIG_FILE` a hálózattól függően pár percig tart.

```js
(async () => {
  // ---- choose the parts ----
  const PARTS = { LOOKUP: true, BIG_LIST: true, BIG_FILE: true, VERSIONS: true };
  const BIG_COUNT = 6500;
  // ---------------------------
  const site = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const J = 'application/json;odata=nometadata';
  const digest = async () => (await (await fetch(site + '/_api/contextinfo', { method: 'POST', headers: { Accept: J } })).json()).FormDigestValue;
  let D = await digest();
  setInterval(async () => (D = await digest()), 20 * 60 * 1000);
  const req = async (method, p, body, raw) => {
    const r = await fetch(site + '/_api/' + p, { method, headers: { Accept: J, 'X-RequestDigest': D, ...(raw ? {} : { 'Content-Type': J }) }, body: raw ? body : body === undefined ? undefined : JSON.stringify(body) });
    if (r.status === 429 || r.status === 503) { await new Promise((ok) => setTimeout(ok, (Number(r.headers.get('Retry-After')) || 5) * 1000)); return req(method, p, body, raw); }
    const t = await r.text();
    if (!r.ok) throw new Error(`${method} ${p}: HTTP ${r.status} ${t.slice(0, 300)}`);
    return t ? JSON.parse(t) : {};
  };
  const q = (s) => encodeURIComponent(`'${s.replace(/'/g, "''")}'`);
  const web = await req('GET', 'web?$select=ServerRelativeUrl');
  const out = {};

  const ensureList = async (title, template = 100) => {
    try { return await req('GET', `web/lists/getbytitle(${q(title)})?$select=Id,ItemCount`); }
    catch (e) { await req('POST', 'web/lists', { Title: title, BaseTemplate: template }); return req('GET', `web/lists/getbytitle(${q(title)})?$select=Id,ItemCount`); }
  };
  const ensureField = async (list, name, xml) => {
    try { await req('GET', `web/lists(guid'${list.Id}')/fields/getbyinternalnameortitle('${name}')?$select=Id`); }
    catch (e) { await req('POST', `web/lists(guid'${list.Id}')/fields/createfieldasxml`, { parameters: { SchemaXml: xml, Options: 8 } }); }
  };
  const addItems = async (list, count, make, parallel = 6) => {
    let next = 0;
    await Promise.all(Array.from({ length: parallel }, async () => { while (next < count) { const i = next++; await req('POST', `web/lists(guid'${list.Id}')/items`, make(i)); } }));
  };

  if (PARTS.LOOKUP) {
    const a = await ensureList('CJ Teszt A');
    const b = await ensureList('CJ Teszt B');
    await ensureField(b, 'CJRefA', `<Field Type="Lookup" Name="CJRefA" DisplayName="A hivatkozás" List="{${a.Id}}" ShowField="Title" />`);
    await ensureField(a, 'CJRefB', `<Field Type="Lookup" Name="CJRefB" DisplayName="B hivatkozás" List="{${b.Id}}" ShowField="Title" />`);
    if (a.ItemCount === 0) await addItems(a, 5, (i) => ({ Title: `A${i + 1}` }), 1);
    if (b.ItemCount === 0) await addItems(b, 5, (i) => ({ Title: `B${i + 1}` }), 1);
    const as = (await req('GET', `web/lists(guid'${a.Id}')/items?$select=Id&$orderby=Id`)).value;
    const bs = (await req('GET', `web/lists(guid'${b.Id}')/items?$select=Id&$orderby=Id`)).value;
    // B[i] → A[i] (chain), A[i] → B[(i+1) % 5] (cycle back).
    for (let i = 0; i < 5; i++) {
      await req('POST', `web/lists(guid'${b.Id}')/items(${bs[i].Id})/validateupdatelistitem`, { formValues: [{ FieldName: 'CJRefA', FieldValue: String(as[i].Id) }], bNewDocumentUpdate: false });
      await req('POST', `web/lists(guid'${a.Id}')/items(${as[i].Id})/validateupdatelistitem`, { formValues: [{ FieldName: 'CJRefB', FieldValue: String(bs[(i + 1) % 5].Id) }], bNewDocumentUpdate: false });
    }
    out.lookup = 'CJ Teszt A ↔ CJ Teszt B: 5–5 elem, kölcsönös lookup';
  }

  if (PARTS.BIG_LIST) {
    const big = await ensureList('CJ Teszt nagy lista');
    await ensureField(big, 'CJSzam', '<Field Type="Number" Name="CJSzam" DisplayName="Szám" />');
    const missing = Math.max(0, BIG_COUNT - big.ItemCount);
    const t0 = Date.now();
    await addItems(big, missing, (i) => ({ Title: `Elem ${big.ItemCount + i + 1}`, CJSzam: big.ItemCount + i + 1 }));
    out.bigList = `CJ Teszt nagy lista: ${big.ItemCount + missing} elem (${missing} új, ${Math.round((Date.now() - t0) / 1000)} s)`;
  }

  const lib = await req('GET', `web/lists/getbytitle(${q('Dokumentumok')})?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`).catch(() => req('GET', `web/lists/getbytitle(${q('Documents')})?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`));
  const folder = `${lib.RootFolder.ServerRelativeUrl}/CJ Teszt`;
  if (PARTS.BIG_FILE || PARTS.VERSIONS) await req('POST', `web/folders/addUsingPath(decodedurl=${q(folder)})`).catch(() => undefined);

  if (PARTS.BIG_FILE) {
    const SIZE = 100 * 1024 * 1024, CHUNK = 10 * 1024 * 1024, name = 'nagy-100MB.bin';
    const block = new Uint8Array(65536);
    crypto.getRandomValues(block);
    const chunk = new Uint8Array(CHUNK);
    for (let o = 0; o < CHUNK; o += block.length) chunk.set(block, o);
    const id = crypto.randomUUID();
    const file = `web/getFileByServerRelativePath(decodedurl=${q(`${folder}/${name}`)})`;
    const t0 = Date.now();
    await req('POST', `web/getFolderByServerRelativePath(decodedurl=${q(folder)})/files/addUsingPath(decodedurl=${q(name)},overwrite=true)`, '', true);
    let offset = 0;
    while (offset < SIZE) {
      const last = offset + CHUNK >= SIZE;
      const part = chunk.slice(0, Math.min(CHUNK, SIZE - offset));
      const op = offset === 0 ? `startUpload(uploadId=guid'${id}')` : last ? `finishUpload(uploadId=guid'${id}',fileOffset=${offset})` : `continueUpload(uploadId=guid'${id}',fileOffset=${offset})`;
      await req('POST', `${file}/${op}`, part, true);
      offset += part.length;
      console.log(`100 MB: ${Math.round(offset / 1048576)} MB`);
    }
    out.bigFile = `${folder.split('/').slice(-2).join('/')}/${name}: 100 MB (${Math.round((Date.now() - t0) / 1000)} s)`;
  }

  if (PARTS.VERSIONS) {
    for (let v = 1; v <= 3; v++) {
      await req('POST', `web/getFolderByServerRelativePath(decodedurl=${q(folder)})/files/addUsingPath(decodedurl=${q('verziozott.txt')},overwrite=true)`, new Blob([`CopyJet verzió ${v}\n`]), true);
    }
    const f = await req('GET', `web/getFileByServerRelativePath(decodedurl=${q(`${folder}/verziozott.txt`)})?$select=UIVersionLabel`);
    out.versions = `verziozott.txt: ${f.UIVersionLabel}`;
  }
  console.log(JSON.stringify(out, null, 2));
})();
```

**Futtatás:** a Forrás egy lapján F12 → Console, a kód bemásolása, Enter. Ha a nagy fájl vagy a nagy lista már megvan, a `PARTS`-ban kikapcsolható.

## 3. Tesztesetek

Minden esetnél: a futás végén a napló exportja (`…-install-….json`) és az ellenőrzés eredménye. Az „Eredmény” oszlopba: OK / hiba (rövid leírás, napló neve).

### 3.1 Lookup-lánc és ciklus

1. Setup a Forráson: „CJ Teszt A” és „CJ Teszt B” szerkezet + tartalom.
2. Install a Célra.
3. **Várt:** a lookup-oszlopok a céllisták után jönnek létre; mivel az elemek kölcsönösen egymásra mutatnak, a lookup-értékek egy külön lépésben (`itemLookups`) kerülnek fel, amikor már mindkét lista elemei megvannak. Mindkét listában 5 elem; minden B elem a megfelelő A elemre, minden A elem a következő B elemre mutat (a Célon az új azonosítókkal).

### 3.2 6000-nél több elemű lista

1. Setup: „CJ Teszt nagy lista” szerkezet + tartalom.
2. Install a Célra.
3. **Várt:** 6500 elem a Célon, a „Szám” oszlop 1–6500; a felület nem fagy le, a haladásjelző folyamatosan mozog. A felderítés és a kiolvasás a 5000-es listanézet-küszöb ellenére hiba nélkül fut.
4. Install még egyszer ugyanoda: **várt** a „céllistában már vannak elemek” jelzés, nem duplikál.

### 3.3 100 MB-os fájl (beágyazott mód)

1. Setup: „Dokumentumok” szerkezet + tartalom, max. fájlméret 250 MB, hivatkozás nélkül.
2. Install a **Cél2**-re (másik tenant).
3. **Várt:** a `.zip` kb. 100 MB-tal nagyobb; a `nagy-100MB.bin` a Cél2-n pontosan 104 857 600 bájt, darabolt feltöltéssel. A max. fájlméret 50 MB-ra állításával a fájl `FILE_TOO_LARGE` figyelmeztetéssel kimarad.

### 3.4 Verziózott tár

1. Setup: „Dokumentumok” szerkezet + tartalom, „Verziók” bekapcsolva, hivatkozás nélkül.
2. Install a Cél2-re.
3. **Várt:** `verziozott.txt` a Cél2-n 3.0-s verzióval, az előzményben 1.0, 2.0, 3.0 a forrás dátumaival; a korábbi verziók szerzője a telepítő (ismert korlát).
4. Ugyanez **hivatkozással** a Célra (azonos tenant): verziók a szerveroldali másolással, `Files copied by reference`.

### 3.5 Lap lista- és kép-webparttal

1. Setup: a kezdőlap (lista-, kép-, gyorshivatkozás-webparttal) és a hivatkozott listák.
2. Install a Célra és a Cél2-re.
3. **Várt:** a lap megjelenik, a lista-webpart a cél listáját mutatja, a kép a cél SiteAssets-ből töltődik, a gyorshivatkozások a cél site-ra mutatnak. Nincs a forrásra mutató hivatkozás (a naplóban nincs „Page references the target cannot serve”).

### 3.6 Magyar ↔ angol nyelvű site

1. A 3.5 sablonja és egy „Dokumentumok” tartalmú sablon telepítése a **Cél EN**-re.
2. **Várt:** a tár a meglévő „Documents” tárba kerül (`LIST_MATCHED`), nem jön létre második; a beépített jogosultsági szintek (Olvasás/Read, Szerkesztés/Edit) helyesen párosulnak; a navigáció és a lap működik.

### 3.7 Hiányzó felhasználó

1. Setup: egy lista „Személy” oszloppal, olyan felhasználóval kitöltve, aki a Cél2 tenantban nincs meg; az egyedi listajogosultságai között is szerepel.
2. Install a Cél2-re; a leképezésnél a felhasználó maradjon leképezés nélkül (helyettesítő felhasználó).
3. **Várt:** az elemek a helyettesítő felhasználóval jönnek létre, a napló jelzi; a listajogosultságnál a felhasználó kimarad (`LIST_SECURITY_PRINCIPAL_MISSING`), a helyettesítő nem kapja meg a jogait.

### 3.8 Megszakított és folytatott telepítés

1. Install a 3.2 sablonjával egy üres célra; a nagy lista elemírása közben **Leállítás**.
2. A lap újratöltése után a sávban **Folytatás**.
3. Ugyanez a lap bezárásával az elemírás közben (nem Leállítással).
4. **Várt:** mindkét esetben pontosan 6500 elem, duplikáció nélkül; a napló `Resuming run …` és az elemeknél `ITEMS_RESUMED` / `ITEMS_RESUME_RECOVERED`. Egy végigfutott telepítés után a sáv nem jelenik meg.

### 3.9 Nem adminisztrátor telepítő

1. Egy olyan felhasználóval, aki a Célon csak tulajdonos (nem site-gyűjtemény-adminisztrátor), telepíts egyedi jogosultságú listát.
2. **Várt:** a lista öröklése megszűnik, a sablon hozzárendelései felkerülnek, a telepítő Teljes hozzáférése megmarad (`LIST_SECURITY_INSTALLER_KEPT`).

### 3.10 Újrafuttatás meglévő elemekre (idempotencia)

1. Bármelyik fenti sablon második telepítése ugyanoda, „Kihagyás” módban.
2. **Várt:** semmi nem duplikálódik, semmi nem íródik felül; a napló „already present; skipped” / „kept as it is” bejegyzésekből áll. „Frissítés” módban a különbségek (pl. nézetmezők, csoportszerepek) frissülnek, törlés nincs.

## 4. Eredmények

| Eset | Verzió | Dátum | Cél | Eredmény | Napló / megjegyzés |
| --- | --- | --- | --- | --- | --- |
| 3.1 | | | Cél | | |
| 3.2 | | | Cél | | |
| 3.3 | | | Cél2 | | |
| 3.4 | | | Cél2, Cél | | |
| 3.5 | | | Cél, Cél2 | | |
| 3.6 | | | Cél EN | | |
| 3.7 | | | Cél2 | | |
| 3.8 | | | Cél | | |
| 3.9 | | | Cél | | |
| 3.10 | | | bármelyik | | |

## 5. Regresszió (automatikus)

`npx jest` része a `tests/regression/` mappa:
- a `schema/examples` teljes mintasablonja csomagként: megnyitás, séma-ellenőrzés, ellenőrzőösszeg; a terv minden elemet tartalmaz, függőségi sorrendben (a sorrend pillanatképe a `__snapshots__` mappában – szándékos változásnál `npx jest -u`);
- a telepítés minden lépést egyszer futtat, az újrafuttatás a mentett állapotból semmit nem ismétel, a megszakított futás a mentett állapotból folytatható;
- a `tests/fixtures/templates/` minden valódi, Setupból exportált (anonimizált) sablonja megnyílik és kizárás nélkül tervezhető. Új kiadás előtt érdemes egy friss, teljes exportot ide tenni (a tenant nevét `contoso`-ra cserélve).
