# Spike 04 – Tartalomtípus listához rendelése, alapértelmezett tartalomtípus, mappák

Állapot: **lezárva** (2026-09-24) · Érintett kód: `ListExtractor`, `ListProvider` (tartalomtípusok és mappák)

## Kérdések

1. **Hozzárendelés:** a `POST <lista>/contenttypes/addAvailableContentType('<site CT ID>')` REST-hívás működik-e? Mit ad vissza: a lista-szintű tartalomtípus ID-ját (site CT ID + `00` + GUID)?
2. **Alapértelmezett tartalomtípus:** a gyökérmappa `UniqueContentTypeOrder` tulajdonsága (az első elem az alapértelmezett) beállítható-e REST-ből (`MERGE`, verbose)?
   - Ha nem: beállítható-e CSOM-mal (`Folder.UniqueContentTypeOrder` + `Folder.Update()`)?
   - A CSOM-hoz az `SP.ContentTypeId` típusazonosítója kell, ezt a kód a site `sp.js`-éből olvassa ki.
3. **Mappák:**
   - F1: egy tárban a `folders/addUsingPath` egy hívással létrehoz-e egymásba ágyazott mappákat (`A/B/C`)?
   - F2: egyedi listában is működik-e a `folders/addUsingPath`?
   - F3: listában működik-e a mappa létrehozása elemként (`AddValidateUpdateItemUsingPath`, `UnderlyingObjectType: 1`)?

## Ideiglenes feltevések

- **Kinyerés:** a lista tartalomtípusai a site-szintű szülő ID-jával kerülnek a sablonba. A lista-szintű ID a site CT ID + `00` + GUID, a szülőt a `parentIdOf` adja. Az első elem az alapértelmezett. A rejtett tartalomtípusok (pl. Mappa) kimaradnak.
- **Mappák:** a provider szintenként halad (előbb `A`, aztán `A/B`), mindegyiknél ellenőrzi, hogy létezik-e már. Így a Q3/F1 válaszától függetlenül működik, az eredmény legfeljebb egyszerűsíthet rajta.

## Kísérlet – Cél site (**ír**)

Csak a Cél **teszt** site-on futtasd. Semmit nem töröl. A következőket módosítja:
- **`CopyJet Spike B2 ékezetes` lista (`Lists/CopyJetSpikeB2`):**
  - bekapcsolja rajta a tartalomtípusokat;
  - hozzáadja a `CopyJet Spike CSOM` tartalomtípust, és megpróbálja alapértelmezetté tenni;
  - létrehoz két mappát (`CopyJetMappa1`, `CopyJetMappa2`).
- **Dokumentumok tár:** létrehozza a `CopyJetSpike/2026/Q3` mappát.

A teszt után a mappák és a lista kézzel törölhetők.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const V = { Accept: 'application/json;odata=verbose' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const WV = { ...V, 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest };
  const short = async (r) => { try { return (await r.text()).slice(0, 300); } catch { return ''; } };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const listUrl = `${rel}/Lists/CopyJetSpikeB2`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  const SITE_CT = '0x0100C0FFEE02C0FFEE02C0FFEE02C0FFEE02';
  const out = { web };

  const listCts = async () => ((await (await fetch(`${LIST}/contenttypes?$select=StringId,Name,Hidden`, { headers: H })).json()).value || []).map((c) => `${c.Name}: ${c.StringId}${c.Hidden ? ' (rejtett)' : ''}`);
  const order = async () => {
    const f = await (await fetch(`${LIST}/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`, { headers: H })).json();
    return { order: (f.ContentTypeOrder || []).map((c) => c.StringValue), unique: f.UniqueContentTypeOrder ? f.UniqueContentTypeOrder.map((c) => c.StringValue) : null };
  };

  // Q1: enable content types, add the site CT to the list
  let r = await fetch(LIST, { method: 'POST', headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ ContentTypesEnabled: true }) });
  out.q1_enableStatus = r.status;
  const before = await listCts();
  if (!before.some((c) => c.indexOf(SITE_CT) >= 0)) {
    r = await fetch(`${LIST}/contenttypes/addAvailableContentType('${SITE_CT}')`, { method: 'POST', headers: W });
    out.q1_addStatus = r.status;
    out.q1_addBody = r.ok ? ((await r.json()).StringId || '') : await short(r);
  }
  out.q1_listCts = await listCts();
  const listCt = ((await (await fetch(`${LIST}/contenttypes?$select=StringId`, { headers: H })).json()).value || []).map((c) => c.StringId).find((id) => id.indexOf(SITE_CT) === 0);
  out.q2_before = await order();

  // Q2a: REST (verbose) MERGE UniqueContentTypeOrder on the root folder
  if (listCt) {
    const rest = out.q2_before.order.filter((id) => id !== listCt && !/^0x0120/i.test(id));
    const ids = [listCt].concat(rest);
    r = await fetch(`${LIST}/RootFolder`, { method: 'POST', headers: { ...WV, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
      body: JSON.stringify({ __metadata: { type: 'SP.Folder' }, UniqueContentTypeOrder: { __metadata: { type: 'Collection(SP.ContentTypeId)' }, results: ids.map((id) => ({ __metadata: { type: 'SP.ContentTypeId' }, StringValue: id })) } }) });
    out.q2a_status = r.status; out.q2a_body = r.ok ? '' : await short(r);
    out.q2a_after = await order();

    // Q2b: CSOM, only if REST did not make it the default
    if (out.q2a_after.order[0] !== listCt) {
      const js = await (await fetch(`${web}/_layouts/15/sp.js`)).text();
      const m = /SP\.ContentTypeId\.prototype\s*=\s*\{[\s\S]{0,3000}?typeId[\s\S]{0,120}?\{([0-9a-fA-F-]{36})\}/.exec(js);
      out.q2b_contentTypeIdTypeId = m ? m[1] : null;
      if (m) {
        const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
        const objs = ids.map((id) => `<Object TypeId="{${m[1]}}"><Property Name="StringValue" Type="String">${esc(id)}</Property></Object>`).join('');
        const xml = '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CopyJet">' +
          `<Actions><SetProperty Id="10" ObjectPathId="5" Name="UniqueContentTypeOrder"><Parameter Type="Array">${objs}</Parameter></SetProperty>` +
          '<Method Name="Update" Id="11" ObjectPathId="5" /></Actions>' +
          '<ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="1" Name="Web" />' +
          `<Method Id="4" ParentId="2" Name="GetList"><Parameters><Parameter Type="String">${esc(listUrl)}</Parameter></Parameters></Method>` +
          '<Property Id="5" ParentId="4" Name="RootFolder" /></ObjectPaths></Request>';
        r = await fetch(web + '/_vti_bin/client.svc/ProcessQuery', { method: 'POST', headers: { 'Content-Type': 'text/xml', Accept: 'application/json', 'X-RequestDigest': digest }, body: xml });
        const body = await r.text();
        out.q2b_status = r.status;
        out.q2b_error = (body.match(/"ErrorMessage":"([^"]*)"/) || [])[1] || null;
        out.q2b_after = await order();
      }
    }
  }

  // F1: nested folders in one call (library)
  const folderExists = async (path) => (await (await fetch(`${web}/_api/web/getFolderByServerRelativePath(decodedurl='${enc(path)}')?$select=Exists`, { headers: H })).json()).Exists === true;
  const deep = `${rel}/Shared Documents/CopyJetSpike/2026/Q3`;
  r = await fetch(`${web}/_api/web/folders/addUsingPath(DecodedUrl='${enc(deep)}')`, { method: 'POST', headers: W });
  out.f1_status = r.status; out.f1_body = r.ok ? '' : await short(r);
  out.f1_exists = { parent: await folderExists(`${rel}/Shared Documents/CopyJetSpike`), deep: await folderExists(deep) };

  // F2: folder in a custom list via folders/addUsingPath
  r = await fetch(`${web}/_api/web/folders/addUsingPath(DecodedUrl='${enc(listUrl + '/CopyJetMappa1')}')`, { method: 'POST', headers: W });
  out.f2_status = r.status; out.f2_body = r.ok ? '' : await short(r);
  out.f2_exists = await folderExists(listUrl + '/CopyJetMappa1');

  // F3: folder in a custom list as an item
  r = await fetch(`${LIST}/AddValidateUpdateItemUsingPath`, { method: 'POST', headers: W,
    body: JSON.stringify({ listItemCreateInfo: { FolderPath: { DecodedUrl: listUrl }, UnderlyingObjectType: 1, LeafName: 'CopyJetMappa2' }, formValues: [{ FieldName: 'Title', FieldValue: 'CopyJetMappa2' }], bNewDocumentUpdate: false }) });
  out.f3_status = r.status; out.f3_body = (await short(r)).slice(0, 200);
  out.f3_exists = await folderExists(listUrl + '/CopyJetMappa2');

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## Mit várunk?

- `q1_addStatus` = 200, és a `q1_listCts` között szerepel egy `0x0100C0FFEE02…00<GUID>` ID-jú „CopyJet Spike CSOM”.
- `q2a_after.order[0]` vagy `q2b_after.order[0]` a lista-szintű „CopyJet Spike CSOM” ID-ja. Ez mondja meg, melyik út állítja be az alapértelmezettet.
- `f1_exists.deep` = `true`, ha egy hívás elég az egymásba ágyazott mappákhoz.
- `f2_exists` / `f3_exists`: melyik út működik egyedi listában.

## Eredmény

### 1. futás (2026-09-24, Cél site)

- **Q1 – nem értékelhető, a kísérlet hibája miatt:** a `CopyJet Spike CSOM` tartalomtípust a spike 02 takarításakor töröltük, a kód mégis erre hivatkozott. A hiba: „A tartalomtípus nem található”, HTTP 500. Emiatt a Q2 sem futott le.
  - **Ami így is kiderült:** a lista-szintű tartalomtípus ID-ja **site-szintű ID + `00` + GUID**. Például az Elem (`0x01`) a listán `0x0100413672950A8A8C4ABF3C45941B6B1EEF`, a Mappa (`0x0120`) `0x012000A8C0F65E086B7545AF19C8E2196FAC2B`. A `parentIdOf` tehát a lista-szintű ID-ból a site-szintűt adja.
  - A `ContentTypeOrder` csak a látható típusokat tartalmazza (`[Elem]`), a `UniqueContentTypeOrder` alapból `null`.
- **F1 – egymásba ágyazott mappa egy hívással: ❌.** A `folders/addUsingPath('…/CopyJetSpike/2026/Q3')` HTTP 500-at adott: „…/CopyJetSpike/2026/Q3 nem található”. A szülő sem jött létre.
  → A provider szintenként hoz létre mappát, a legfelső szinttel kezdve.
- **F2 – `folders/addUsingPath` egyedi listában: ✅.** HTTP 200, a mappa létrejött. Nem kell hozzá elem-alapú létrehozás, és a lista `EnableFolderCreation` beállítása sem akadályozta.
- **F3 – mappa elemként:** a hívás formátuma hibás volt (HTTP 400, OData-olvasási hiba). Mivel az F2 működik, nem folytatjuk.

### 2. futás – Q1/Q2 beépített site-tartalomtípussal

Csak a Cél **teszt** site-on futtasd. A `CopyJet Spike B2 ékezetes` listához hozzáadja a beépített **Hirdetmény** (`0x0104`) vagy, ha az nincs, a **Hivatkozás** (`0x0105`) tartalomtípust, és megpróbálja alapértelmezetté tenni. Semmit nem töröl.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const V = { Accept: 'application/json;odata=verbose' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const WV = { ...V, 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest };
  const short = async (r) => { try { return (await r.text()).slice(0, 300); } catch { return ''; } };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const listUrl = `${rel}/Lists/CopyJetSpikeB2`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  const out = { web };

  // A built-in site content type that always exists: Announcement (0x0104) or Link (0x0105).
  let SITE_CT = null;
  for (const id of ['0x0104', '0x0105']) {
    const c = await (await fetch(`${web}/_api/web/availablecontenttypes('${id}')?$select=StringId,Name`, { headers: H })).json();
    if (c.StringId) { SITE_CT = c.StringId; out.siteCt = `${c.Name}: ${c.StringId}`; break; }
  }
  const listCtIds = async () => ((await (await fetch(`${LIST}/contenttypes?$select=StringId,Name,Hidden`, { headers: H })).json()).value || []);
  const order = async () => {
    const f = await (await fetch(`${LIST}/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`, { headers: H })).json();
    return { order: (f.ContentTypeOrder || []).map((c) => c.StringValue), unique: f.UniqueContentTypeOrder ? f.UniqueContentTypeOrder.map((c) => c.StringValue) : null };
  };
  const childOf = (id) => id.toUpperCase().indexOf(SITE_CT.toUpperCase() + '00') === 0 && id.length === SITE_CT.length + 34;

  // Q1: add the site CT to the list
  if (!(await listCtIds()).some((c) => childOf(c.StringId))) {
    const r = await fetch(`${LIST}/contenttypes/addAvailableContentType('${SITE_CT}')`, { method: 'POST', headers: W });
    out.q1_addStatus = r.status;
    out.q1_addBody = r.ok ? ((await r.json()).StringId || '') : await short(r);
  }
  out.q1_listCts = (await listCtIds()).map((c) => `${c.Name}: ${c.StringId}${c.Hidden ? ' (rejtett)' : ''}`);
  const listCt = (await listCtIds()).map((c) => c.StringId).find(childOf);
  out.q2_before = await order();

  if (listCt) {
    const ids = [listCt].concat(out.q2_before.order.filter((id) => id !== listCt && !/^0x0120/i.test(id)));
    // Q2a: REST (verbose) MERGE UniqueContentTypeOrder
    let r = await fetch(`${LIST}/RootFolder`, { method: 'POST', headers: { ...WV, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
      body: JSON.stringify({ __metadata: { type: 'SP.Folder' }, UniqueContentTypeOrder: { __metadata: { type: 'Collection(SP.ContentTypeId)' }, results: ids.map((id) => ({ __metadata: { type: 'SP.ContentTypeId' }, StringValue: id })) } }) });
    out.q2a_status = r.status; out.q2a_body = r.ok ? '' : await short(r);
    out.q2a_after = await order();

    // Q2b: CSOM, only if REST did not make it the default
    if (out.q2a_after.order[0] !== listCt) {
      const js = await (await fetch(`${web}/_layouts/15/sp.js`)).text();
      const m = /SP\.ContentTypeId\.prototype\s*=\s*\{[\s\S]{0,3000}?typeId[\s\S]{0,120}?\{([0-9a-fA-F-]{36})\}/.exec(js);
      out.q2b_contentTypeIdTypeId = m ? m[1] : null;
      if (m) {
        const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
        const objs = ids.map((id) => `<Object TypeId="{${m[1]}}"><Property Name="StringValue" Type="String">${esc(id)}</Property></Object>`).join('');
        const xml = '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CopyJet">' +
          `<Actions><SetProperty Id="10" ObjectPathId="5" Name="UniqueContentTypeOrder"><Parameter Type="Array">${objs}</Parameter></SetProperty>` +
          '<Method Name="Update" Id="11" ObjectPathId="5" /></Actions>' +
          '<ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="1" Name="Web" />' +
          `<Method Id="4" ParentId="2" Name="GetList"><Parameters><Parameter Type="String">${esc(listUrl)}</Parameter></Parameters></Method>` +
          '<Property Id="5" ParentId="4" Name="RootFolder" /></ObjectPaths></Request>';
        r = await fetch(web + '/_vti_bin/client.svc/ProcessQuery', { method: 'POST', headers: { 'Content-Type': 'text/xml', Accept: 'application/json', 'X-RequestDigest': digest }, body: xml });
        const body = await r.text();
        out.q2b_status = r.status;
        out.q2b_error = (body.match(/"ErrorMessage":"([^"]*)"/) || [])[1] || null;
        out.q2b_after = await order();
      }
    }
  }

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- `q1_addStatus` = 200, és a `q1_listCts` között a Hirdetmény `0x0104` + `00` + GUID ID-val szerepel.
- `q2a_after.order[0]` vagy `q2b_after.order[0]` a lista-szintű Hirdetmény ID-ja. Ez mondja meg, melyik út állítja be az alapértelmezettet.

### 2. futás – eredmény (2026-09-24, Cél site): ✅

- **Q1 – `addAvailableContentType` REST-ből: működik.** A `POST <lista>/contenttypes/addAvailableContentType('0x0104')` 200-at adott, és a lista-szintű ID-t adta vissza: `0x0104007B48E5823A4D6E4C8C5BF88502FD1CDA` (site-ID + `00` + GUID).
- **Q2a – alapértelmezett tartalomtípus REST-ből: működik.** A `MERGE <lista>/RootFolder` hívás `odata=verbose` törzzsel (`UniqueContentTypeOrder`: `Collection(SP.ContentTypeId)`) 204-et adott.
  - **Előtte:** `ContentTypeOrder = [Elem, Hirdetmény]`, `UniqueContentTypeOrder = null`.
  - **Utána:** mindkettő `[Hirdetmény, Elem]`, vagyis a Hirdetmény lett az alapértelmezett. A CSOM-ra (Q2b) nem volt szükség.

→ **Döntés:**
- A `ListProvider` REST-tel rendeli a listához a tartalomtípusokat (`addAvailableContentType`), és REST-tel állítja a sorrendet.
- A sorrend-beállítás közvetlen `fetch`-csel megy (`core/http/raw.ts`), `odata=verbose` fejlécekkel. Ennek oka, hogy a PnPjs alapfejlécei a `Content-Type`-ot `application/json;charset=utf-8`-ra írják felül, a sorrend-beállítást pedig csak `verbose` formátumban teszteltük.
- A sorrendből a provider soha nem vesz ki meglévő tartalomtípust: a sablon típusai kerülnek előre, a cél többi látható típusa utánuk marad.

**Takarítás:** a Cél site-on a `CopyJet Spike B2 ékezetes` lista (Hirdetmény tartalomtípussal és a `CopyJetMappa1` mappával) kézzel törölhető.
