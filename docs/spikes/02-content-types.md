# Spike 02 – Tartalomtípusok: egyedi vagy beépített, ID megtartása, mezőhivatkozás REST-ből

Állapot: **lezárva** (2026-09-24) · Érintett kód: `src/core/contentTypes/`, `src/core/http/csom*.ts`, `ContentTypeExtractor`, `ContentTypeProvider`

## Összefoglaló

| Kérdés | Válasz | Megoldás |
| --- | --- | --- |
| Egyedi vagy beépített? | A csoport és az ID alakja nem, a `FeatureId` a `SchemaXml`-ben igen, kivéve a modern lapok családját | `isCustomContentType`: nem rejtett, nem `_Hidden`, nincs `FeatureId`, nincs a 7 elemű pontos ID-listán |
| Megtartja-e a REST az ID-t? | **Nem**, sem nometadata, sem verbose formátumban; a szülőt sem | **CSOM** `ContentTypes.Add(ContentTypeCreationInformation)` |
| Nem létező tartalomtípus lekérdezése | HTTP 200 + `{"odata.null": true}`, nem 404 | `_get` a `StringId` meglétét vizsgálja |
| Mezőhivatkozás REST-ből | **Nem működik**, létező oszlopra is „nem létezik” hibát ad | **CSOM** `FieldLinks.Add(FieldLinkCreationInformation)` |
| Hivatkozás jelzői REST-ből | **Nem** (`SP.FieldLink` nem támogatja a PATCH-et) | **CSOM** `SetProperty Required/Hidden` + `ContentType.Update(false)` |

## Kérdések

1. **Egyedi vagy beépített?** Mi különbözteti meg a felhasználó által létrehozott tartalomtípust a beépítettektől?
   - A mezőknél a `CanBeDeleted` és a csoport sem volt jó ismérv (spike 01), itt is hasonlóra számítunk.
   - **Ideiglenes feltevés:** a beépítettek `SchemaXml`-jében van `FeatureId`, az egyedieknél nincs.
   - További jelöltek: `Group`, `Hidden`, `ReadOnly`, `Sealed`, az ID alakja (szülő + `00` + 32 hex).
2. **Megmarad-e az ID?** A `POST /_api/web/contenttypes` hívás `{ "Id": { "StringValue": "0x0100…" } }` törzzsel pontosan ezt az ID-t adja-e az új tartalomtípusnak? Az öröklés és a listák tartalomtípus-hivatkozásai ezen múlnak.
3. **Mezőhivatkozás REST-ből:** a `POST /_api/web/contenttypes('<id>')/fieldlinks` hívás `{ "FieldInternalName": "…", "Required": false, "Hidden": false }` törzzsel hozzáadja-e a mezőt? A PnPjs 4-ben erre nincs metódus.
4. **Meglévő mezőhivatkozás jelzői:** módosítható-e egy meglévő (akár örökölt) mezőhivatkozás `Required`/`Hidden` jelzője REST-ből (`MERGE` a `fieldlinks('<guid>')`-re)?
   - Addig a provider ezt nem próbálja meg, csak figyelmeztet (`CT_FIELDLINK_FLAGS`). Ez akkor fordul elő, ha például egy gyerek tartalomtípus rejtetté teszi az örökölt Title-t.

## A. kísérlet – Forrás site (csak olvas)

A futtatás módja ugyanaz, mint a spike 01-nél: F12 → Console, beillesztés, Enter. Ha a Chrome elsőre letiltja a beillesztést, írd be: `allow pasting`. A konzolba kiírt JSON-t küldd vissza.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const r = await fetch(web + '/_api/web/contenttypes?$select=StringId,Name,Group,Hidden,ReadOnly,Sealed,SchemaXml', { headers: H });
  const cts = (await r.json()).value || [];
  const rows = cts.map((ct) => {
    const el = new DOMParser().parseFromString(ct.SchemaXml, 'application/xml').documentElement;
    return {
      id: ct.StringId, name: ct.Name, group: ct.Group, hidden: ct.Hidden, readOnly: ct.ReadOnly, sealed: ct.Sealed,
      featureId: el.getAttribute('FeatureId') || null,
      guidSuffix: /00[0-9A-F]{32}$/i.test(ct.StringId),
      attrs: Array.from(el.attributes).map((a) => a.name).join(',')
    };
  });
  const json = JSON.stringify({ web, status: r.status, count: rows.length, rows }, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

Érdemes előtte a Forrás site-on létrehozni 1–2 saját tartalomtípust (*Webhelybeállítások → Webhely-tartalomtípusok → Létrehozás*), például egy „Projekt” típust „Elem” szülővel és egy saját oszloppal. Így a kimenetben lesz egyedi minta is.

## B. kísérlet – Cél site (**ír**: létrehoz egy teszt tartalomtípust)

Csak a Cél **teszt** site-on futtasd. A kód a következőket hozza létre:
- a `CopyJet Spike CT` nevű tartalomtípust `0x0100C0FFEE00C0FFEE00C0FFEE00C0FFEE00` ID-val, „Elem” szülővel;
- ehhez hozzáadja a beépített `Priority` (Prioritás) oszlopot mezőhivatkozásként.

Semmit nem töröl. A tartalomtípust a teszt után a *Webhely-tartalomtípusok* oldalon kézzel törölheted. Ha már létezik, a kód csak a meglévőt kérdezi le.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const ID = '0x0100C0FFEE00C0FFEE00C0FFEE00C0FFEE00';
  const out = { web };

  // A missing content type comes back as 200 {"odata.null": true}, not 404.
  const existing = await (await fetch(`${web}/_api/web/contenttypes('${ID}')?$select=StringId,Name`, { headers: H })).json();
  if (!existing.StringId) {
    const add = await fetch(web + '/_api/web/contenttypes', { method: 'POST', headers: W,
      body: JSON.stringify({ Id: { StringValue: ID }, Name: 'CopyJet Spike CT', Group: 'CopyJet spike', Description: 'Törölhető teszt' }) });
    out.q2_addStatus = add.status;
    out.q2_addBody = add.ok ? await add.json() : await add.text();
  } else {
    out.q2_existing = existing;
  }
  const check = await fetch(`${web}/_api/web/contenttypes('${ID}')?$select=StringId,Name`, { headers: H });
  out.q2_getById = { status: check.status, body: check.ok ? await check.json() : await check.text() };

  const link = await fetch(`${web}/_api/web/contenttypes('${ID}')/fieldlinks`, { method: 'POST', headers: W,
    body: JSON.stringify({ FieldInternalName: 'Priority', Required: false, Hidden: false }) });
  out.q3_linkStatus = link.status;
  out.q3_linkBody = link.ok ? await link.json() : await link.text();
  const links = await fetch(`${web}/_api/web/contenttypes('${ID}')/fieldlinks?$select=Id,Name,FieldInternalName,Required,Hidden`, { headers: H });
  const linkRows = links.ok ? (await links.json()).value : [];
  out.q3_links = links.ok ? linkRows.map((l) => l.FieldInternalName || l.Name) : links.status;

  // Q4: try to make the inherited Title link required via MERGE, then read it back.
  const title = linkRows.find((l) => (l.FieldInternalName || l.Name) === 'Title');
  if (title) {
    const merge = await fetch(`${web}/_api/web/contenttypes('${ID}')/fieldlinks('${title.Id}')`, { method: 'POST',
      headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ Required: !title.Required }) });
    out.q4_mergeStatus = merge.status;
    out.q4_mergeBody = merge.ok ? '' : await merge.text();
    const after = await fetch(`${web}/_api/web/contenttypes('${ID}')/fieldlinks('${title.Id}')?$select=Required`, { headers: H });
    out.q4_requiredBefore = title.Required;
    out.q4_requiredAfter = after.ok ? (await after.json()).Required : after.status;
  }

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## Mit várunk?

- **A:** legyen egy olyan ismérv (vagy kettő kombinációja), amely az egyedi tartalomtípusokat pontosan elválasztja a beépítettektől.
- **B:**
  - `q2_getById.body.StringId` pontosan `0x0100C0FFEE00C0FFEE00C0FFEE00C0FFEE00`. Ha nem, az ID megtartásához más út kell, például a `CreateContentType` hívás vagy a CSOM `ContentTypeCreationInformation.Id`.
  - `q3_linkStatus` 200 vagy 201, és a `q3_links` listában szerepel a `Priority`.
  - `q4_requiredAfter` ≠ `q4_requiredBefore`: ekkor a jelzők REST-ből állíthatók, és a provider beállíthatja őket figyelmeztetés helyett.

## Eredmény

### A – egyedi vagy beépített (2026-09-24, Forrás site, magyar nyelvű)

A web 66 tartalomtípusa közül 58-nak van `FeatureId`-ja. Ezek mind beépítettek, és egyiknél sem volt `FeatureId`, ami egyedi lett volna. A **8 `FeatureId` nélküli**:

| ID | Név | Csoport | Valójában |
| --- | --- | --- | --- |
| `0x0100573FB3…2051` | TestSiteContentType | Egyéni tartalomtípusok | **egyedi** |
| `0x0101009D1CB255DA76424F860D91F20E6C4118` | Webhelylap | Dokumentumtartalom-típusok | beépített (modern lap) |
| `…4118002A50BFCFB7614729B56886FADA02339B` | Újraközzétételi lap | Dokumentumtartalom-típusok | beépített |
| `…411800EE3B4FE2C29A4DCDB55229ADEB3DBFE4` | HTML-oldal | Dokumentumtartalom-típusok | beépített |
| `…411800C57E7DC123744954AACC08D1D5260154` | Tartalom frissessége | `_Hidden`, rejtett | beépített |
| `0x0101002039C03B61C64EC4A04F5361F3851066` | Megjelenítési sablon | `_Hidden` | beépített |
| `…385106605` | Megjelenítési sablon kódja | `_Hidden` | beépített |
| `0x010100C5033D6CFB8447359FB795C8A73A2B19` | Tervezési fájl | `_Hidden` | beépített |

A további vizsgált jelöltekből ez derült ki:
- **Az ID alakja nem jó ismérv.** Rengeteg beépített tartalomtípus ID-ja is szülő + `00` + GUID alakú (Körözvény, Foglalások, Webhelylap).
- **A csoport nem jó ismérv.** A beépítettek egy része nem `_Hidden` csoportú, például a Csoportmunka- és a Dokumentumtartalom-típusok.
- **A `Sealed` és a `ReadOnly` nem jó ismérv.**

→ **Döntés (`isCustomContentType`):** egyedi az a tartalomtípus, amelyre egyszerre teljesül:
- nem rejtett, és a csoportja nem `_Hidden`;
- a `SchemaXml`-jében nincs `FeatureId`;
- nincs rajta a `FeatureId` nélküli beépítettek pontos ID-listáján (a fenti 7 beépített).

A lista **pontos ID-kat** tartalmaz, nem ágat. Így a Webhelylap egyedi gyereke, vagyis egy saját lap-tartalomtípus, egyedinek számít. Ha a SharePoint új, `FeatureId` nélküli beépített típust vezet be, az a Setupban egyediként jelenik meg. A célon azonban „már létezik” (`same`) diffet ad, tehát nem okoz kárt. Ilyenkor a listát bővíteni kell.

### B – ID megtartása, mezőhivatkozás, jelzők

**Első futás (2026-09-24, Cél site):** a kísérleti kód hibás volt, a tartalomtípus nem jött létre. Közben kiderült egy fontos viselkedés:
- **A nem létező tartalomtípus nem 404-et ad.** A `GET contenttypes('<nem létező ID>')` HTTP 200-zal és `{"odata.null": true}` törzzsel válaszol, a kód viszont 404-et várt.
- A `ContentTypeProvider._get` a `StringId` meglétét is ellenőrzi, így erre felkészült. A teszt-mock mostantól ezt a valódi választ adja.
- Egy nem létező tartalomtípus `fieldlinks` gyűjteménye viszont már 404-et ad (`ResourceNotFoundException`).

**Második futás (2026-09-24, Cél site): a REST nem tartja meg az ID-t.**
- A `POST /_api/web/contenttypes` hívás `odata=nometadata` törzzsel és `Id: { StringValue: "0x0100C0FFEE00C0FFEE00C0FFEE00C0FFEE00" }` értékkel 201-et adott, de a tartalomtípus **saját ID-t kapott**: `0x0100827E5C29C68D314D9295EA993AE5BD59`. A kért ID-ra végzett lekérdezés `odata.null`.
- **A szülő helyes maradt** (`0x01` + `00` + új GUID), vagyis a SharePoint a kért ID-ból a szülőt kiolvasta, csak a saját részt generálta újra.
- A REST-tel létrehozott tartalomtípus `SchemaXml`-jében `FeatureIds` (többes szám) attribútum van, `FeatureId` nincs. Az `isCustomContentType` szabálya csak a `FeatureId="`-re illeszkedik, így ez nem zavarja.
- A mezőhivatkozás (Q3) és a jelzők (Q4) kérdése nyitva maradt, mert a kód a kért ID-n kereste a tartalomtípust.

→ A `ContentTypeProvider` jelenleg előbb létrehozza a rossz ID-jú típust, és csak utána dob `CT_ID_NOT_KEPT` hibát. Ezt a C kísérlet eredménye alapján javítani kell.

## C. kísérlet – Cél site (**ír**): ID-megtartó létrehozási utak, mezőhivatkozás

Csak a Cél **teszt** site-on futtasd. A kód legfeljebb 3 új tartalomtípust hoz létre (`CopyJet Spike verbose`, `CopyJet Spike CSOM`, `CopyJet Spike dokumentum`). Ezen felül a B-ben már létrejött `CopyJet Spike CT`-hez hozzáadja a Priority oszlopot, és átállítja rajta a Title hivatkozás `Required` jelzőjét. Semmit nem töröl.

- **C1:** REST, `odata=verbose` formátum, típusjelöléssel (`SP.ContentType` / `SP.ContentTypeId`).
- **C2:** CSOM (`/_vti_bin/client.svc/ProcessQuery`), `ContentTypeCreationInformation.Id`. A CSOM típusazonosítókat emlékezetből írtam. Ha hibásak, a szerver hibaüzenete ezt megmutatja.
- **C3:** REST (nometadata), dokumentum szülő (`0x0101` + `00` + GUID). Azt ellenőrzi, hogy a szülőt a kért ID-ból veszi-e a SharePoint.
- **Q3/Q4:** a mezőhivatkozás és a jelzők a B-ben létrejött `0x0100827E5C29C68D314D9295EA993AE5BD59` tartalomtípuson.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const getCt = async (id) => (await (await fetch(`${web}/_api/web/contenttypes('${id}')?$select=StringId,Name`, { headers: H })).json()).StringId || null;
  const text = async (r) => { try { return await r.text(); } catch { return ''; } };
  const out = { web };

  // C1: odata=verbose with type annotations
  const ID1 = '0x0100C0FFEE01C0FFEE01C0FFEE01C0FFEE01';
  if (!(await getCt(ID1))) {
    const r = await fetch(web + '/_api/web/contenttypes', { method: 'POST',
      headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest },
      body: JSON.stringify({ __metadata: { type: 'SP.ContentType' }, Id: { __metadata: { type: 'SP.ContentTypeId' }, StringValue: ID1 },
        Name: 'CopyJet Spike verbose', Group: 'CopyJet spike', Description: 'Törölhető teszt' }) });
    const body = await text(r);
    out.c1_status = r.status;
    out.c1_createdId = (body.match(/"StringId":"([^"]+)"/) || [])[1] || body.slice(0, 400);
  }
  out.c1_requestedIdExists = !!(await getCt(ID1));

  // C2: CSOM ContentTypeCreationInformation.Id
  const ID2 = '0x0100C0FFEE02C0FFEE02C0FFEE02C0FFEE02';
  if (!(await getCt(ID2))) {
    const xml = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CopyJet">` +
      `<Actions><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="StringId" ScalarProperty="true" /></Properties></Query></Query></Actions>` +
      `<ObjectPaths><StaticProperty Id="5" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="6" ParentId="5" Name="Web" /><Property Id="8" ParentId="6" Name="ContentTypes" />` +
      `<Method Id="9" ParentId="8" Name="Add"><Parameters><Parameter TypeId="{168f3091-4554-4f14-8866-b20d48e45b54}">` +
      `<Property Name="Description" Type="String">Törölhető teszt</Property><Property Name="Group" Type="String">CopyJet spike</Property>` +
      `<Property Name="Id" Type="String">${ID2}</Property><Property Name="Name" Type="String">CopyJet Spike CSOM</Property><Property Name="ParentContentType" Type="Null" />` +
      `</Parameter></Parameters></Method></ObjectPaths></Request>`;
    const r = await fetch(web + '/_vti_bin/client.svc/ProcessQuery', { method: 'POST', headers: { 'Content-Type': 'text/xml', 'X-RequestDigest': digest }, body: xml });
    out.c2_status = r.status;
    out.c2_body = (await text(r)).slice(0, 600);
  }
  out.c2_requestedIdExists = !!(await getCt(ID2));

  // C3: nometadata REST, Document parent – is the parent taken from the requested ID?
  const r3 = await fetch(web + '/_api/web/contenttypes', { method: 'POST', headers: W,
    body: JSON.stringify({ Id: { StringValue: '0x010100C0FFEE03C0FFEE03C0FFEE03C0FFEE03' }, Name: 'CopyJet Spike dokumentum', Group: 'CopyJet spike', Description: 'Törölhető teszt' }) });
  const b3 = await text(r3);
  out.c3_status = r3.status;
  out.c3_createdId = (b3.match(/"StringId":"([^"]+)"/) || [])[1] || b3.slice(0, 300);

  // Q3/Q4 on the content type created in run B
  const B = '0x0100827E5C29C68D314D9295EA993AE5BD59';
  const link = await fetch(`${web}/_api/web/contenttypes('${B}')/fieldlinks`, { method: 'POST', headers: W,
    body: JSON.stringify({ FieldInternalName: 'Priority', Required: false, Hidden: false }) });
  out.q3_linkStatus = link.status;
  out.q3_linkBody = link.ok ? '' : (await text(link)).slice(0, 400);
  const links = await fetch(`${web}/_api/web/contenttypes('${B}')/fieldlinks?$select=Id,Name,FieldInternalName,Required,Hidden`, { headers: H });
  const rows = links.ok ? (await links.json()).value : [];
  out.q3_links = rows.map((l) => `${l.FieldInternalName || l.Name}${l.Required ? ' (kötelező)' : ''}${l.Hidden ? ' (rejtett)' : ''}`);
  const title = rows.find((l) => (l.FieldInternalName || l.Name) === 'Title');
  if (title) {
    const m = await fetch(`${web}/_api/web/contenttypes('${B}')/fieldlinks('${title.Id}')`, { method: 'POST',
      headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ Required: !title.Required }) });
    out.q4_mergeStatus = m.status;
    out.q4_mergeBody = m.ok ? '' : (await text(m)).slice(0, 400);
    const after = await fetch(`${web}/_api/web/contenttypes('${B}')/fieldlinks('${title.Id}')?$select=Required`, { headers: H });
    out.q4_requiredBefore = title.Required;
    out.q4_requiredAfter = after.ok ? (await after.json()).Required : after.status;
  }

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- `c1_requestedIdExists` vagy `c2_requestedIdExists` értéke `true`: ekkor ezen az úton megmarad az ID, és a provider ezt használja.
- Ha mindkettő `false`: a provider új ID-val hoz létre, és az ID-t tokenként (`{ctid:X}`) viszi tovább. Ehhez a `c3_createdId`-nek `0x010100`-val kell kezdődnie, vagyis a szülőt a kért ID-ból kell vennie.
- `q3_linkStatus` 200 vagy 201, és a `q3_links` listában szerepel a Priority. `q4_requiredAfter` ≠ `q4_requiredBefore`.

### C – eredmény (2026-09-24, Cél site)

A konzolkimenet csonkán érkezett, ezért a C1 kimenetét egy csak olvasó lekérdezés egészítette ki:

| Út | Kért ID | Kapott ID | Eredmény |
| --- | --- | --- | --- |
| C1 – REST `odata=verbose` | `0x0100C0FFEE01…` | `0x010090359CBB24D86C4FBB4026BFD483073B` | ❌ az ID elveszett |
| **C2 – CSOM `ContentTypeCreationInformation.Id`** | `0x0100C0FFEE02…` | `0x0100C0FFEE02C0FFEE02C0FFEE02C0FFEE02` | ✅ **az ID megmaradt** |
| C3 – REST, dokumentum szülő | `0x010100C0FFEE03…` | `0x0100CCB62FB272FCED4EB47621B1815FE980` | ❌ az ID **és a szülő** is elveszett (Elem lett) |

A B-ben létrejött tartalomtípusra a REST-es mezőhivatkozás-hozzáadás (nometadata) nem működött: a Priority nem került fel, a hivatkozások csak `ContentType, Title (kötelező)`. A pontos hibaüzenet a csonkolás miatt nem látszott.

→ **Döntés:** a tartalomtípus **CSOM-mal** jön létre (`core/http/csom.ts` → `csomCreateContentType`). A CSOM típusazonosítói a C2 alapján helyesek (`RequestContext` `{3747adcd-…}`, `ContentTypeCreationInformation` `{168f3091-…}`). A provider létrehozás után ellenőrzi az ID-t (`CT_ID_NOT_KEPT`). Az ID-t tokenként továbbvivő kerülőút a C3 miatt nem járható, mert a REST a szülőt sem tartja meg.

## D. kísérlet – Cél site (**ír**): mezőhivatkozás és jelzők

Csak a Cél **teszt** site-on futtasd. Új tartalomtípust nem hoz létre. A B-ben létrejött `CopyJet Spike CT`-hez (`0x0100827E5C29C68D314D9295EA993AE5BD59`) megpróbálja hozzáadni a Priority oszlopot, és átállítani a Title `Required` jelzőjét. A `sp.js`-t csak olvassa.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const CT = `${web}/_api/web/contenttypes('0x0100827E5C29C68D314D9295EA993AE5BD59')`;
  const short = async (r) => { try { return (await r.text()).slice(0, 300); } catch { return ''; } };
  const linkNames = async () => { const r = await fetch(`${CT}/fieldlinks?$select=Id,Name,FieldInternalName,Required,Hidden`, { headers: H }); return r.ok ? (await r.json()).value : []; };
  const out = { web };

  // D1: REST nometadata
  let r = await fetch(`${CT}/fieldlinks`, { method: 'POST', headers: W, body: JSON.stringify({ FieldInternalName: 'Priority', Required: false, Hidden: false }) });
  out.d1_status = r.status; out.d1_body = await short(r);
  out.d1_hasPriority = (await linkNames()).some((l) => (l.FieldInternalName || l.Name) === 'Priority');

  // D2: REST verbose with type annotation
  if (!out.d1_hasPriority) {
    r = await fetch(`${CT}/fieldlinks`, { method: 'POST',
      headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest },
      body: JSON.stringify({ __metadata: { type: 'SP.FieldLink' }, FieldInternalName: 'Priority', Required: false, Hidden: false }) });
    out.d2_status = r.status; out.d2_body = await short(r);
    out.d2_hasPriority = (await linkNames()).some((l) => (l.FieldInternalName || l.Name) === 'Priority');
  }

  // D3: MERGE Required on the existing Title link
  const title = (await linkNames()).find((l) => (l.FieldInternalName || l.Name) === 'Title');
  if (title) {
    r = await fetch(`${CT}/fieldlinks('${title.Id}')`, { method: 'POST', headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ Required: !title.Required }) });
    out.d3_status = r.status; out.d3_body = await short(r);
    out.d3_requiredBefore = title.Required;
    out.d3_requiredAfter = ((await linkNames()).find((l) => l.Id === title.Id) || {}).Required;
  }
  out.links = (await linkNames()).map((l) => `${l.FieldInternalName || l.Name}${l.Required ? ' (kötelező)' : ''}${l.Hidden ? ' (rejtett)' : ''}`);

  // D4: CSOM type IDs from the site's own sp.js (read-only)
  for (const file of ['sp.js', 'sp.debug.js']) {
    const js = await (await fetch(`${web}/_layouts/15/${file}`)).text();
    const typeIdOf = (name) => { const m = new RegExp(name + '[\\s\\S]{0,4000}?typeId[\\s\\S]{0,120}?\\{([0-9a-fA-F-]{36})\\}').exec(js); return m ? m[1] : null; };
    out[`d4_${file}`] = {
      size: js.length,
      FieldLinkCreationInformation: typeIdOf('FieldLinkCreationInformation'),
      ContentTypeCreationInformation: typeIdOf('ContentTypeCreationInformation')
    };
  }

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- `d1_hasPriority` vagy `d2_hasPriority` értéke `true`: ekkor a mezőhivatkozás REST-tel hozzáadható, és a provider azt a formátumot használja.
- Ha egyik sem: CSOM `FieldLinks.Add` kell. Ehhez a D4 adja a `FieldLinkCreationInformation` típusazonosítóját. A D4 `ContentTypeCreationInformation` értékének egyeznie kell a C2-ben bevált `168f3091-4554-4f14-8866-b20d48e45b54`-gyel; ez igazolja, hogy a kiolvasás jó.
- A `d3_requiredAfter` ≠ `d3_requiredBefore`: ekkor a jelzők REST-tel állíthatók.

### D – eredmény (2026-09-24, Cél site)

- **D1/D2:** mindkét REST-formátum (nometadata és verbose) HTTP 400-at adott ezzel az üzenettel: „Nem létezik a(z) 'Priority' oszlop.” Maga a hívás tehát elfogadott formátumú, és a SharePoint kikereste a mezőt. **A tesztoszlopot választottam rosszul:** a Priority a Forrás site-on létezett, a Cél site-on nem. Ez tervezési hiba volt a kísérletben.
  → Ha a mezőhivatkozás hiányzó site column miatt nem jön létre, a provider `CT_FIELDLINK_FAILED` hibát ad, és a mögöttes SharePoint-hiba a `detail`-ben marad. Ez a viselkedés helyes.
- **D3:** `MERGE` a `fieldlinks('<id>')`-re → HTTP 400, „A(z) SP.FieldLink típus nem támogatja a(z) PATCH HTTP-metódust.”
  → Egy meglévő mezőhivatkozás `Required`/`Hidden` jelzője **REST-ből nem módosítható**, csak CSOM-mal (`FieldLink.Required` + `ContentType.Update(false)`). Addig marad a `CT_FIELDLINK_FLAGS` figyelmeztetés.
- **D4:** a site saját `sp.js`-éből kiolvasott CSOM-típusazonosítók: `FieldLinkCreationInformation` = `{63fb2c92-8f65-4bbb-a658-b6cd294403f4}`, `ContentTypeCreationInformation` = `{168f3091-4554-4f14-8866-b20d48e45b54}`. Az utóbbi egyezik a C2-ben bevált értékkel, tehát a kiolvasás megbízható. Az `sp.js` és az `sp.debug.js` ugyanazt adta.

## E. kísérlet – Cél site (**ír**): mezőhivatkozás létező oszloppal

Csak a Cél **teszt** site-on futtasd. Új tartalomtípust nem hoz létre. A kód a `CopyJet Spike CT`-hez hozzáadja az első olyan oszlopot, amely a Cél site-on létezik, és még nincs a tartalomtípuson. Ezt **kötelezőként** (`Required: true`) adja hozzá, majd visszaolvassa.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const CT = `${web}/_api/web/contenttypes('0x0100827E5C29C68D314D9295EA993AE5BD59')`;
  const links = async () => (await (await fetch(`${CT}/fieldlinks?$select=Name,FieldInternalName,Required,Hidden`, { headers: H })).json()).value || [];
  const out = { web };

  const present = (await links()).map((l) => l.FieldInternalName || l.Name);
  let field = null;
  for (const name of ['SiteColumn1', 'Comments', '_Comments', 'Categories', 'Keywords', 'Location', 'StartDate', 'URL']) {
    if (present.includes(name)) continue;
    const r = await fetch(`${web}/_api/web/availablefields?$filter=InternalName eq '${name}'&$select=InternalName`, { headers: H });
    if (r.ok && ((await r.json()).value || []).length) { field = name; break; }
  }
  out.field = field;
  if (field) {
    const r = await fetch(`${CT}/fieldlinks`, { method: 'POST', headers: W, body: JSON.stringify({ FieldInternalName: field, Required: true, Hidden: false }) });
    out.e1_status = r.status;
    out.e1_body = r.ok ? '' : (await r.text()).slice(0, 300);
    const link = (await links()).find((l) => (l.FieldInternalName || l.Name) === field);
    out.e1_added = !!link;
    out.e1_required = link ? link.Required : null;
  }
  out.links = (await links()).map((l) => `${l.FieldInternalName || l.Name}${l.Required ? ' (kötelező)' : ''}${l.Hidden ? ' (rejtett)' : ''}`);

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- `e1_added` = `true`: ekkor a mezőhivatkozás REST-tel hozzáadható, és a provider jelenlegi megoldása jó.
- `e1_required` = `true`: ekkor új hivatkozásnál a `Required` jelző megmarad, és csak a meglévő, örökölt hivatkozások jelzőihez kellene CSOM.

### E – eredmény (2026-09-24, Cél site)

- A `Comments` oszlop **létezik** a Cél site-on (az `availablefields?$filter=InternalName eq 'Comments'` megtalálta), a `POST …/fieldlinks` mégis HTTP 400-at adott: „Nem létezik a(z) 'Comments' oszlop.”
  → **A REST-es mezőhivatkozás-hozzáadás nem működik**, a D1/D2 hibája sem a rossz tesztoszlop miatt volt. A D eredményénél leírt „a hívás elfogadott formátumú” feltevés téves volt.

→ **Döntés:** a mezőhivatkozásokat is **CSOM**-mal kezeljük. A művelet `ContentType.FieldLinks.Add(FieldLinkCreationInformation { Field = web.AvailableFields.GetByInternalNameOrTitle(név) })`, majd a jelzők `SetProperty` beállítása és `ContentType.Update(false)`. Ez egy kérésben kezeli:
- az új hivatkozásokat;
- az új és a meglévő (örökölt) hivatkozások `Required`/`Hidden` jelzőit, amit a REST nem tud (D3).

Az XML-t a `src/core/http/csomXml.ts` (`fieldLinksBody`) állítja elő. A `CT_FIELDLINK_FLAGS` figyelmeztetés megszűnt, a jelzőket a provider beállítja.

## F. kísérlet – Cél site (**ír**): CSOM mezőhivatkozás és jelzők

Csak a Cél **teszt** site-on futtasd. Új tartalomtípust nem hoz létre. A `CopyJet Spike CT`-n **pontosan azt a kérést** küldi el, amelyet a provider küldene:
- a `Comments` oszlopot hozzáadja kötelezőként;
- a meglévő Title hivatkozást nem kötelezővé teszi.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const CT = `${web}/_api/web/contenttypes('0x0100827E5C29C68D314D9295EA993AE5BD59')`;
  const links = async () => (await (await fetch(`${CT}/fieldlinks?$select=Id,Name,FieldInternalName,Required,Hidden`, { headers: H })).json()).value || [];
  const show = (ls) => ls.map((l) => `${l.FieldInternalName || l.Name}${l.Required ? ' (kötelező)' : ''}${l.Hidden ? ' (rejtett)' : ''}`);
  const before = await links();
  const title = before.find((l) => (l.FieldInternalName || l.Name) === 'Title');
  // Exactly what ContentTypeProvider sends (generated by src/core/http/csomXml.ts fieldLinksBody):
  const xml = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CopyJet"><Actions><ObjectPath Id="10" ObjectPathId="4" /><ObjectPath Id="11" ObjectPathId="100" /><ObjectPath Id="12" ObjectPathId="101" /><SetProperty Id="13" ObjectPathId="101" Name="Required"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="14" ObjectPathId="101" Name="Hidden"><Parameter Type="Boolean">false</Parameter></SetProperty><ObjectPath Id="15" ObjectPathId="102" /><SetProperty Id="16" ObjectPathId="102" Name="Required"><Parameter Type="Boolean">false</Parameter></SetProperty><SetProperty Id="17" ObjectPathId="102" Name="Hidden"><Parameter Type="Boolean">false</Parameter></SetProperty><Method Name="Update" Id="18" ObjectPathId="4"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method></Actions><ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="1" Name="Web" /><Property Id="3" ParentId="2" Name="ContentTypes" /><Method Id="4" ParentId="3" Name="GetById"><Parameters><Parameter Type="String">0x0100827E5C29C68D314D9295EA993AE5BD59</Parameter></Parameters></Method><Property Id="5" ParentId="4" Name="FieldLinks" /><Property Id="6" ParentId="2" Name="AvailableFields" /><Method Id="100" ParentId="6" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">Comments</Parameter></Parameters></Method><Method Id="101" ParentId="5" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="100" /></Parameter></Parameters></Method><Method Id="102" ParentId="5" Name="GetById"><Parameters><Parameter Type="Guid">{${title.Id}}</Parameter></Parameters></Method></ObjectPaths></Request>`;
  const r = await fetch(web + '/_vti_bin/client.svc/ProcessQuery', { method: 'POST', headers: { 'Content-Type': 'text/xml', Accept: 'application/json', 'X-RequestDigest': digest }, body: xml });
  const body = await r.text();
  const out = { web, before: show(before), f_status: r.status, f_error: (body.match(/"ErrorMessage":"([^"]*)"/) || [])[1] || null, after: show(await links()) };
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- `f_status` = 200, `f_error` = `null`;
- az `after` listában: `Title` (kötelező jelző nélkül) és `Comments (kötelező)`.

### F – eredmény (2026-09-24, Cél site): ✅

A `fieldLinksBody` által előállított kérés, amelyet a provider is küld, HTTP 200-at adott, `ErrorInfo` nélkül:
- **előtte:** `ContentType`, `Title (kötelező)`;
- **utána:** `ContentType`, `Title`, `Comments (kötelező)`.

Egy kérésben működött az új hivatkozás hozzáadása a `Required` jelzővel együtt, és a meglévő (örökölt) Title hivatkozás jelzőjének módosítása is.

**Takarítás:** a Cél site-on maradt tesztelemeket (`CopyJet Spike CT`, `CopyJet Spike verbose`, `CopyJet Spike CSOM`, `CopyJet Spike dokumentum`, csoport: „CopyJet spike”) a *Webhely-tartalomtípusok* oldalon kézzel kell törölni. A CopyJet maga soha nem töröl.
