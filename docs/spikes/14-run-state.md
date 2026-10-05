# Spike 14 – Futásállapot a cél site-on (`CopyJetLog`)

Állapot: **lezárva** · Érintett kód (4. fázis): `core/state`, `engine` (folytatás), `ResumeBanner`

## Kérdések

1. **Rejtett lista:** létrehozható-e REST-tel `Hidden: true` és `NoCrawl: true` beállítással egy egyéni lista (`BaseTemplate 100`) a `Lists/CopyJetLog` címen? Látszik-e utána a Webhely tartalmában?
2. **Hosszú szöveg:** mekkora JSON fér egy több soros, sima szöveges (`Note`, `RichText="FALSE"`) mezőbe? A kísérlet 10 000, 70 000 és 300 000 karaktert ír, és visszaolvassa a hosszt.
3. **Melléklet felülírása:** egy elem mellékletét (`state.json`) lehet-e helyben felülírni (`AttachmentFiles('…')/$value`, `PUT`), vagy törölni és újra felvenni kell?
4. **Csak tulajdonosoknak:** `breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)` után ki marad a listán? Megmarad-e a futtató felhasználó? Felvehető-e utána a tulajdonoscsoport teljes hozzáféréssel?

A döntés tétje: a futás állapota (kész lépések, tokenek, elem-ID-leképezések) a mezőben vagy a mellékletben legyen-e. A leképezés nagy lista esetén több százezer karakter is lehet.

## Kísérlet – Cél site (ír, majd töröl)

A kísérlet a Cél site-on létrehoz egy **„CopyJetSpike14”** rejtett listát, ír bele egy elemet mezővel és melléklettel, felülírja a mellékletet, majd megszünteti a lista jogosultság-öröklését. A végén a listát a lomtárba teszi. Ha meg szeretnéd nézni, írd át a `CLEANUP` értékét `false`-ra; ekkor utána kézzel kell törölnöd (Webhely tartalma → `/_layouts/15/viewlsts.aspx` nem mutatja, mert rejtett; a lista címe: `…/Lists/CopyJetSpike14`).

```js
(async () => {
  const CLEANUP = true;
  const site = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const J = 'application/json;odata=nometadata';
  const digest = (await (await fetch(site + '/_api/contextinfo', { method: 'POST', headers: { Accept: J } })).json()).FormDigestValue;
  const call = async (url, method = 'GET', body, extra = {}) => {
    const headers = { Accept: J, 'X-RequestDigest': digest, ...extra };
    if (body !== undefined && typeof body !== 'string' && !(body instanceof Blob)) { headers['Content-Type'] = J; body = JSON.stringify(body); }
    const r = await fetch(site + '/_api/' + url, { method: method === 'GET' ? 'GET' : 'POST', headers: method === 'MERGE' || method === 'DELETE' || method === 'PUT' ? { ...headers, 'X-HTTP-Method': method, 'IF-MATCH': '*' } : headers, body });
    const text = await r.text();
    if (!r.ok) return { httpStatus: r.status, error: text.slice(0, 300) };
    try { return text ? JSON.parse(text) : {}; } catch { return text; }
  };
  const out = {};
  const name = 'CopyJetSpike14';
  // 1. Hidden custom list.
  const list = await call('web/lists', 'POST', { Title: name, BaseTemplate: 100, Hidden: true, NoCrawl: true, Description: 'CopyJet spike 14' });
  if (list.httpStatus) { console.log(JSON.stringify({ createList: list }, null, 2)); return; }
  const L = `web/lists(guid'${list.Id}')`;
  const info = await call(`${L}?$select=Hidden,NoCrawl,RootFolder/ServerRelativeUrl&$expand=RootFolder`);
  out.list = { hidden: info.Hidden, noCrawl: info.NoCrawl, url: info.RootFolder && info.RootFolder.ServerRelativeUrl };
  const visible = await call(`web/lists?$select=Title&$filter=Hidden eq false and Title eq '${name}'`);
  out.list.inSiteContentsQuery = (visible.value || []).length > 0;
  // 2. Plain multi-line text field, written with growing JSON.
  out.addField = await call(`${L}/fields/createfieldasxml`, 'POST', { parameters: { SchemaXml: '<Field Type="Note" Name="StateJson" StaticName="StateJson" DisplayName="StateJson" RichText="FALSE" NumLines="6" />', Options: 0 } });
  out.addField = out.addField.httpStatus ? out.addField : 'ok';
  const item = await call(`${L}/items`, 'POST', { Title: 'run-1' });
  out.item = item.httpStatus ? item : item.Id;
  if (item.httpStatus) { console.log(JSON.stringify(out, null, 2)); return; }
  const I = `${L}/items(${item.Id})`;
  out.noteField = {};
  for (const size of [10000, 70000, 300000]) {
    const value = JSON.stringify({ pad: 'x'.repeat(size - 10) });
    const w = await call(I, 'MERGE', { StateJson: value });
    const back = w.httpStatus ? null : await call(`${I}?$select=StateJson`);
    out.noteField[size] = w.httpStatus ? w : { written: value.length, readBack: back && back.StateJson ? back.StateJson.length : back };
  }
  // 3. Attachment: add, overwrite in place, read back.
  const big = JSON.stringify({ pad: 'y'.repeat(500000) });
  const add = await call(`${I}/AttachmentFiles/add(FileName='state.json')`, 'POST', new Blob(['{"v":1}']));
  out.attachment = { add: add.httpStatus ? add : 'ok' };
  const put = await call(`${I}/AttachmentFiles('state.json')/$value`, 'PUT', new Blob([big]));
  out.attachment.overwritePut = put.httpStatus ? put : 'ok';
  const readBack = await fetch(`${site}/_api/${I}/AttachmentFiles('state.json')/$value`);
  out.attachment.readBackLength = readBack.ok ? (await readBack.text()).length : readBack.status;
  out.attachment.expectedLength = big.length;
  const again = await call(`${I}/AttachmentFiles/add(FileName='state.json')`, 'POST', new Blob(['{"v":2}']));
  out.attachment.addSameNameAgain = again.httpStatus ? again : 'ok';
  // 4. Owners only: break inheritance without copying, then add the owner group.
  const me = await call('web/currentuser?$select=Id,LoginName,IsSiteAdmin');
  out.me = { id: me.Id, siteAdmin: me.IsSiteAdmin };
  out.breakInheritance = await call(`${L}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`, 'POST');
  out.breakInheritance = out.breakInheritance.httpStatus ? out.breakInheritance : 'ok';
  const ra = await call(`${L}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=Member/Title,Member/Id,Member/PrincipalType,RoleDefinitionBindings/Name`);
  out.afterBreak = (ra.value || []).map((a) => `${a.Member.Title} (${a.Member.Id}, type ${a.Member.PrincipalType}): ${a.RoleDefinitionBindings.map((b) => b.Name).join(', ')}`);
  const owner = await call('web/AssociatedOwnerGroup?$select=Id,Title');
  const full = await call('web/roledefinitions/getbytype(5)?$select=Id,Name');
  out.addOwnerGroup = await call(`${L}/roleassignments/addroleassignment(principalid=${owner.Id},roledefid=${full.Id})`, 'POST');
  out.addOwnerGroup = out.addOwnerGroup.httpStatus ? out.addOwnerGroup : `ok: ${owner.Title} → ${full.Name}`;
  const ra2 = await call(`${L}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=Member/Title,Member/Id,RoleDefinitionBindings/Name`);
  out.afterOwnerGroup = (ra2.value || []).map((a) => `${a.Member.Title} (${a.Member.Id}): ${a.RoleDefinitionBindings.map((b) => b.Name).join(', ')}`);
  const stillRead = await call(`${I}?$select=Id`);
  out.installerCanStillRead = !stillRead.httpStatus;
  // Cleanup: to the recycle bin.
  if (CLEANUP) { const r = await call(`${L}/recycle`, 'POST'); out.cleanup = r.httpStatus ? r : 'recycled'; }
  console.log(JSON.stringify(out, null, 2));
})();
```

**Futtatás:** a Cél site bármelyik lapján F12 → Console, másold be, Enter. A kimenő JSON-t kérem vissza.

## Eredmények (2026-10-05, Cél site, site-gyűjtemény-adminisztrátorként)

1. **Rejtett lista:** a `Hidden: true`, `NoCrawl: true` beállítás létrehozáskor működik; a lista a `…/Lists/CopyJetSpike14` címen jött létre, és a `Hidden eq false` lekérdezés (Webhely tartalma) nem adja vissza.
2. **Hosszú szöveg:** a sima szöveges `Note` mező 10 000, 70 000 és 300 000 karaktert is hiánytalanul visszaadott.
3. **Melléklet:**
   - a `state.json` felvétele működik;
   - a helyben felülírás (`AttachmentFiles('state.json')/$value`, `X-HTTP-Method: PUT`) is működik, 500 010 karakter hiánytalanul visszaolvasva;
   - ugyanazon a néven újra felvenni nem lehet: HTTP 400, `-2130575257` („A megadott név már használatban van”).
4. **Jogosultság:**
   - `breakroleinheritance(false, true)` után csak a futtató felhasználó maradt a listán, Teljes hozzáféréssel;
   - a tulajdonoscsoport (`AssociatedOwnerGroup`) Teljes hozzáféréssel felvehető;
   - a futtató utána is olvas és ír.

## Döntések

- **A `CopyJetLog` lista:**
  - rejtett egyéni lista (`Hidden`, `NoCrawl`), az első mentéskor jön létre;
  - az öröklést megszünteti: a futtató felhasználó és a tulajdonoscsoport marad rajta, mindkettő Teljes hozzáféréssel;
  - ha a jogosultság beállítása nem sikerül, a lista örökölt jogokkal marad, és a napló figyelmeztet (`STATE_LIST_PERMISSIONS`).
- **Egy futás = egy elem:**
  - oszlopok: `CJRunId`, `CJTemplate`, `CJChecksum`, `CJStatus`;
  - a teljes állapot (kész lépések, tokenek, elem-ID-leképezések) a `state.json` mellékletbe kerül, és minden mentéskor helyben felülíródik (`PUT`): egy hívás lépésenként.
  - Miért melléklet: a jegyzetmező 300 000 karaktert is elbírt, de egy nagy lista ID-leképezése ennél nagyobb lehet, a melléklet mérete pedig nem korlát.
- **Mikor ment:**
  - minden lépés után;
  - a listaelemek írása közben 500 elemenként;
  - a fájlok feltöltése közben 25 fájlonként.

  A mentések sorba rendeződnek: amíg egy fut, a következő a legfrissebb állapotot írja.
- **A folytatás:**
  - csak ugyanazt a sablont ajánlja fel: a manifest SHA-256-ja `meta.checksum` nélkül;
  - a kész lépéseket kihagyja, a tokeneket és az ID-leképezéseket visszatölti;
  - a site saját tokenjeit és a felhasználó-leképezést frissen olvassa.
- **Félbemaradt elemírás:**
  - írás előtt a program elmenti, mely forráselemek következnek (`pendingItems`, 500-as csomag);
  - folytatáskor a feljegyzett elemeket ismertnek veszi;
  - a feljegyzés nélküli célelemeket ID-sorrendben a mentett csomag elejéhez párosítja, és a címükkel ellenőrzi (`ITEMS_RESUME_RECOVERED`). A SharePoint a kötegen belül sorban ír, a kötegek egymás után futnak;
  - utána csak a hiányzókat írja;
  - ha ez nem egyértelmű (több ismeretlen elem, vagy eltérő cím), nem ír semmit (`ITEMS_RESUME_UNKNOWN_ITEMS`).

## Utólagos megjegyzés (1.8.1.0)

Az első valódi telepítés `CopyJetLog` listájának oszlopai `RunId`, `Template`, `Checksum`, `Status` belső névvel jöttek létre, nem `CJ…`-vel. A `createfieldasxml` `Options: 0` mellett ugyanis a `DisplayName`-ből képzi a belső nevet; a spike-ban a kettő egyezett, ezért ott nem jött elő. Javítás: `AddFieldInternalNameHint` (8). A már meglévő listához a program olvasás előtt hozzáadja a hiányzó `CJ…` oszlopokat, a régieket nem törli.

## Utólagos megjegyzés (1.8.2.0)

A lapbezárásos teszt elemírás közben szakadt meg: a két elem már a listában volt, de a feljegyzés üres maradt (`"Teszt_lookup_forrs": {}`), ezért a folytatás nem írt a listába, és a lookupok üresek maradtak. Javítás: a fent leírt `pendingItems` csomagmentés és párosítás.
