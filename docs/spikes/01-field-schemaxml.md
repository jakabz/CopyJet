# Spike 01 – Site columnok `SchemaXml`-je és a mező-lekérdezések

Állapot: **lezárva** (2026-09-24, Forrás site, magyar nyelvű) · Érintett kód: `src/core/fields/schemaXml.ts`, `SiteFieldExtractor`, `SiteFieldProvider`

## Kérdések

1. **Szűrés:** a `/_api/web/fields?$filter=CanBeDeleted eq true and Hidden eq false` szerveroldalon működik-e, és pontosan az egyedi site columnokat adja-e vissza?
2. **Attribútumok:** a valós `SchemaXml`-ekben milyen gyökér-attribútumok és gyerekelemek fordulnak elő? Ezek közül melyek hiányoznak az engedélylistáról, és melyiket kell megtartani?
3. **ID-ütközés:** működik-e a `/_api/web/availablefields?$filter=Id eq guid'…'` szűrő? A provider ezzel ellenőrzi, hogy az ID-t nem foglalja-e már egy másik belső nevű mező.
4. **Választási lehetőségek frissítése** (írási művelet, a telepítési teszthez tartozik, nem ehhez a kísérlethez): a `MERGE` `{"Choices": [...]}` törzzsel `odata=nometadata` módban frissíti-e a Choice mező választási lehetőségeit?

## Jelenlegi feltevések (a kód ezekre épül)

- A szűrés, a sanitizálás és a tokenizálás a Setup oldalon történik. Az Install a feloldás után újra sanitizál, mert egy sablon kézzel is szerkeszthető.
- **Eldobott attribútumok:**
  - `SourceID`, `WebId`, `ColName`, `RowOrdinal`: forrás-site-hoz kötött értékek;
  - `Version`: sémaverzió;
  - `JSLink`: a sablonból nem futhat kód;
  - minden, ami nincs az engedélylistán.
- **Megtartott gyerekelemek:** `Default`, `DefaultFormula`, `CHOICES`, `MAPPINGS`, `Formula`, `FieldRefs`, `Validation`, `Customization`.
- **A site column `ID`-ja változatlanul marad** a sablonban, mert a tartalomtípusok erre hivatkoznak. A többi GUID tokenizálódik, pl. a lookup `List="{{listkey:X}}"`.
- **A taxonómiamezőket a provider az 1. fázisban nem hozza létre** (`unsupported` diff). A termkészlet-leképezés a 2. fázis `TermMapper` feladata.
- **Átnevezés (`rename`) site columnnál nem támogatott:** új belső név és új ID keletkezne, ami eltörné a tartalomtípus- és nézethivatkozásokat. Ilyenkor a provider figyelmeztet és kihagy.

## Kísérlet (csak olvas, a site-on nem módosít semmit)

1. Nyisd meg a **Forrás** site bármely lapját tulajdonosként, és nyomd meg az F12-t, majd válaszd a **Console** fület.
2. Illeszd be az alábbi kódot, és nyomj Entert.
3. A konzolba kiírt JSON-t küldd vissza. Ha a böngésző engedi, a kód a vágólapra is másolja.

```js
(async () => {
  // Modern pages do not always expose _spPageContextInfo: start from the page URL, then ask the web for its own URL.
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url', { headers: { Accept: 'application/json;odata=nometadata' } })).json();
  const web = webInfo.Url;
  const get = async (path) => {
    const r = await fetch(web + path, { headers: { Accept: 'application/json;odata=nometadata' } });
    return { status: r.status, body: r.ok ? await r.json() : await r.text() };
  };
  const allowed = ['ID','Name','StaticName','DisplayName','Type','Group','Description','Required','Hidden','ReadOnly','Sealed','Indexed','EnforceUniqueValues','Viewable','Filterable','Sortable','ShowInNewForm','ShowInEditForm','ShowInDisplayForm','ShowInViewForms','ShowInFileDlg','ShowInListSettings','CustomFormatter','ClientSideComponentId','ClientSideComponentProperties','Format','MaxLength','NumLines','RichText','RichTextMode','AppendOnly','UnlimitedLengthInDocumentLibrary','IsolateStyles','RestrictedMode','Min','Max','Decimals','Percentage','LCID','CommaSeparator','ResultType','FillInChoice','FriendlyDisplayFormat','CalType','StorageTZ','DisplayFormat','List','ShowField','Mult','RelationshipDeleteBehavior','PrependId','FieldRef','UserSelectionMode','UserSelectionScope','Presence'];
  const filtered = await get("/_api/web/fields?$filter=CanBeDeleted eq true and Hidden eq false&$select=Id,InternalName,TypeAsString,Group,SchemaXml");
  const all = await get('/_api/web/fields?$select=InternalName,CanBeDeleted,Hidden');
  const fields = Array.isArray(filtered.body?.value) ? filtered.body.value : [];
  const attrs = {}, elements = {}, perType = {};
  for (const f of fields) {
    const el = new DOMParser().parseFromString(f.SchemaXml, 'application/xml').documentElement;
    for (const a of el.attributes) attrs[a.name] = (attrs[a.name] || 0) + 1;
    for (const c of el.children) elements[c.nodeName] = (elements[c.nodeName] || 0) + 1;
    perType[f.TypeAsString] = (perType[f.TypeAsString] || 0) + 1;
  }
  const firstId = fields[0]?.Id;
  const byId = firstId ? await get(`/_api/web/availablefields?$filter=Id eq guid'${firstId}'&$select=Id,InternalName`) : null;
  const result = {
    web,
    q1_filterStatus: filtered.status,
    q1_filteredCount: fields.length,
    q1_customByClientFilter: Array.isArray(all.body?.value) ? all.body.value.filter(f => f.CanBeDeleted && !f.Hidden).length : all.status,
    q1_fields: fields.map(f => `${f.InternalName} (${f.TypeAsString}, ${f.Group})`),
    q2_attributeCounts: attrs,
    q2_notAllowed: Object.keys(attrs).filter(a => !allowed.includes(a)),
    q2_elementCounts: elements,
    q2_types: perType,
    q3_byIdStatus: byId && byId.status,
    q3_byIdResult: byId && byId.body
  };
  const json = JSON.stringify(result, null, 2);
  console.log(json);
  // copy() only exists in some DevTools consoles; the clipboard API may need page focus.
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## Mit várunk?

- `q1_filterStatus` = 200, és `q1_filteredCount` = `q1_customByClientFilter`. Ha a szerver elutasítja a szűrőt, kliensoldali szűrésre váltunk.
- `q2_notAllowed` minden tételéről döntünk: felvesszük az engedélylistára, vagy tudatosan eldobjuk. Az eldobottak listája a Setup naplójában is megjelenik (`FIELD_XML_SANITIZED`).
- `q3_byIdStatus` = 200, és egy találatot ad. Ha nem, az ID-ütközést `getById` hívással és 404-kezeléssel ellenőrizzük.

## Eredmény

**Q1 – a szűrő működik, de nem az egyedi oszlopokat adja.** A `CanBeDeleted eq true and Hidden eq false` HTTP 200-zal 226 mezőt adott vissza, és a kliensoldali ellenőrzés is 226-ot számolt. Ezek szinte mind beépített oszlopok (`FirstName`, `Author`, `DocIcon`, `Created` …, köztük `_Hidden` csoportúak is). A valódi egyedi site column csak kettő volt: `SiteColumn1` (Text) és `SiteColumn2` (Number). A csoport sem jó ismérv: az „Egyéni oszlopok” csoportban beépített mezők is vannak (`TaskOutcome`, `_ColorTag`, `ComplianceAssetId`, `TagEventDate`).

→ **Döntés:**
- Szerveroldalon csak a `Hidden eq false` szűrő marad.
- Az egyedi oszlopot a `SchemaXml` `SourceID` attribútuma azonosítja (`isCustomField`): a felhasználó által létrehozottnál ez a létrehozó web GUID-ja, a beépítettnél `http://schemas.microsoft.com/sharepoint/v3…`. A 226 mezőből 215-nek van `SourceID`-ja.
- A `SchemaXml` szerveroldalon nem szűrhető, ezért a szűrés a kliensen történik. Ehhez a feltérképezés is lekéri a `SchemaXml`-t: kb. 230 mező, egyszeri hívás.

**Q2 – attribútumok.** Az engedélylistán nem szereplő 29 attribútum gyakorlatilag mind beépített mezőkből jön, és a Q1-es javítás után ezek nem kerülnek a sablonba.
- **Felvettük** (leíró, ártalmatlan): `AuthoringInfo`, `Dir`, `IMEMode`, `ShowInVersionHistory`.
- **Tudatosan eldobjuk:**
  - forrás-site-hoz kötött értékek: `SourceID`, `WebId`, `ColName`, `RowOrdinal`, `JoinColName`, `JoinRowOrdinal`, `JoinType`;
  - verzió- és telepítési metaadatok: `Version`, `DisplaceOnUpgrade`, `DelayActivateTemplateBinding`, `FromBaseType`, `AllowDeletion`, `CanToggleHidden`, `NoCustomize`;
  - beépített vagy Computed mezőkre jellemzők: `DefaultListField`, `DisplayNameSrcField`, `Node`, `ClassInfo`, `TextOnly`, `CalloutMenu`, `DisplaySize`, `WikiLinking`, `ListItemMenuAllowed`, `LinkToItemAllowed`, `HeaderImage`.
- A gyerekelemek közül a `DisplayPattern` csak Computed mezőkben fordult elő, ezért nem került az engedélylistára.

**Q3 – az ID szerinti szűrő működik.** Az `availablefields?$filter=Id eq guid'…'` HTTP 200-at adott egy találattal. A provider marad ennél a megoldásnál.

**Q4 – nyitott.** A választási lehetőségek `MERGE`-es frissítését az első telepítési teszten ellenőrizzük.

**Tanulság a teszteléshez:** a Forrás site-on csak Text és Number típusú egyedi oszlop van. A lookup, a Choice, a Person és a taxonómia ágak ellenőrzéséhez ilyen oszlopokat is létre kell hozni.
