# Spike 11 – Managed Metadata oszlopok és termek

Állapot: **lezárva** · Érintett kód (2. fázis): lista- és site-oszlopok (`TaxonomyFieldType`), `FieldValueSerializer`, `TermMapper`, Install `MappingStep`

## Kérdések

1. **Oszlop:** mit tárol a SharePoint egy Managed Metadata oszlop `SchemaXml`-jében?
   - termtár (`SspId`), termkészlet (`TermSetId`), horgony (`AnchorId`);
   - a rejtett jegyzetoszlop (`TextField`), többes érték, nyitottság.

   Ebből mi tenant-függő, és mi kell az újralétrehozáshoz?
2. **Érték:** milyen alakban adja vissza a REST egy- és többértékű mező értékét (`Label`, `TermGuid`, `WssId`)?
3. **Termtár elérése:** működik-e a `_api/v2.1/termStore` a böngészőből?
   - termkészlet szülőcsoporttal;
   - a termek címkéi és gyerekei;
   - egy term szülője (a címke-útvonalhoz);
   - a csoportok.
4. **Írás** (B kísérlet, az A eredménye után):
   - létrehozható-e egy Managed Metadata oszlop `createFieldAsXml`-lel;
   - milyen alakban fogadja az értéket az `AddValidateUpdateItemUsingPath`.

## Előkészítés (Forrás tenant)

1. **Termtár** (SharePoint felügyeleti központ → Tartalomszolgáltatások → Termtár), vagy egy site-szintű termcsoport:
   - csoport: **CopyJet teszt**;
   - termkészlet: **Terület**;
   - termek: **Pénzügy**, **Logisztika**, és a Logisztika alatt egy gyerek-term, a **Raktár**.
2. A Forrás „Teszt lista” listájához két oszlop, mindkettő **Felügyelt metaadatok** típusú, a „Terület” termkészlethez kötve:
   - **Terület** – egy érték;
   - **Címkék** – több érték engedélyezve.
3. Tölts ki legalább két elemet: az egyiknél Terület = Pénzügy, a másiknál Címkék = Logisztika és Raktár.

## A. kísérlet – Forrás site (csak olvas)

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (url) => { const r = await fetch(url, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status, error: (await r.text()).slice(0, 200) }; };
  const webInfo = await get(guess + '/_api/web?$select=Url,ServerRelativeUrl');
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const LIST = `${web}/_api/web/getList('${enc(rel + '/Lists/Teszt lista')}')`;
  const out = {};
  // 1. Managed Metadata columns: the settings SharePoint keeps in their SchemaXml.
  const all = (await get(`${LIST}/fields?$select=Id,InternalName,Title,TypeAsString,Hidden,SchemaXml`)).value || [];
  const props = (xml) => { const o = {}; const re = /<Name>(\w+)<\/Name>\s*<Value[^>]*>([^<]*)<\/Value>/g; let m; while ((m = re.exec(xml))) o[m[1]] = m[2]; return o; };
  const attr = (xml, a) => (new RegExp(`\\s${a}="([^"]*)"`).exec(xml) || [])[1];
  const tax = all.filter((f) => /^TaxonomyFieldType/.test(f.TypeAsString));
  out.columns = tax.map((f) => {
    const p = props(f.SchemaXml);
    const note = all.filter((x) => x.Id.toLowerCase() === (p.TextField || '').replace(/[{}]/g, '').toLowerCase())[0];
    return { name: f.InternalName, type: f.TypeAsString, mult: attr(f.SchemaXml, 'Mult'), showField: attr(f.SchemaXml, 'ShowField'), list: attr(f.SchemaXml, 'List') ? 'set' : null,
      sspId: p.SspId, termSetId: p.TermSetId, anchorId: p.AnchorId, textField: p.TextField, open: p.Open, isPathRendered: p.IsPathRendered, hiddenNoteField: note ? `${note.InternalName} (Hidden: ${note.Hidden})` : null };
  });
  // Site column for comparison: Enterprise Keywords (TaxKeyword) carries the default term store id.
  const keyword = await get(`${web}/_api/web/availablefields/getbyinternalnameortitle('TaxKeyword')?$select=SchemaXml`);
  out.taxKeywordSspId = keyword.SchemaXml ? props(keyword.SchemaXml).SspId : keyword;
  // 2. Item values as REST returns them.
  const names = tax.map((f) => f.InternalName);
  if (names.length) {
    const items = (await get(`${LIST}/items?$select=Id,Title,${names.join(',')}&$top=5&$orderby=ID`)).value || [];
    out.items = items.map((i) => { const o = { id: i.Id, title: i.Title }; names.forEach((n) => (o[n] = i[n])); return o; });
  }
  // 3. Term store (v2.1): the store, the column's term set, its group and its terms.
  const TS = `${web}/_api/v2.1/termStore`;
  const store = await get(TS);
  out.termStore = store.httpStatus ? store : { id: store.id, defaultLanguageTag: store.defaultLanguageTag, languageTags: store.languageTags };
  const first = out.columns[0];
  if (first && first.termSetId) {
    const setId = first.termSetId;
    const set = await get(`${TS}/sets/${setId}?$expand=parentGroup`);
    out.set = set.httpStatus ? set : { id: set.id, names: set.localizedNames, group: set.parentGroup && { id: set.parentGroup.id, name: set.parentGroup.name } };
    const children = await get(`${TS}/sets/${setId}/children?$select=id,labels,childrenCount`);
    out.topTerms = children.httpStatus ? children : (children.value || []).map((t) => ({ id: t.id, labels: (t.labels || []).map((l) => `${l.languageTag}:${l.name}${l.isDefault ? '*' : ''}`), children: t.childrenCount }));
    const withChildren = (children.value || []).filter((t) => t.childrenCount > 0)[0];
    if (withChildren) {
      const sub = await get(`${TS}/sets/${setId}/terms/${withChildren.id}/children?$select=id,labels`);
      out.subTerms = sub.httpStatus ? sub : (sub.value || []).map((t) => ({ id: t.id, label: (t.labels || [])[0] && t.labels[0].name }));
    }
    // Flat list of every term of the set (would give paths in one call), and a term's parent.
    const flat = await get(`${TS}/sets/${setId}/terms?$select=id,labels&$top=5`);
    out.flatTerms = flat.httpStatus ? `HTTP ${flat.httpStatus}` : `${(flat.value || []).length} terms`;
    const someTerm = (out.subTerms && out.subTerms[0]) || (out.topTerms && out.topTerms[0]);
    if (someTerm && someTerm.id) {
      const parent = await get(`${TS}/sets/${setId}/terms/${someTerm.id}?$expand=parent`);
      out.termWithParent = parent.httpStatus ? `HTTP ${parent.httpStatus}` : { id: parent.id, parent: parent.parent ? parent.parent.id : null };
    }
    const groups = await get(`${TS}/groups?$select=id,name`);
    out.groups = groups.httpStatus ? groups : (groups.value || []).map((g) => g.name);
  }
  const json = JSON.stringify(out, null, 2).split(location.host).join('contoso.sharepoint.com');
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- a `columns` mutatja a két oszlop `SchemaXml`-beállításait (`sspId`, `termSetId`, `textField`, `hiddenNoteField`);
- az `items` a REST-értékeket;
- a `termStore`, `set`, `topTerms`, `subTerms`, `termWithParent` és `groups` alapján kiderül, mely termtár-hívások működnek a böngészőből.

## Eredmény

### A – kiolvasás (2026-09-30, Forrás site)

- **Oszlop (`SchemaXml` → `Customization`):**
  - `SspId` = a termtár azonosítója. Egyezik a `termStore.id`-vel és a `TaxKeyword` `SspId`-jével.
  - `TermSetId`, `AnchorId` (nulla GUID = a teljes készlet), `Open: false`, `IsPathRendered: false`.
  - `TextField` = a rejtett jegyzetoszlop (Note, GUID-szerű belső név, pl. `g6245…`).
  - `ShowField="Term1033"`, `List` = a TaxonomyHiddenList.
  - Többértékű: `TaxonomyFieldTypeMulti`, `Mult="TRUE"`.
  - **Tenant-függő:** `SspId`, `TermSetId`, `TextField`, `List`. A többi hordozható.
- **Értékek (REST):**
  - Többértékű: tömb, `{ Label, TermGuid, WssId }` elemekkel.
  - ⚠️ **Egyértékű:** `{ Label: "8", TermGuid, WssId: 8 }`, vagyis a `Label` a WssId, nem a címke. A címke a termtárból jön, a `TermGuid` alapján.
- **Termtár (`_api/v2.1/termStore`), minden működik a böngészőből:**
  - store: `id`, `defaultLanguageTag: en-US`, `languageTags: [en-US]` (a magyar címkék is `en-US` alatt vannak);
  - `sets/{id}?$expand=parentGroup` → név és csoport;
  - `sets/{id}/children` → felső termek (címkék, `childrenCount`);
  - `sets/{id}/terms/{id}/children`;
  - `sets/{id}/terms` → a teljes lapos lista;
  - `sets/{id}/terms/{id}?$expand=parent` → a szülő;
  - `groups` → a csoportnevek (a site-szintű csoport neve tartalmazza a tenant hosztját).
- Címke-útvonal: `CopyJet teszt / Terület / Logisztika / Raktár`.

## B. kísérlet – Cél site (**ír**, önálló)

Csak a Cél **teszt** site-on futtasd. Ha a Cél másik tenantban van, ott is legyen meg a **CopyJet teszt / Terület** termkészlet a három termmel (ugyanazokkal a nevekkel).

- **Keresés útvonal szerint:** csoport → termkészlet → termek, a szülőkkel együtt.
- **Oszlopok:** a `CopyJetSpike11` listán (létrehozza, ha nincs) egy egy- és egy többértékű Managed Metadata oszlop `createFieldAsXml`-lel, `TextField` nélkül. Megnézi, hogy a SharePoint maga köti-e be a rejtett jegyzetoszlopot.
- **Értékek:**
  - T1: `Címke|GUID`;
  - T2: `-1;#Címke|GUID`;
  - T3: csak GUID;
  - T4: rossz címke;
  - T5: ismeretlen GUID.

```js
(async () => {
  // The term set to look for by path (group / set), as a template would carry it.
  const GROUP = 'CopyJet teszt', SET = 'Terület';
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (url) => { const r = await fetch(url, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status, error: (await r.text()).slice(0, 200) }; };
  const webInfo = await get(guess + '/_api/web?$select=Url,ServerRelativeUrl');
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const TS = `${web}/_api/v2.1/termStore`;
  const out = { find: {}, setup: {}, tests: {} };

  // 1. Find the term set by group name and set name (what a cross-tenant install has to do).
  const store = await get(TS);
  const groups = (await get(`${TS}/groups?$select=id,name`)).value || [];
  const group = groups.filter((g) => g.name === GROUP)[0];
  out.find.group = group ? 'found' : `missing (groups: ${groups.map((g) => g.name).join(', ')})`;
  if (!group) { console.log(JSON.stringify(out, null, 2)); return; }
  const sets = (await get(`${TS}/groups/${group.id}/sets?$select=id,localizedNames`)).value || [];
  const set = sets.filter((s) => (s.localizedNames || []).some((n) => n.name === SET))[0];
  out.find.set = set ? 'found' : 'missing';
  if (!set) { console.log(JSON.stringify(out, null, 2)); return; }
  // Every term with its parent, for label paths: flat list with $expand=parent.
  const flat = await get(`${TS}/sets/${set.id}/terms?$select=id,labels&$expand=parent`);
  out.find.flatWithParent = flat.httpStatus ? `HTTP ${flat.httpStatus}` : (flat.value || []).map((t) => ({ label: t.labels[0].name, parent: t.parent ? t.parent.id.slice(0, 8) : null }));
  const terms = (await get(`${TS}/sets/${set.id}/terms?$select=id,labels`)).value || [];
  const byLabel = {};
  terms.forEach((t) => (byLabel[t.labels[0].name] = t.id));
  out.find.terms = Object.keys(byLabel);

  // 2. A list with a single and a multi value Managed Metadata column, created from SchemaXml (no TextField given).
  const listUrl = `${rel}/Lists/CopyJetSpike11`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  if ((await fetch(`${LIST}?$select=Id`, { headers: H })).status === 404) {
    out.setup.listCreated = (await fetch(`${web}/_api/web/lists`, { method: 'POST', headers: W, body: JSON.stringify({ Title: 'CopyJetSpike11', BaseTemplate: 100 }) })).status;
  }
  const prop = (n, v) => `<Property><Name>${n}</Name><Value xmlns:q1="http://www.w3.org/2001/XMLSchema" p4:type="q1:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">${v}</Value></Property>`;
  const custom = `<Customization><ArrayOfProperty>${prop('SspId', store.id)}${prop('TermSetId', set.id)}${prop('AnchorId', '00000000-0000-0000-0000-000000000000')}${prop('Open', 'false')}${prop('IsPathRendered', 'false')}</ArrayOfProperty></Customization>`;
  const addField = async (name, xml) => {
    if ((await fetch(`${LIST}/fields/getbyinternalnameortitle('${name}')?$select=Id`, { headers: H })).ok) return 'exists';
    const r = await fetch(`${LIST}/fields/createfieldasxml`, { method: 'POST', headers: W, body: JSON.stringify({ parameters: { SchemaXml: xml, Options: 12 } }) });
    return r.ok ? 'created' : `HTTP ${r.status}: ${(await r.text()).slice(0, 200)}`;
  };
  out.setup.single = await addField('CJTax', `<Field Type="TaxonomyFieldType" Name="CJTax" StaticName="CJTax" DisplayName="CJ taxonómia" ShowField="Term1033">${custom}</Field>`);
  out.setup.multi = await addField('CJTaxMulti', `<Field Type="TaxonomyFieldTypeMulti" Name="CJTaxMulti" StaticName="CJTaxMulti" DisplayName="CJ taxonómia multi" ShowField="Term1033" Mult="TRUE">${custom}</Field>`);
  // Did SharePoint wire the hidden note field (TextField) and the TaxonomyHiddenList?
  const fields = (await get(`${LIST}/fields?$select=Id,InternalName,Hidden,TypeAsString,SchemaXml`)).value || [];
  const describe = (name) => {
    const f = fields.filter((x) => x.InternalName === name)[0];
    if (!f) return 'missing';
    const tf = (/<Name>TextField<\/Name>\s*<Value[^>]*>([^<]*)</.exec(f.SchemaXml) || [])[1];
    const note = tf ? fields.filter((x) => x.Id.toLowerCase() === tf.replace(/[{}]/g, '').toLowerCase())[0] : null;
    return { type: f.TypeAsString, textField: tf || null, note: note ? `${note.InternalName} (${note.TypeAsString}, hidden: ${note.Hidden})` : null, hiddenList: /\sList="[^"]+"/.test(f.SchemaXml) };
  };
  out.setup.singleField = describe('CJTax');
  out.setup.multiField = describe('CJTaxMulti');

  // 3. Values through AddValidateUpdateItemUsingPath, in several notations.
  const a = byLabel['Pénzügy'], b = byLabel['Logisztika'], c = byLabel['Raktár'];
  const add = async (title, values) => {
    const r = await fetch(`${LIST}/AddValidateUpdateItemUsingPath`, { method: 'POST', headers: W, body: JSON.stringify({
      listItemCreateInfo: { FolderPath: { DecodedUrl: listUrl }, UnderlyingObjectType: 0 },
      formValues: [{ FieldName: 'Title', FieldValue: title }].concat(Object.keys(values).map((k) => ({ FieldName: k, FieldValue: values[k] }))), bNewDocumentUpdate: false }) });
    const body = r.ok ? await r.json() : { value: [{ FieldName: 'HTTP', HasException: true, ErrorMessage: String(r.status) }] };
    const id = Number(((body.value || []).filter((v) => v.FieldName === 'Id')[0] || {}).FieldValue);
    const errors = (body.value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`);
    let back = null;
    if (id) {
      const i = await get(`${LIST}/items(${id})?$select=CJTax,CJTaxMulti`);
      back = { CJTax: i.CJTax, CJTaxMulti: i.CJTaxMulti };
    }
    return { sent: values, errors, back };
  };
  out.tests.T1_label_guid = await add('T1', { CJTax: `Pénzügy|${a}`, CJTaxMulti: `Logisztika|${b};Raktár|${c}` });
  out.tests.T2_wssid_notation = await add('T2', { CJTax: `-1;#Pénzügy|${a}`, CJTaxMulti: `-1;#Logisztika|${b};#-1;#Raktár|${c}` });
  out.tests.T3_guid_only = await add('T3', { CJTax: a });
  out.tests.T4_wrong_label = await add('T4', { CJTax: `Rossz címke|${a}` });
  out.tests.T5_unknown_guid = await add('T5', { CJTax: 'Pénzügy|00000000-1111-2222-3333-444444444444' });
  const json = JSON.stringify(out, null, 2).split(location.host).join('contoso.sharepoint.com');
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

### B – eredmény (2026-09-30, másik tenant és Forrás)

- ✅ **Keresés név szerint** (csoport → termkészlet → termek) mindkét tenantban működik.
- ⚠️ A lapos `sets/{id}/terms?$expand=parent` **nem adja a szülőt** (a Raktárnál is `null`). A címke-útvonalhoz a `children` hívásokon kell végigmenni (az A-ban működött).
- ❌ **`createFieldAsXml` csak `Customization`-nel** (`SspId`, `TermSetId`, `TextField`, `List` nélkül): HTTP 500, `TargetInvocationException`, mindkét tenantban.
  - ⚠️ **Utána a lista használhatatlan:** a `fields` lekérdezése is 500-at ad, és minden elemírás 500. Egy félresikerült taxonómia-oszlop tönkreteszi a listát, ezért a CopyJet csak igazolt módon hozhat létre ilyen oszlopot. A `CopyJetSpike11` listát kézzel törölni kell.
- A T1–T5 értékírás emiatt nem futott le.

## C. kísérlet – Cél site (**ír**): két létrehozási mód, külön friss listán

Minden változat új, időbélyeges listát kap (`CJ11C1…`, `CJ11C2…`), így egy hiba nem rontja el a másikat.

- **C1, teljes XML:** előbb a rejtett jegyzetoszlop (Note, `Hidden`), utána a taxonómia-oszlop `List` (TaxonomyHiddenList), `WebId`, `TextField` és a teljes `Customization` megadásával.
- **C2, CSOM** (a PnP módszere): `AddFieldAsXml` egyszerű oszloppal, utána `SspId`, `TermSetId`, `AnchorId` és `Update()`, egy kérésben.

Mindkettőnél megnézi, hogy létrejött-e az oszlop és a rejtett segédoszlop, és hogy működik-e az értékírás: T1 `Címke|GUID`, T2 `-1;#Címke|GUID`, T3 ismeretlen GUID.

```js
(async () => {
  const GROUP = 'CopyJet teszt', SET = 'Terület';
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (url) => { const r = await fetch(url, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status, error: (await r.text()).slice(0, 200) }; };
  const webInfo = await get(guess + '/_api/web?$select=Id,Url,ServerRelativeUrl');
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const TS = `${web}/_api/v2.1/termStore`;
  const out = { find: {}, C1_fullXml: {}, C2_csom: {} };

  // The term set by path, and its terms (labels → IDs).
  const store = await get(TS);
  const group = ((await get(`${TS}/groups?$select=id,name`)).value || []).filter((g) => g.name === GROUP)[0];
  const set = group ? ((await get(`${TS}/groups/${group.id}/sets?$select=id,localizedNames`)).value || []).filter((s) => (s.localizedNames || []).some((n) => n.name === SET))[0] : null;
  if (!set) { console.log('Term set not found: ' + GROUP + ' / ' + SET); return; }
  const byLabel = {};
  ((await get(`${TS}/sets/${set.id}/terms?$select=id,labels`)).value || []).forEach((t) => (byLabel[t.labels[0].name] = t.id));
  const a = byLabel['Pénzügy'], b = byLabel['Logisztika'], c = byLabel['Raktár'];
  out.find = { set: 'found', terms: Object.keys(byLabel) };
  const hidden = await get(`${web}/_api/web/getList('${enc(rel + '/Lists/TaxonomyHiddenList')}')?$select=Id`);
  out.find.taxonomyHiddenList = hidden.Id ? 'found' : hidden;

  const stamp = Date.now();
  const newList = async (title) => {
    const r = await fetch(`${web}/_api/web/lists`, { method: 'POST', headers: W, body: JSON.stringify({ Title: title, BaseTemplate: 100 }) });
    return r.ok ? `${rel}/Lists/${title}` : null;
  };
  const listApi = (url) => `${web}/_api/web/getList('${enc(url)}')`;
  const describe = async (listUrl, name) => {
    const fields = (await get(`${listApi(listUrl)}/fields?$select=Id,InternalName,Hidden,TypeAsString,SchemaXml`));
    if (fields.httpStatus) return `fields: HTTP ${fields.httpStatus}`;
    const f = (fields.value || []).filter((x) => x.InternalName === name)[0];
    if (!f) return 'missing';
    const p = (n) => (new RegExp(`<Name>${n}</Name>\\s*<Value[^>]*>([^<]*)<`).exec(f.SchemaXml) || [])[1];
    const tf = p('TextField');
    const note = tf ? (fields.value || []).filter((x) => x.Id.toLowerCase() === tf.replace(/[{}]/g, '').toLowerCase())[0] : null;
    return { type: f.TypeAsString, sspId: p('SspId') === store.id ? 'ok' : p('SspId'), termSetId: p('TermSetId') === set.id ? 'ok' : p('TermSetId'), textField: note ? `${note.InternalName} (hidden: ${note.Hidden})` : tf || null, hiddenList: /\sList="[^"]+"/.test(f.SchemaXml) };
  };
  const write = async (listUrl, single, multi) => {
    const run = async (title, values) => {
      const r = await fetch(`${listApi(listUrl)}/AddValidateUpdateItemUsingPath`, { method: 'POST', headers: W, body: JSON.stringify({
        listItemCreateInfo: { FolderPath: { DecodedUrl: listUrl }, UnderlyingObjectType: 0 },
        formValues: [{ FieldName: 'Title', FieldValue: title }].concat(Object.keys(values).map((k) => ({ FieldName: k, FieldValue: values[k] }))), bNewDocumentUpdate: false }) });
      const body = r.ok ? await r.json() : { value: [{ FieldName: 'HTTP', HasException: true, ErrorMessage: `${r.status} ${(await r.text()).slice(0, 120)}` }] };
      const id = Number(((body.value || []).filter((v) => v.FieldName === 'Id')[0] || {}).FieldValue);
      const errors = (body.value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`);
      const back = id ? await get(`${listApi(listUrl)}/items(${id})?$select=${single},${multi}`) : null;
      return { errors, back: back && { [single]: back[single], [multi]: back[multi] } };
    };
    return {
      T1_label_guid: await run('T1', { [single]: `Pénzügy|${a}`, [multi]: `Logisztika|${b};Raktár|${c}` }),
      T2_wssid: await run('T2', { [single]: `-1;#Pénzügy|${a}`, [multi]: `-1;#Logisztika|${b};#-1;#Raktár|${c}` }),
      T3_unknown_guid: await run('T3', { [single]: 'Pénzügy|00000000-1111-2222-3333-444444444444' })
    };
  };

  // C1: full SchemaXml, as SharePoint itself writes it – hidden note field first, then List, WebId and TextField.
  const prop = (n, v) => `<Property><Name>${n}</Name><Value xmlns:q1="http://www.w3.org/2001/XMLSchema" p4:type="q1:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">${v}</Value></Property>`;
  const createXml = async (listUrl, xml) => {
    const r = await fetch(`${listApi(listUrl)}/fields/createfieldasxml`, { method: 'POST', headers: W, body: JSON.stringify({ parameters: { SchemaXml: xml, Options: 12 } }) });
    return r.ok ? 'created' : `HTTP ${r.status}: ${(await r.text()).slice(0, 160)}`;
  };
  const l1 = await newList(`CJ11C1${stamp}`);
  if (l1 && hidden.Id) {
    const full = async (name, type, mult) => {
      const noteId = crypto.randomUUID(), fieldId = crypto.randomUUID();
      const note = await createXml(l1, `<Field Type="Note" ID="{${noteId}}" Name="${name}_0" StaticName="${name}_0" DisplayName="${name}_0" ShowInViewForms="FALSE" Required="FALSE" Hidden="TRUE" CanToggleHidden="TRUE" />`);
      const custom = prop('SspId', store.id) + prop('TermSetId', set.id) + prop('AnchorId', '00000000-0000-0000-0000-000000000000') + prop('TextField', `{${noteId}}`) + prop('IsPathRendered', 'false') + prop('IsKeyword', 'false') + prop('Open', 'false') + prop('CreateValuesInEditForm', 'false');
      const field = await createXml(l1, `<Field Type="${type}" ID="{${fieldId}}" Name="${name}" StaticName="${name}" DisplayName="${name}" ShowField="Term1033" List="{${hidden.Id}}" WebId="{${webInfo.Id}}" Mult="${mult}" Required="FALSE"><Default></Default><Customization><ArrayOfProperty>${custom}</ArrayOfProperty></Customization></Field>`);
      return { note, field };
    };
    out.C1_fullXml.single = await full('CJTaxC1', 'TaxonomyFieldType', 'FALSE');
    out.C1_fullXml.multi = await full('CJTaxMultiC1', 'TaxonomyFieldTypeMulti', 'TRUE');
    out.C1_fullXml.singleField = await describe(l1, 'CJTaxC1');
    out.C1_fullXml.multiField = await describe(l1, 'CJTaxMultiC1');
    out.C1_fullXml.values = await write(l1, 'CJTaxC1', 'CJTaxMultiC1');
  }

  // C2: CSOM, like PnP: AddFieldAsXml with a minimal field, then SspId / TermSetId / AnchorId and Update – one request.
  const l2 = await newList(`CJ11C2${stamp}`);
  const csom = async (listUrl, name, type, mult) => {
    const xml = `<Field Type="${type}" Name="${name}" StaticName="${name}" DisplayName="${name}" ShowField="Term1033" Mult="${mult}" Required="FALSE"><Default></Default></Field>`;
    const guid = (g) => `<Parameter Type="Guid">{${g}}</Parameter>`;
    const body =
      '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CopyJet">' +
      '<Actions><ObjectPath Id="10" ObjectPathId="5" />' +
      `<SetProperty Id="11" ObjectPathId="5" Name="SspId">${guid(store.id)}</SetProperty>` +
      `<SetProperty Id="12" ObjectPathId="5" Name="TermSetId">${guid(set.id)}</SetProperty>` +
      `<SetProperty Id="13" ObjectPathId="5" Name="AnchorId">${guid('00000000-0000-0000-0000-000000000000')}</SetProperty>` +
      '<Method Name="Update" Id="14" ObjectPathId="5" /></Actions><ObjectPaths>' +
      '<StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="1" Name="Web" />' +
      `<Method Id="3" ParentId="2" Name="GetList"><Parameters><Parameter Type="String">${esc(listUrl)}</Parameter></Parameters></Method>` +
      '<Property Id="4" ParentId="3" Name="Fields" />' +
      `<Method Id="5" ParentId="4" Name="AddFieldAsXml"><Parameters><Parameter Type="String">${esc(xml)}</Parameter><Parameter Type="Boolean">true</Parameter><Parameter Type="Enum">12</Parameter></Parameters></Method>` +
      '</ObjectPaths></Request>';
    const r = await fetch(`${web}/_vti_bin/client.svc/ProcessQuery`, { method: 'POST', headers: { 'Content-Type': 'text/xml', Accept: 'application/json', 'X-RequestDigest': digest }, body });
    const json = r.ok ? await r.json() : null;
    const err = json && json[0] && json[0].ErrorInfo;
    return r.ok ? (err ? `ErrorInfo: ${err.ErrorMessage} (${err.ErrorTypeName})` : 'created') : `HTTP ${r.status}`;
  };
  if (l2) {
    out.C2_csom.single = await csom(l2, 'CJTaxC2', 'TaxonomyFieldType', 'FALSE');
    out.C2_csom.multi = await csom(l2, 'CJTaxMultiC2', 'TaxonomyFieldTypeMulti', 'TRUE');
    out.C2_csom.singleField = await describe(l2, 'CJTaxC2');
    out.C2_csom.multiField = await describe(l2, 'CJTaxMultiC2');
    out.C2_csom.values = await write(l2, 'CJTaxC2', 'CJTaxMultiC2');
  }
  out.lists = [l1, l2].map((u) => (u || '').split('/').pop());
  const json = JSON.stringify(out, null, 2).split(location.host).join('contoso.sharepoint.com');
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

### C – eredmény (2026-09-30, azonos tenant)

- ❌ **C1, teljes XML:** a rejtett Note oszlop létrejött, de a taxonómia-oszlop **ismét HTTP 500** (`TargetInvocationException`), és utána a lista `fields` lekérdezése és minden elemírás is 500. XML-lel (`createFieldAsXml`) Managed Metadata oszlop **nem hozható létre**, még teljes leírással sem.
- ✅ **C2, CSOM:** `Web.GetList(url).Fields.AddFieldAsXml(<egyszerű TaxonomyFieldType>, true, 12)`, utána ugyanazon az objektumon `SetProperty SspId / TermSetId / AnchorId` és `Update()`, egyetlen `ProcessQuery`-ben. Egy- és többértékű is jó.
  - A `SspId` és a `TermSetId` a megadott.
  - **A rejtett jegyzetoszlopot (`TextField`) és a TaxonomyHiddenList-kötést (`List`) a SharePoint maga hozza létre.**
- **Értékek (`AddValidateUpdateItemUsingPath`):**
  - ✅ T1 `Címke|GUID`: egyértékű `Pénzügy|dae0…`, többértékű `Logisztika|3483…;Raktár|e8da…`. Visszaolvasva helyes (az egyértékű `Label`-je itt is a WssId).
  - ❌ T2 `-1;#Címke|GUID`: mezőszintű hiba („A címkézési felhasználói felületről visszaadott adatok formázása helytelen”), az elem nem jön létre.
  - ❌ T3 ismeretlen GUID: mezőszintű hiba („A megadott globálisan egyedi azonosító nem szerepel a kifejezéstárban”), **az elem nem jön létre**.

→ **Döntések:**
- **Oszlop:** csak CSOM-mal (`AddFieldAsXml` egyszerű XML-lel, utána `SspId`, `TermSetId`, `AnchorId` és `Update`), a site-oszlopoknál is. A cél termtár és termkészlet a sablonbeli „csoport/termkészlet” útvonal szerint oldódik fel. Ha nincs meg, az oszlop kimarad, figyelmeztetéssel, és **nem** jön létre félkész oszlop.
- **Érték:** `Címke|GUID`, többértékűnél `;`-vel elválasztva. **A nem feloldott termeket írás előtt el kell hagyni**, különben az egész elem elbukik.
- **Kiolvasás:**
  - a címke és az útvonal a termtárból jön a `TermGuid` alapján, nem a REST `Label`-ből;
  - az útvonalhoz a `children` hívásokon kell végigmenni, mert a lapos lista nem adja a szülőt.

### C – eredmény, másik tenant (2026-09-30)

Ugyanaz, mint az azonos tenantban:
- a C1 (teljes XML) 500-at ad, és tönkreteszi a listát;
- a C2 (CSOM) működik egy- és többértékű oszlopra is, a rejtett jegyzetoszlopot és a TaxonomyHiddenList-kötést a SharePoint hozza létre;
- a `Címke|GUID` alak helyes, a `-1;#…` és az ismeretlen GUID mezőszintű hibát ad.

A termkészletet a másik tenantban név szerint (csoport → termkészlet) találta meg, a termek GUID-ja ott más.

Állapot: **lezárva** – mindkét tenantban igazolva; a valódi telepítés ellenőrzi.
