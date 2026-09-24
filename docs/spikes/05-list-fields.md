# Spike 05 – Lista mezők: saját oszlop vagy site column, létrehozás listán

Állapot: **lezárva** (2026-09-24) · Érintett kód: `ListFieldExtractor`, `ListFieldProvider`

## Kérdések

1. **Felismerés:** egy lista mezői közül melyik a lista saját oszlopa, melyik egy site column listára került példánya, és melyik beépített?
   - **Feltevés:** a `SchemaXml` `SourceID` attribútuma dönt. A saját oszlopnál ez a lista GUID-ja, a site column példányánál a web GUID-ja, a beépítettnél egy `http://schemas.microsoft.com/…` URI.
2. **Létrehozás listán:** a `POST <lista>/fields/createfieldasxml` hívás `Options: 12` beállítással (`AddFieldInternalNameHint` + `AddToAllContentTypes`) létrehozza-e a mezőt? Megtartja-e a belső nevet és az `ID`-t? Lookup mezőnél elfogadja-e a `List="{<GUID>}"` hivatkozást?
3. **Site column listára tétele tartalomtípus nélkül:** ha a site column `SchemaXml`-jét (azonos `ID`-val) a lista `createfieldasxml`-jével adjuk hozzá, a lista mezője a site columnhoz kötött példány lesz-e? Ugyanazt az ID-t kapja-e?

## A. kísérlet – Forrás site (csak olvas)

**Előkészítés a Forrás site-on, a „Teszt lista” listán:**
1. *+ Oszlop hozzáadása* → **Szöveg**, neve: `ListaSzoveg`.
2. *+ Oszlop hozzáadása* → **Keresés** (Lookup), neve: `ListaLookup`, forrása a „Teszt lookup forrás” lista Cím oszlopa.
3. *Listabeállítások* → *Hozzáadás meglévő webhelyoszlopokból* → `SiteColumn1`.

Utána futtasd a Forrás site bármely lapján: F12 → Console → beillesztés → Enter. A kiírt JSON-t küldd vissza.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,Id,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const out = { web, webId: webInfo.Id };
  for (const url of ['Lists/Teszt lista', 'Lists/Teszt lookup forrs']) {
    const LIST = `${web}/_api/web/getList('${enc(rel + '/' + url)}')`;
    const list = await (await fetch(`${LIST}?$select=Id,Title`, { headers: H })).json();
    const fields = (await (await fetch(`${LIST}/fields?$select=Id,InternalName,TypeAsString,Hidden,CanBeDeleted,FromBaseType,Sealed,ReadOnlyField,SchemaXml&$filter=Hidden eq false`, { headers: H })).json()).value || [];
    out[url] = {
      listId: list.Id,
      fields: fields.map((f) => {
        const el = new DOMParser().parseFromString(f.SchemaXml, 'application/xml').documentElement;
        const src = (el.getAttribute('SourceID') || '').replace(/[{}]/g, '').toLowerCase();
        const kind = src === String(list.Id).toLowerCase() ? 'LISTA' : src === String(webInfo.Id).toLowerCase() ? 'WEB' : /^http/.test(src) ? 'beépített' : (src || '-');
        return `${f.InternalName} | ${f.TypeAsString} | ${kind} | CanBeDeleted=${f.CanBeDeleted} FromBaseType=${f.FromBaseType} ReadOnly=${f.ReadOnlyField}` +
          (f.TypeAsString.indexOf('Lookup') === 0 ? ` | List=${el.getAttribute('List')} ShowField=${el.getAttribute('ShowField')}` : '');
      })
    };
  }
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## B. kísérlet – Cél site (**ír**: új tesztlista és 3 mező)

Csak a Cél **teszt** site-on futtasd. Semmit nem töröl, és nem függ korábbi tesztelemektől. Létrehoz egy `CopyJet Spike 05` listát (`Lists/CopyJetSpike05`), és ezen:
- **Q2a:** egy saját szöveges oszlopot, megadott `ID`-val;
- **Q2b:** egy lookup oszlopot, amely a lista saját Címére mutat;
- **Q3:** a beépített `StartDate` site column példányát, a site column saját `SchemaXml`-jéből.

A teszt után a lista kézzel törölhető.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,Id,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const short = async (r) => { try { return (await r.text()).slice(0, 300); } catch { return ''; } };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const listUrl = `${rel}/Lists/CopyJetSpike05`;
  const LIST = `${web}/_api/web/getList('${enc(listUrl)}')`;
  const out = { web };

  let list = await fetch(`${LIST}?$select=Id`, { headers: H });
  if (list.status === 404) {
    const r = await fetch(`${web}/_api/web/lists`, { method: 'POST', headers: W, body: JSON.stringify({ Title: 'CopyJetSpike05', BaseTemplate: 100 }) });
    out.createList = r.status;
    list = await fetch(`${LIST}?$select=Id`, { headers: H });
  }
  const listId = (await list.json()).Id;
  out.listId = listId;
  const field = async (name) => {
    const r = await fetch(`${LIST}/fields?$select=Id,InternalName,TypeAsString,SchemaXml&$filter=InternalName eq '${name}'`, { headers: H });
    const f = ((await r.json()).value || [])[0];
    if (!f) return null;
    const el = new DOMParser().parseFromString(f.SchemaXml, 'application/xml').documentElement;
    return { id: f.Id, type: f.TypeAsString, sourceId: el.getAttribute('SourceID'), list: el.getAttribute('List') };
  };
  const create = async (key, schemaXml, options) => {
    const r = await fetch(`${LIST}/fields/createfieldasxml`, { method: 'POST', headers: W, body: JSON.stringify({ parameters: { SchemaXml: schemaXml, Options: options } }) });
    out[key + '_status'] = r.status;
    out[key + '_body'] = r.ok ? '' : await short(r);
  };

  // Q2a: own text column with a given ID
  if (!(await field('CJSpikeText'))) {
    await create('q2a', '<Field ID="{c0ffee05-0000-4000-8000-000000000001}" Name="CJSpikeText" StaticName="CJSpikeText" DisplayName="CJ Spike szöveg" Type="Text" />', 12);
  }
  out.q2a_field = await field('CJSpikeText');

  // Q2b: lookup to the list itself (Title)
  if (!(await field('CJSpikeLookup'))) {
    await create('q2b', `<Field ID="{c0ffee05-0000-4000-8000-000000000002}" Name="CJSpikeLookup" StaticName="CJSpikeLookup" DisplayName="CJ Spike lookup" Type="Lookup" List="{${listId}}" ShowField="Title" />`, 12);
  }
  out.q2b_field = await field('CJSpikeLookup');

  // Q3: a site column's own SchemaXml added to the list
  const site = ((await (await fetch(`${web}/_api/web/availablefields?$select=Id,SchemaXml&$filter=InternalName eq 'StartDate'`, { headers: H })).json()).value || [])[0];
  out.q3_siteColumnId = site ? site.Id : null;
  if (site && !(await field('StartDate'))) {
    await create('q3', site.SchemaXml, 8);
  }
  out.q3_field = await field('StartDate');

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## Mit várunk?

- **A:** a `ListaSzoveg` és a `ListaLookup` `LISTA`, a `SiteColumn1` `WEB`, a Title/Created stb. `beépített`. A lookup `List` attribútuma a „Teszt lookup forrás” lista GUID-ja.
- **B:**
  - `q2a_field.id` = `c0ffee05-…-000000000001`, és a típus Text;
  - `q2b_field.list` a lista saját GUID-ja;
  - `q3_field.id` = `q3_siteColumnId`, a `sourceId` a web GUID-ja vagy a beépített séma-URI. Ez azt jelenti, hogy a site column példánya azonos ID-val került a listára.

## Eredmény

### A – felismerés (2026-09-24, Forrás site): ✅ a feltevés igaz

**„Teszt lista”** (lista ID: `f1768c8f-…`, web ID: `8b98bd9e-…`):

| Csoport | Mezők | `SourceID` |
| --- | --- | --- |
| Saját oszlop (LISTA) | `Valaki` (User), `V_x00e1_lassz` (Choice), `Lookup`, `ListaSzoveg` (Text), `ListaLookup` | a lista GUID-ja |
| Site column példány (WEB) | `SiteColumn1` (Text), `SiteColumn2` (Number) | a web GUID-ja |
| Beépített | Title, ID, Created, Author, Attachments, Compliance-mezők stb. | `http://schemas.microsoft.com/…` |

- A beépítettek mind `CanBeDeleted=false`, `FromBaseType=true` értékűek, de az ismérv a `SourceID` (egyértelmű, és a site columnoknál is ez vált be, spike 01).
- A lookup `List` attribútuma a céllista GUID-ja kapcsos zárójelben (`{74c85f18-…}`, a „Teszt lookup forrás”). A tokenizálás után ebből `{{listkey:Teszt_lookup_forrs}}` lesz.
- A belső név kódolt ékezetet tartalmazhat (`V_x00e1_lassz` = „Válassz”). Ez megfelel a séma mintájának, és változatlanul kell átvinni.

→ **Döntés:**
- A `ListFieldExtractor` a LISTA csoportot a teljes definícióval viszi át (tisztított, tokenizált `SchemaXml`).
- A WEB csoportot hivatkozásként viszi át: a site column belső nevét és ID-ját. A cél site-on a site column már létezik, mert a site columnok a listák előtt települnek.
- A beépített mezőket nem viszi át.

### B – létrehozás listán (2026-09-24, Cél site): ✅

A `POST <lista>/fields/createfieldasxml` mindhárom esetben HTTP 200-at adott:

| Eset | Kért | Eredmény |
| --- | --- | --- |
| Q2a – saját szöveges oszlop, `Options: 12` | `ID` `c0ffee05-…-01` | ID megmaradt; a `SourceID`-t a SharePoint a lista GUID-jára állította |
| Q2b – lookup a lista saját Címére, `Options: 12` | `List="{<lista GUID>}"` | létrejött, a `List` a megadott GUID |
| Q3 – a `StartDate` site column saját `SchemaXml`-je, `Options: 8` | site column ID `64cd368d-…` | **azonos ID**, `SourceID` = a site columné (`http://schemas.microsoft.com/sharepoint/v3`), vagyis a site columnhoz kötött példány |

→ **Döntés (`ListFieldProvider`):**
- A sablonban minden lista mező `IField`, a sémát nem bővítjük.
- Ha a mező ID-jával **van** site column a cél site-on, a példány a cél site column saját `SchemaXml`-jéből jön létre (Q3). Ez a `SchemaXml` nincs tisztítva, mert a `SourceID` köti a site columnhoz.
- Ha **nincs**, a mező saját oszlopként jön létre a sablon tisztított, feloldott `SchemaXml`-jéből (Q2).
- Mindkét úton `Options: 12` (`AddFieldInternalNameHint` + `AddToAllContentTypes`), hogy a mező a lista űrlapjain is megjelenjen, ahogy a felületen hozzáadott oszlopoknál. A Q3 esetet `Options: 8`-cal teszteltük; a 12-es értéket site column példányra az első telepítési próba igazolja.

**Takarítás:** a Cél site-on a `CopyJetSpike05` lista kézzel törölhető.
