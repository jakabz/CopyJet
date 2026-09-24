# Spike 06 – Nézetek: kiolvasás, létrehozás, mezők, alapértelmezett nézet

Állapot: **lezárva** (2026-09-24) · Érintett kód: `ViewExtractor`, `ViewProvider`

## Kérdések

1. **Kiolvasás:** mely tulajdonságok írják le a nézetet (`ViewQuery`, `RowLimit`, `Paged`, `Scope`, `ViewType`, `ViewType2`, `CustomFormatter`, mezők)? Hogyan szűrhetők ki a rejtett és a személyes nézetek?
2. **URL és cím:** milyen URL-t kap egy ékezetes című nézet? A sémában a nézetnek nincs URL-mezője.
   - **Terv:** az alapnézetet a `default` jelző szerint párosítjuk (a címe nyelvfüggő: „Minden elem” / „All Items”), a többit cím szerint.
3. **Létrehozás:** a `POST <lista>/views` hívás elfogadja-e létrehozáskor a `ViewQuery`, `RowLimit`, `Paged` és `ViewTypeKind` (GRID = 2048) értékeket?
4. **Mezők:** működik-e a `viewfields/removeallviewfields` + `viewfields/addviewfield('<név>')` páros?
5. **Alapértelmezett és hatókör:** beállítható-e `MERGE`-dzsel a `DefaultView: true` és a `Scope`?

## A. kísérlet – Forrás site (csak olvas)

**Előkészítés a Forrás site-on, a „Teszt lista” listán:**
1. Hozz létre egy nyilvános nézetet **„Nyitott elemek”** néven, szűrővel (pl. `ListaSzoveg` nem üres) és egy rendezéssel.
2. Hozz létre egy másik nyilvános nézetet **„Ügyfél nézet”** néven, ékezetes címmel, néhány oszloppal.
3. Ha szeretnéd, készíts egy személyes nézetet is, hogy lássuk, kimarad-e.

Utána: F12 → Console → beillesztés → Enter, és küldd vissza a kiírt JSON-t.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const out = { web };
  for (const url of ['Lists/Teszt lista', 'Shared Documents']) {
    const LIST = `${web}/_api/web/getList('${enc(rel + '/' + url)}')`;
    const sel = 'Id,Title,ServerRelativeUrl,DefaultView,Hidden,PersonalView,ViewType,ViewType2,BaseViewId,RowLimit,Paged,Scope,ViewQuery,CustomFormatter,TabularView,MobileView';
    const views = (await (await fetch(`${LIST}/views?$select=${sel}`, { headers: H })).json()).value || [];
    out[url] = [];
    for (const v of views) {
      const vf = await (await fetch(`${LIST}/views('${v.Id}')/viewfields`, { headers: H })).json();
      out[url].push({
        title: v.Title, url: v.ServerRelativeUrl.replace(rel + '/', ''), default: v.DefaultView, hidden: v.Hidden, personal: v.PersonalView,
        type: v.ViewType, type2: v.ViewType2, baseViewId: v.BaseViewId, rowLimit: v.RowLimit, paged: v.Paged, scope: v.Scope,
        query: v.ViewQuery, fields: vf.Items, formatterLength: (v.CustomFormatter || '').length, tabular: v.TabularView
      });
    }
  }
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## B. kísérlet – Cél site (**ír**: új tesztlista és 2 nézet)

Csak a Cél **teszt** site-on futtasd. Semmit nem töröl, és nem függ korábbi tesztelemektől. Létrehoz egy `CopyJetSpike06` listát, és ezen két nézetet:
- „CJ Nézet ékezetes” (HTML, lekérdezéssel, lapozással);
- „CJ Rács” (GRID).

Az első nézet mezőit lecseréli, a nézetet alapértelmezetté teszi, és beállítja a hatókörét. A teszt után a lista kézzel törölhető.

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
  const LIST = `${web}/_api/web/getList('${enc(rel + '/Lists/CopyJetSpike06')}')`;
  const out = { web };

  if ((await fetch(`${LIST}?$select=Id`, { headers: H })).status === 404) {
    out.createList = (await fetch(`${web}/_api/web/lists`, { method: 'POST', headers: W, body: JSON.stringify({ Title: 'CopyJetSpike06', BaseTemplate: 100 }) })).status;
  }
  const viewBy = async (title) => ((await (await fetch(`${LIST}/views?$select=Id,Title,ServerRelativeUrl,DefaultView,ViewType,RowLimit,Paged,Scope,ViewQuery&$filter=Title eq '${title.replace(/'/g, "''")}'`, { headers: H })).json()).value || [])[0];
  const describe = async (v) => v && ({ ...v, ServerRelativeUrl: v.ServerRelativeUrl.replace(rel + '/', ''), fields: (await (await fetch(`${LIST}/views('${v.Id}')/viewfields`, { headers: H })).json()).Items });

  // Q3: HTML view with query, row limit and paging at creation
  const q = '<OrderBy><FieldRef Name="Created" Ascending="FALSE" /></OrderBy><Where><IsNotNull><FieldRef Name="Title" /></IsNotNull></Where>';
  if (!(await viewBy('CJ Nézet ékezetes'))) {
    const r = await fetch(`${LIST}/views`, { method: 'POST', headers: W, body: JSON.stringify({ Title: 'CJ Nézet ékezetes', PersonalView: false, ViewQuery: q, RowLimit: 50, Paged: true }) });
    out.q3_status = r.status; out.q3_body = r.ok ? '' : await short(r);
  }
  // Q3b: GRID view via ViewTypeKind at creation
  if (!(await viewBy('CJ Rács'))) {
    const r = await fetch(`${LIST}/views`, { method: 'POST', headers: W, body: JSON.stringify({ Title: 'CJ Rács', PersonalView: false, ViewTypeKind: 2048, RowLimit: 30 }) });
    out.q3b_status = r.status; out.q3b_body = r.ok ? '' : await short(r);
  }

  const v = await viewBy('CJ Nézet ékezetes');
  if (v) {
    // Q4: replace fields
    let r = await fetch(`${LIST}/views('${v.Id}')/viewfields/removeallviewfields`, { method: 'POST', headers: W });
    out.q4_removeStatus = r.status;
    for (const name of ['LinkTitle', 'Created', 'Author']) {
      r = await fetch(`${LIST}/views('${v.Id}')/viewfields/addviewfield('${name}')`, { method: 'POST', headers: W });
      out['q4_add_' + name] = r.status;
    }
    // Q5: default view + scope (1 = Recursive)
    r = await fetch(`${LIST}/views('${v.Id}')`, { method: 'POST', headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ DefaultView: true, Scope: 1 }) });
    out.q5_status = r.status; out.q5_body = r.ok ? '' : await short(r);
  }
  out.view = await describe(await viewBy('CJ Nézet ékezetes'));
  out.grid = await describe(await viewBy('CJ Rács'));
  out.allItemsDefault = ((await (await fetch(`${LIST}/views?$select=Title,DefaultView,ServerRelativeUrl`, { headers: H })).json()).value || []).map((x) => `${x.Title} | ${x.ServerRelativeUrl.replace(rel + '/', '')}${x.DefaultView ? ' (alapértelmezett)' : ''}`);

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## Mit várunk?

- **A:** a nyilvános nézetek `hidden=false`, `personal=false`; a „Minden elem” alapnézet URL-je `AllItems.aspx`. Az „Ügyfél nézet” URL-jéből kiderül, mit kezd a SharePoint az ékezetekkel.
- **B:**
  - `view.RowLimit` = 50, `Paged` = true, a `ViewQuery` a kért lekérdezés, a mezők `LinkTitle, Created, Author`, `DefaultView` = true, `Scope` = 1;
  - `grid.ViewType` = `GRID`;
  - az `allItemsDefault` listában a „Minden elem” már nem alapértelmezett.

## Eredmény

### A – kiolvasás (2026-09-24, Forrás site)

| Lista | Nézet | URL | Jelzők |
| --- | --- | --- | --- |
| Teszt lista | Minden elem | `AllItems.aspx` | alapértelmezett |
| Teszt lista | Nyitott elemek | `Nyitott elemek.aspx` | lekérdezéssel (`OrderBy` + `Where IsNotNull`) |
| Teszt lista | Ügyfél nézet | `gyfl nzet.aspx` | – |
| Teszt lista | Személyes nézet | `PersonalViews.aspx` | **personal** |
| Dokumentumok | Minden dokumentum | `Forms/AllItems.aspx` | alapértelmezett |
| Dokumentumok | Dokumentumok újbóli csatolása, assetLibTemp, Dokumentumok egyesítése | `Forms/repair.aspx`, `Thumbnails.aspx`, `Combine.aspx` | **hidden** |

- **Szűrés:** a másolandó nézetek azok, amelyek nem `Hidden` és nem `PersonalView`.
- **URL:** az ékezetes betűk kimaradnak („Ügyfél nézet” → `gyfl nzet.aspx`), mint a listáknál. Az alapnézet címe nyelvfüggő, az URL-je `AllItems.aspx`.
  → A párosítás: az alapnézet a `default` jelző szerint, a többi cím szerint. A sémát nem kell bővíteni.
- **Tartalom:**
  - A mezők belső névvel szerepelnek (`LinkTitle`, `V_x00e1_lassz`, `Lookup` …).
  - A `ViewQuery` egyszerű CAML.
  - A `RowLimit` 30, a `Paged` igaz, a `Scope` 0, a `ViewType` HTML, a `ViewType2` üres.
  - `CustomFormatter` egyik nézetben sem volt.

### B – létrehozás (2026-09-24, Cél site)

- **Q3 – HTML nézet létrehozása: ✅.** A `POST <lista>/views` hívás `{ Title, PersonalView: false, ViewQuery, RowLimit: 50, Paged: true }` törzzsel 201-et adott, és minden érték megmaradt.
  - Az URL a címből képződik, **ugyanazzal az ékezetkihagyással**, mint a forráson: „CJ Nézet ékezetes” → `CJ Nzet kezetes.aspx`. A cím szerinti létrehozás tehát a forrással azonos URL-t ad.
- **Q3b – GRID nézet létrehozáskor: ❌.** A hiba: „The property 'ViewTypeKind' does not exist on type 'SP.View'”, HTTP 400. A nézettípus a `ViewCreationInformation` része (`views/add`), ezt nem teszteltük.
- **Q4 – mezők cseréje: ✅.** A `viewfields/removeallviewfields`, majd az `addviewfield('LinkTitle' | 'Created' | 'Author')` hívás 200-at adott, az eredmény pontosan ez a sorrend.
- **Q5 – alapértelmezett és hatókör: ✅.** A `MERGE { DefaultView: true, Scope: 1 }` 204-et adott. A nézet alapértelmezett lett, a „Minden elem” már nem az, a `Scope` 1.

→ **Döntés (`ViewProvider`):**
- **Létrehozás:** `views.add(title, false, { ViewQuery, RowLimit, Paged })`, majd a mezők cseréje, végül egy `MERGE` a hatókörrel, a formázással és szükség esetén a `DefaultView`-val.
- **Párosítás:** az alapnézet a célon a `DefaultView` szerint, a többi cím szerint.
- **GRID és CALENDAR nézet:** az 1. verzióban figyelmeztetéssel kimarad (`VIEW_TYPE_UNSUPPORTED`). Az A kísérletben nem is fordult elő ilyen.

**Takarítás:** a Cél site-on a `CopyJetSpike06` lista kézzel törölhető.
