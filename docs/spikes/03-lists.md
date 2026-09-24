# Spike 03 – Listák és tárak: szűrés, létrehozás URL-lel, keresés URL szerint

Állapot: **lezárva** (2026-09-24); a tár URL-je az első telepítési próbán igazolandó · Érintett kód: `src/core/lists/`, `ListExtractor`, `ListProvider`

## Kérdések

1. **Szűrés:** mely tulajdonságok választják el biztosan a felhasználói listákat és tárakat a rendszer-, katalógus- és segédlistáktól? A `Hidden eq false` a Forrás site-on 8 listát adott, és ez egyezett a *Webhely tartalma* nézettel (0. fázis). Ebből viszont nem derül ki, hogy mind a 8 átvihető-e. Jelöltek: `BaseTemplate`, `IsCatalog`, `IsApplicationList`, `IsSystemList`, `IsPrivate`, `IsSiteAssetsLibrary`, `NoCrawl`.
2. **Létrehozás megadott URL-lel:** a sablonban a lista azonosítója a site-relatív URL (pl. `Lists/Ugyfelek`), a cím ettől eltérhet (pl. „Ügyfelek”). Hogyan lehet ezt REST-ből létrehozni?
   - **B1:** `POST /_api/web/lists/add` hívás, `ListCreationInformation` paraméterben `Url` értékkel.
   - **B2:** a lista létrehozása az URL-névvel mint címmel, majd átnevezés (`MERGE Title`).
   - **B3:** ugyanez tárnál (101), ahol az URL nem `Lists/` alatt van.
3. **Keresés URL szerint:** a `GET /_api/web/getList('<szerver-relatív URL>')` nem létező listánál 404-et vagy `odata.null`-t ad? (A tartalomtípusoknál `odata.null` volt, spike 02.)

## Ideiglenes feltevések (a kód addig ezekre épül)

- **Szűrés:** ~~`Hidden eq false`, `BaseTemplate` ∈ {100, 101, 119}, és se nem `IsCatalog`, se nem `IsApplicationList`, se nem `IsSystemList`, se nem `IsSiteAssetsLibrary`.~~ Az A kísérlet alapján egyszerűsítve, lásd az Eredményt.
- **Kulcs:** a lista URL-jének utolsó szegmense, a kulcs-mintának megfelelően (`[A-Za-z0-9_.-]`, a szóköz `_` lesz). Ütközésnél `_2`, `_3` utótag kerül rá. Ugyanezt a függvényt használja a forrás-tokenkörnyezet (`{listkey:X}`) is.
- **Létrehozás:** a lista az URL-névvel mint címmel jön létre, utána kapja meg a valódi címét (B2). Ha a B1 működik, arra váltunk.
- **Keresés:** `getList(szerver-relatív URL)`, és mind a 404-et, mind az `odata.null`-t „nem létezik”-nek vesszük.

## A. kísérlet – Forrás site (csak olvas)

F12 → Console, beillesztés, Enter; ha kell, előtte `allow pasting`. A konzolba kiírt JSON-t küldd vissza.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const sel = 'Title,BaseTemplate,BaseType,Hidden,IsCatalog,IsApplicationList,IsSystemList,IsPrivate,IsSiteAssetsLibrary,NoCrawl,ItemCount,EntityTypeName,RootFolder/ServerRelativeUrl';
  const r = await fetch(`${web}/_api/web/lists?$select=${sel}&$expand=RootFolder`, { headers: H });
  const lists = (await r.json()).value || [];
  const flag = (l) => ['Hidden', 'IsCatalog', 'IsApplicationList', 'IsSystemList', 'IsPrivate', 'IsSiteAssetsLibrary', 'NoCrawl'].filter((k) => l[k]).join(',');
  const rows = lists.map((l) => `${l.Hidden ? '  ' : '* '}${l.Title} | ${l.RootFolder.ServerRelativeUrl} | tpl ${l.BaseTemplate} | ${l.ItemCount} elem | ${flag(l) || '-'}`);
  const json = JSON.stringify({ web, status: r.status, count: lists.length, visible: lists.filter((l) => !l.Hidden).length, rows }, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

A `*` jelöli a nem rejtett listákat. Érdemes előtte a Forrás site-on létrehozni egy ékezetes nevű listát, például „Ügyfelek”, hogy lássuk, milyen URL-t kap.

## B. kísérlet – Cél site (**ír**: legfeljebb 3 tesztlistát hoz létre)

Csak a Cél **teszt** site-on futtasd. Semmit nem töröl; a tesztlistákat (`CopyJet Spike …`) utána a *Webhely tartalma* oldalon kézzel törölheted. Ha egy tesztlista már létezik, a kód nem hozza létre újra.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,ServerRelativeUrl', { headers: H })).json();
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const short = async (r) => { try { return (await r.text()).slice(0, 300); } catch { return ''; } };
  const getList = async (url) => {
    const r = await fetch(`${web}/_api/web/getList('${encodeURIComponent(url)}')?$select=Id,Title&$expand=RootFolder`, { headers: H });
    const body = r.ok ? await r.json() : await short(r);
    return { status: r.status, title: body && body.Title, url: body && body.RootFolder && body.RootFolder.ServerRelativeUrl, raw: body && body.Id ? undefined : body };
  };
  const out = { web };

  // Q3: lookup by URL – missing and existing
  out.q3_missing = await getList(`${rel}/Lists/CopyJetNincsIlyenLista`);
  out.q3_existing = await getList(`${rel}/Shared Documents`);

  // B1: lists/add with ListCreationInformation.Url
  if ((await getList(`${rel}/Lists/CopyJetSpikeB1`)).title === undefined) {
    const r = await fetch(`${web}/_api/web/lists/add`, { method: 'POST', headers: W,
      body: JSON.stringify({ parameters: { Title: 'CopyJet Spike B1 ékezetes', Url: 'Lists/CopyJetSpikeB1', BaseTemplate: 100, Description: 'Törölhető teszt' } }) });
    out.b1_status = r.status; out.b1_body = r.ok ? '' : await short(r);
  }
  out.b1_result = await getList(`${rel}/Lists/CopyJetSpikeB1`);

  // B2: create with the URL name as title, then rename
  if ((await getList(`${rel}/Lists/CopyJetSpikeB2`)).title === undefined) {
    const r = await fetch(`${web}/_api/web/lists`, { method: 'POST', headers: W,
      body: JSON.stringify({ Title: 'CopyJetSpikeB2', BaseTemplate: 100, Description: 'Törölhető teszt' }) });
    out.b2_createStatus = r.status; out.b2_createBody = r.ok ? '' : await short(r);
    const m = await fetch(`${web}/_api/web/getList('${encodeURIComponent(rel + '/Lists/CopyJetSpikeB2')}')`, { method: 'POST',
      headers: { ...W, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body: JSON.stringify({ Title: 'CopyJet Spike B2 ékezetes' }) });
    out.b2_renameStatus = m.status; out.b2_renameBody = m.ok ? '' : await short(m);
  }
  out.b2_result = await getList(`${rel}/Lists/CopyJetSpikeB2`);

  // B3: document library via lists/add with Url (libraries live at the web root, not under Lists/)
  if ((await getList(`${rel}/CopyJetSpikeB3`)).title === undefined) {
    const r = await fetch(`${web}/_api/web/lists/add`, { method: 'POST', headers: W,
      body: JSON.stringify({ parameters: { Title: 'CopyJet Spike B3 tár', Url: 'CopyJetSpikeB3', BaseTemplate: 101, Description: 'Törölhető teszt' } }) });
    out.b3_status = r.status; out.b3_body = r.ok ? '' : await short(r);
  }
  out.b3_result = await getList(`${rel}/CopyJetSpikeB3`);

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## Mit várunk?

- **A:** egy vagy két ismérv, amely a nem rejtett, de nem átvihető listákat kiszűri (pl. Stíluskönyvtár, Űrlapsablonok, Webhelyeszközök).
- **B1:**
  - ha `b1_result.url` = `…/Lists/CopyJetSpikeB1` és `title` = „CopyJet Spike B1 ékezetes”, akkor az URL és a cím egy lépésben megadható;
  - ha nem, a B2 a tartalék.
- **B2:** `b2_result.url` = `…/Lists/CopyJetSpikeB2`, `title` = „CopyJet Spike B2 ékezetes”.
- **B3:** `b3_result.url` = `…/CopyJetSpikeB3` (tár a web gyökerében).
- **Q3:** mit ad a `q3_missing` (404 vagy 200 + `odata.null`), és a `q3_existing` megtalálja-e a Dokumentumok tárat.

## Eredmény

### A – szűrés (2026-09-24, Forrás site, magyar nyelvű)

21 lista van a site-on, ebből 8 nem rejtett:

| Lista | URL | Sablon | Jelzők | Átvihető? |
| --- | --- | --- | --- | --- |
| Dokumentumok | `Shared Documents` | 101 | – | ✅ |
| Események | `Lists/Events` | 106 | – | ⚠️ nem támogatott sablon (naptár) |
| Teszt lista | `Lists/Teszt lista` | 100 | – | ✅ |
| Teszt lookup forrás | `Lists/Teszt lookup forrs` | 100 | – | ✅ |
| Űrlapsablonok | `FormServerTemplates` | 101 | IsSystemList | ❌ |
| Webhelyeszközök | `SiteAssets` | 101 | IsApplicationList, IsSystemList, IsSiteAssetsLibrary | ❌ (a képek a 3. fázisban) |
| Weblapok | `SitePages` | 119 | IsApplicationList, IsSystemList | ❌ (a lapok a 3. fázisban) |
| Stílustár | `Style Library` | 101 | IsCatalog, IsSystemList | ❌ |

A 13 rejtett lista mind `IsSystemList`, ezek közül ráadásul sok `IsCatalog` is.

→ **Döntés (`isUserList`):**
- A feltérképezés azokat a listákat veszi fel, amelyek nem `Hidden`, nem `IsSystemList` és nem `IsCatalog`.
- Közülük a 100-as és a 101-es sablonú lista támogatott. A többi, például a 106-os naptár, *nem támogatott* jelzést kap, és kinyeréskor figyelmeztetéssel (`LIST_TEMPLATE_UNSUPPORTED`) kimarad.
- A 119-es sablont (Site Pages) a séma megengedi, de rendszerlista, ezért a lista-pár nem viszi át. A lapokat a 3. fázis kezeli.

**URL és cím:** a „Teszt lookup forrás” lista URL-je `Lists/Teszt lookup forrs`, mert a SharePoint az ékezetes betűt kihagyta. A „Teszt lista” URL-jében szóköz van. A lista azonosítója ezért az URL, nem a cím. A kulcs az utolsó URL-szegmensből képződik, pl. `Teszt_lookup_forrs` vagy `Shared_Documents`.

### B – létrehozás és keresés (2026-09-24, Cél site)

- **Q3 – nem létező lista:** a `getList('<url>')` **HTTP 404**-et ad (`System.IO.FileNotFoundException`), nem `odata.null`-t, szemben a tartalomtípusokkal. A meglévő Dokumentumok tárat (`/sites/Cel/Shared Documents`) megtalálta.
- **B2 – létrehozás az URL-névvel mint címmel, majd átnevezés:** ✅ működik.
  - A `POST /_api/web/lists` `{ Title: "CopyJetSpikeB2", BaseTemplate: 100 }` törzzsel 201-et adott, a lista URL-je `/sites/Cel/Lists/CopyJetSpikeB2` lett.
  - Ezután a `MERGE { Title: "CopyJet Spike B2 ékezetes" }` 204-et adott. A cím megváltozott, az URL maradt.
- **B1/B3 – nem értékelhető, a kísérleti kód hibája miatt:** a `ListCreationInformation`-ben a sablon mezőneve `TemplateType`, nem `BaseTemplate`. Emiatt mindkét hívás HTTP 400-at adott: „The property 'BaseTemplate' does not exist on type 'SP.ListCreationInformation'”. A tárat (101) így nem teszteltük.

→ **Döntés:**
- A `ListProvider` a B2 utat használja: létrehozás az URL utolsó szegmensével mint címmel, majd egy `MERGE` a valódi címmel és a beállításokkal.
- Létrehozás után a lista URL-jét ellenőrzi, és eltérésnél `LIST_URL_MISMATCH` hibával leáll. Ez fedi le a tárak nem tesztelt esetét is.
- A listát URL szerint keresi, a 404-et „nem létezik”-nek veszi.

**Takarítás:** a Cél site-on a `CopyJet Spike B2 ékezetes` lista kézzel törölhető. A B1 és a B3 nem jött létre.
