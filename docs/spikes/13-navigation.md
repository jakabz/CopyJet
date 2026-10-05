# Spike 13 – Navigáció és kezdőlap

Állapot: **lezárva** · Érintett kód (3. fázis): `NavigationExtractor`, `NavigationProvider`, kezdőlap-beállítás

## Kérdések

1. **Kiolvasás:**
   - milyen alakban adja a REST a bal oldali (`QuickLaunch`) és a felső (`TopNavigationBar`) menüt;
   - kibontható-e több szint egy kérésben;
   - mit ad hozzá a `MenuState` API (célközönség, rejtett csomópont, mélyebb szintek)?
2. **Site-függő értékek:** milyen URL-ek vannak a csomópontokon (abszolút, relatív, lista, lap, külső)? Mit kell tokenizálni?
3. **Beállítások:** megosztott navigáció, vízszintes bal menü, megamenü, `QuickLaunchEnabled`. A kommunikációs site felső menüje melyik gyűjteményben van?
4. **Kezdőlap:** honnan olvasható (`rootfolder/WelcomePage`), és hogyan írható?
5. **Írás** (B kísérlet, a Célon):
   - csomópont és alcsomópont felvétele;
   - abszolút és relatív URL kezelése, külső hivatkozás;
   - a kezdőlap átállítása, majd visszaállítása.

## Előkészítés (Forrás)

A Forrás site bal oldali menüjébe vegyél fel (Szerkesztés a menü alján):
- egy **„CopyJet teszt”** fejlécet, alatta két hivatkozással: a **„Teszt lista”** listára és a **`CopyJet-teszt.aspx`** lapra;
- egy **külső hivatkozást** (pl. `https://www.microsoft.com`);
- ha a site-on van felső menü (kommunikációs site vagy bekapcsolt vízszintes menü), oda is egy hivatkozást;
- ha van kedved: az egyik hivatkozásnak célközönséget.

## A. kísérlet – Forrás site (csak olvas)

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (url) => { const r = await fetch(url, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status, error: (await r.text()).slice(0, 200) }; };
  const webInfo = await get(guess + '/_api/web?$select=Id,Url,ServerRelativeUrl,WebTemplate,Configuration,QuickLaunchEnabled,HorizontalQuickLaunch,MegaMenuEnabled,NavAudienceTargetingEnabled');
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const out = { web: { template: `${webInfo.WebTemplate}#${webInfo.Configuration}`, quickLaunchEnabled: webInfo.QuickLaunchEnabled, horizontalQuickLaunch: webInfo.HorizontalQuickLaunch, megaMenu: webInfo.MegaMenuEnabled, audienceTargeting: webInfo.NavAudienceTargetingEnabled } };
  const nav = await get(`${web}/_api/web/navigation?$select=UseShared`);
  out.web.useShared = nav.httpStatus ? nav : nav.UseShared;
  const shortUrl = (u) => (u || '').replace(web, '{site}').replace(rel + '/', '{siterelative}/');
  // 1. REST collections, one level of children expanded.
  const tree = async (name) => {
    const r = await get(`${web}/_api/web/navigation/${name}?$expand=Children,Children/Children&$select=Id,Title,Url,IsExternal,IsDocLib,IsVisible,AudienceIds,Children/Id,Children/Title,Children/Url,Children/IsExternal,Children/Children/Title,Children/Children/Url`);
    if (r.httpStatus) return r;
    return (r.value || []).map((n) => ({ id: n.Id, title: n.Title, url: shortUrl(n.Url), ext: n.IsExternal, docLib: n.IsDocLib, visible: n.IsVisible, audiences: n.AudienceIds, children: (n.Children || []).map((c) => ({ id: c.Id, title: c.Title, url: shortUrl(c.Url), ext: c.IsExternal, children: (c.Children || []).map((g) => `${g.Title} → ${shortUrl(g.Url)}`) })) }));
  };
  out.quickLaunch = await tree('QuickLaunch');
  out.topNavigationBar = await tree('TopNavigationBar');
  // 2. MenuState: the tree the modern UI edits (keys, hidden nodes, audiences, deeper levels).
  const menu = async (key) => {
    const r = await get(`${web}/_api/navigation/MenuState?menuNodeKey='${key}'&mapProviderName='SPNavigationProvider'&depth=5`);
    if (r.httpStatus) return r;
    const walk = (n) => ({ key: n.Key, title: n.Title, url: shortUrl(n.SimpleUrl), hidden: n.IsHidden, deleted: n.IsDeleted, nodeType: n.NodeType, audiences: n.AudienceIds, children: (n.Nodes || []).map(walk) });
    return { startingNodeKey: r.StartingNodeKey, startingNodeTitle: r.StartingNodeTitle, simpleUrl: shortUrl(r.SimpleUrl), version: r.Version, nodes: (r.Nodes || []).map(walk), keys: Object.keys(r) };
  };
  out.menuStateQuickLaunch = await menu('1025');
  out.menuStateTopNav = await menu('1002');
  // 3. Home page.
  const root = await get(`${web}/_api/web/rootfolder?$select=WelcomePage`);
  out.welcomePage = root.httpStatus ? root : root.WelcomePage;
  // 4. Which node URLs point at lists of this site (for {listurl:K} tokens)?
  const lists = (await get(`${web}/_api/web/lists?$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false`)).value || [];
  const allUrls = JSON.stringify([out.quickLaunch, out.topNavigationBar]);
  out.listUrlsInNav = lists.map((l) => ({ title: l.Title, url: shortUrl(l.RootFolder.ServerRelativeUrl) })).filter((l) => allUrls.indexOf(l.url) >= 0);
  console.log(JSON.stringify(out, null, 2));
})();
```

**Futtatás:** a Forrás site bármelyik lapján F12 → Console, másold be, Enter. A kimenő JSON-t kérem vissza.

## B. kísérlet – Cél site (ír, majd visszaállít)

A kísérlet a Cél site bal oldali menüjébe ír:
- felvesz egy **„CopyJet spike13”** csomópontot két alcsomóponttal: egy relatív és egy abszolút URL-lel megadott belső hivatkozással;
- felvesz egy külső hivatkozást;
- visszaolvassa, mit tárolt el a SharePoint;
- a kezdőlapot átállítja a `CopyJet-teszt.aspx` lapra (ezt a 3. fázis telepítése már létrehozta), visszaolvassa, majd visszaállítja az eredetire;
- a külső hivatkozásra REST-tel célközönséget ír (a Forrás tesztjének csoportjával; a Cél ugyanabban a tenantban van), és `MenuState`-tel visszaolvassa;
- a végén törli a saját tesztcsomópontjait. Ha a menüben látni szeretnéd őket, írd át a `CLEANUP` értékét `false`-ra; ekkor utána kézzel kell törölnöd őket.

```js
(async () => {
  const CLEANUP = true;
  const PAGE = 'SitePages/CopyJet-teszt.aspx';
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const ctx = await (await fetch(guess + '/_api/contextinfo', { method: 'POST', headers: { Accept: 'application/json;odata=nometadata' } })).json();
  const H = { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': ctx.FormDigestValue };
  const call = async (method, url, body, extra) => {
    const r = await fetch(url, { method, headers: Object.assign({}, H, extra || {}), body: body ? JSON.stringify(body) : undefined });
    const text = await r.text();
    if (!r.ok) return { httpStatus: r.status, error: text.slice(0, 300) };
    try { return text ? JSON.parse(text) : { ok: true }; } catch { return { ok: true }; }
  };
  const webInfo = await call('GET', guess + '/_api/web?$select=Url,ServerRelativeUrl');
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const out = {};
  const ql = `${web}/_api/web/navigation/QuickLaunch`;
  // 1. Heading without link (the UI stores http://linkless.header/, spike 13 A) + children: relative, absolute, external.
  const head = await call('POST', ql, { Title: 'CopyJet spike13', Url: 'http://linkless.header/', IsExternal: true });
  out.heading = head.httpStatus ? head : { id: head.Id, storedUrl: head.Url };
  const created = [];
  if (!head.httpStatus) {
    created.push(head.Id);
    const kids = `${ql}/GetById(${head.Id})/Children`;
    const relChild = await call('POST', kids, { Title: 'Relatív', Url: `${rel}/${PAGE}`, IsExternal: false });
    const absChild = await call('POST', kids, { Title: 'Abszolút', Url: `${web}/${PAGE}`, IsExternal: false });
    const extChild = await call('POST', kids, { Title: 'Külső', Url: 'https://www.microsoft.com', IsExternal: true });
    const notFound = await call('POST', kids, { Title: 'Nincs ilyen lap', Url: `${rel}/SitePages/Nincs-ilyen.aspx`, IsExternal: false });
    out.children = { relChild, absChild, extChild, notFound };
    Object.keys(out.children).forEach((k) => { const c = out.children[k]; out.children[k] = c.httpStatus ? c : { id: c.Id, storedUrl: c.Url, isExternal: c.IsExternal }; });
    // Read back the way the extractor will read.
    const back = await call('GET', `${ql}/GetById(${head.Id})?$expand=Children&$select=Title,Url,Children/Title,Children/Url,Children/IsExternal`);
    out.readBack = back.httpStatus ? back : { url: back.Url, children: (back.Children || []).map((c) => `${c.Title} → ${c.Url} (ext ${c.IsExternal})`) };
    // Order: the new heading is last? (append semantics)
    const order = await call('GET', `${ql}?$select=Title`);
    out.quickLaunchOrder = (order.value || []).map((n) => n.Title);
  }
  // 1b. Audience on the external child through REST MERGE, read back through MenuState (only it returns audiences).
  if (!head.httpStatus && out.children.extChild.id) {
    const AUDIENCE = 'a64a6140-20c7-44c7-ae78-754d724f63e8'; // the group of the Forrás audience test (same tenant)
    out.audienceMerge = await call('POST', `${ql}/GetById(${out.children.extChild.id})`, { AudienceIds: [AUDIENCE] }, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
    const ms = await call('GET', `${web}/_api/navigation/MenuState?menuNodeKey='1025'&mapProviderName='SPNavigationProvider'&depth=5`);
    const mine = (ms.Nodes || []).filter((n) => n.Key === String(head.Id))[0];
    out.menuStateReadBack = mine ? (mine.Nodes || []).map((c) => ({ key: c.Key, title: c.Title, url: c.SimpleUrl, audiences: c.AudienceIds })) : ms;
  }
  // 2. Home page: set to the test page, read back, restore.
  const rootUrl = `${web}/_api/web/rootfolder`;
  const before = await call('GET', `${rootUrl}?$select=WelcomePage`);
  out.welcomeBefore = before.WelcomePage;
  const set = await call('POST', rootUrl, { WelcomePage: PAGE }, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
  out.setResult = set;
  out.welcomeAfter = (await call('GET', `${rootUrl}?$select=WelcomePage`)).WelcomePage;
  if (before.WelcomePage !== undefined) {
    const restore = await call('POST', rootUrl, { WelcomePage: before.WelcomePage }, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
    out.restoreResult = restore;
    out.welcomeRestored = (await call('GET', `${rootUrl}?$select=WelcomePage`)).WelcomePage;
  }
  // 3. Cleanup of the spike's own nodes (deleting the heading deletes its children).
  if (CLEANUP) out.cleanup = await Promise.all(created.map((id) => call('POST', `${ql}/GetById(${id})`, null, { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' })));
  console.log(JSON.stringify(out, null, 2));
})();
```

**Futtatás:** a Cél site bármelyik lapján F12 → Console, másold be, Enter. A kimenő JSON-t kérem vissza, és azt is, hogy a kezdőlap a végén az eredeti maradt-e.

## Eredmények

### A – Forrás (2026-10-05, kommunikációs site, `SITEPAGEPUBLISHING#0`)

- **Hol van a menü:** a felső menü (`TopNavigationBar`, `MenuState` 1002) üres, minden a `QuickLaunch`-ban (1025) van.
  - Beállítások: `QuickLaunchEnabled` igaz, `HorizontalQuickLaunch` hamis, `MegaMenuEnabled` igaz, `UseShared` hamis, `NavAudienceTargetingEnabled` igaz.
- **Szintek:** a REST `$expand=Children,Children/Children` egy kérésben adja a fát. A `MenuState` `depth=5`-tel ugyanazt adja kulcsokkal (`Key` = csomópont-ID), `IsHidden` / `IsDeleted` / `NodeType` mezőkkel.
- **Célközönség:** csak a `MenuState` adja (`AudienceIds`: csoport GUID), a REST `AudienceIds` `null`. A GUID tenantfüggő.
- **Hivatkozás nélküli fejléc:** a felület `http://linkless.header/` URL-t és `IsExternal = true` értéket tárol. Változatlanul kell átvinni.
- **Beépített csomópontok** (ID < 2000):
  - 1031 „Kezdőlap”: URL a site relatív címe, záró `/` nélkül;
  - 1034 „Webhely tartalma”: `_layouts/15/viewlsts.aspx`, `IsExternal = true`;
  - 1033 „Legutóbbiak”: üres URL, a gyerekeit a SharePoint maga tölti.

  Ezek a célon is megvannak, ezért nem másolandók; a „Legutóbbiak” teljesen kimarad.
- **A listák hivatkozásai:**
  - kétféle URL fordul elő: a nézetre mutató (`…/Lists/Teszt lista/AllItems.aspx`) és a gyökérre mutató (`…/Lists/Teszt lista`; a `MenuState` záró `/`-rel adja);
  - mindkettő a lista gyökér-URL-jének előtagjával kezdődik → `{listurl:K}` + maradék;
  - a sablonon kívüli lista hivatkozása `{siterelative}`-vel tokenizálva megy át, és a célon törött lesz, ha ott nincs ilyen lista (figyelmeztetés).
- **Kezdőlap:** `rootfolder/WelcomePage` = `SitePages/TopicHome.aspx` (site-relatív).

**Döntés (a B eredményétől függően véglegesítendő):**
- **Kiolvasás** `MenuState`-tel (célközönség, rejtett csomópontok).
- **Írás** a REST csomópont-API-val, hozzáfűzéssel. A `SaveMenuState` az egész fát cseréli, és a meglévő csomópontokat törölné; ez ütközik a „soha nem törlünk” szabállyal.
- **Egyezés:** azonos szinten azonos címmel és feloldott URL-lel → „már megvan”.

### B – Cél (2026-10-05)

- **Belső hivatkozás** (`IsExternal = false`): a SharePoint megnyitja a célt, és ha nem létezik, HTTP 500-zal elutasítja: `-2130247147`, „nincs ilyen fájl vagy mappa”. A külső hivatkozást és a link nélküli fejlécet nem ellenőrzi.
  - **Döntés:** a C kísérlet szerint lásd lent.
- **Célközönség:** csomópontonként REST `MERGE {AudienceIds: [...]}` írja; a `MenuState` visszaadja. Más tenantba nem vihető át: kihagyás figyelmeztetéssel.
- **Sorrend:** a `QuickLaunch`-hoz fűzött új csomópont a „Webhely tartalma” elé kerül. Az alcsomópontok a felvétel sorrendjében állnak.
- **Kezdőlap:** a `rootfolder` `MERGE {WelcomePage}` hívása írja; a visszaállítás is működött. A SharePoint nem ellenőrzi, hogy a lap létezik-e, ezért a provider maga ellenőrzi.
- **Törlés:** a fejléc törlése a gyerekeit is törli (csak a spike takarításához kellett).
- **Mellékes hiba (spike 12):** a 3. fázis telepítése a lapot `CopyJet-teszt(1).aspx` néven hozta létre.
  - **Ok:** a még közzé nem tett lap első `savepageasdraft` mentése a címből újranevezi a fájlt, és ez ütközött a korábbi átnevezéssel.
  - **Javítás (1.6.1.0):** az átnevezés a vázlatmentés után, a közzététel előtt történik, utána a provider ellenőrzi a végleges nevet (`PAGE_NAME_DIFFERS`).

### C – Cél (2026-10-05)

- **URL-tárolás:** a relatív és az azonos site-ra mutató abszolút URL egyaránt szerver-relatívként tárolódik (`/sites/Cel/SitePages/TopicHome.aspx`). A lista gyökérre mutató URL-je változatlan marad.
- **Hiányzó cél:** a szerver-relatív URL-t a SharePoint `IsExternal = true` mellett is ellenőrzi (HTTP 500, `-2130247147`).
  - **Döntés:** a célon nem létező belső hivatkozás kimarad (`NAV_LINK_BROKEN`). Ha a kihagyott menüpontnak vannak gyerekei, link nélküli fejlécként (`http://linkless.header/`) kerül fel, hogy a gyerekek megmaradjanak.
- **A javított lapnév:** az 1.6.1.0 telepítése után a lap `CopyJet-teszt.aspx` néven van a Célon.
