# Spike 12 – Modern lapok

Állapot: **lezárva** · Érintett kód (3. fázis): `PageExtractor`, `PageProvider`, `deepTokenize` / `deepResolve`, webpart-ellenőrzés

## Kérdések

1. **Kiolvasás:** milyen alakban adja a `_api/sitepages/pages` a lap vásznát (`CanvasContent1`) és fejlécét (`LayoutWebpartsContent`)? Ugyanazt adja-e, mint a lap listaeleme a Webhelylapok tárban?
2. **Site-függő értékek a vásznon:**
   - hol vannak a site URL-ek (abszolút, relatív, kódolt), a web- és lista-azonosítók és a kép-URL-ek;
   - mit kell tokenizálni, és mi marad forrásra mutató hivatkozás?
3. **Webpartok:** mit ad a `GetClientSideWebParts`? Megtalálhatók-e benne a lapon használt webpartok, és megkülönböztethető-e a beépített és az egyedi (SPFx)?
4. **Írás** (B kísérlet, az A eredménye után): lap létrehozása a Célon, a vászon mentése, közzététel, kezdőlapnak beállítás.

## Előkészítés (Forrás)

Hozz létre a Forráson egy modern lapot **`CopyJet-teszt.aspx`** néven (a lap címe lehet „CopyJet teszt”), és tegyél rá:
- egy **Szöveg** webpartot, benne egy hivatkozással a „Teszt lista” listára;
- egy **Kép** webpartot egy feltöltött képpel;
- egy **Lista** webpartot, amely a „Teszt lista” listát mutatja;
- egy **Gyorshivatkozások** webpartot, benne egy belső (pl. a „Dokumentumok” tár) és egy külső hivatkozással;
- a lap fejlécébe egy képet, ha van kedved.

Utána tedd közzé a lapot.

## A. kísérlet – Forrás site (csak olvas)

```js
(async () => {
  const PAGE = 'CopyJet-teszt.aspx';
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (url) => { const r = await fetch(url, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status, error: (await r.text()).slice(0, 200) }; };
  const webInfo = await get(guess + '/_api/web?$select=Id,Url,ServerRelativeUrl,Title');
  const web = webInfo.Url, rel = webInfo.ServerRelativeUrl.replace(/\/$/, '');
  const out = {};
  // 1. Site pages as the sitepages API lists them.
  const pages = await get(`${web}/_api/sitepages/pages?$select=Id,Title,FileName,PageLayoutType,PromotedState,Url&$top=50`);
  out.pages = pages.httpStatus ? pages : (pages.value || []).map((p) => `${p.Id}: ${p.FileName} (${p.PageLayoutType}, promoted ${p.PromotedState})`);
  const page = (pages.value || []).filter((p) => p.FileName.toLowerCase() === PAGE.toLowerCase())[0];
  if (!page) { out.note = `Page ${PAGE} not found`; console.log(JSON.stringify(out, null, 2)); return; }
  // 2. The page's canvas through the sitepages API and through its SitePages list item.
  const api = await get(`${web}/_api/sitepages/pages(${page.Id})?$select=Title,PageLayoutType,CanvasContent1,LayoutWebpartsContent,BannerImageUrl,Description,TopicHeader,IsPageCheckedOutToCurrentUser`);
  const item = await get(`${web}/_api/web/lists/getbytitle('Site Pages')/items(${page.Id})?$select=Title,CanvasContent1,LayoutWebpartsContent,PageLayoutType,BannerImageUrl,PromotedState`);
  const itemOrUrl = item.httpStatus ? await get(`${web}/_api/web/getList('${encodeURIComponent(rel + '/SitePages')}')/items(${page.Id})?$select=Title,CanvasContent1,LayoutWebpartsContent,PageLayoutType,BannerImageUrl,PromotedState`) : item;
  out.api = api.httpStatus ? api : { layout: api.PageLayoutType, canvasType: typeof api.CanvasContent1, canvasLength: (api.CanvasContent1 || '').length, layoutWebparts: typeof api.LayoutWebpartsContent, banner: api.BannerImageUrl, description: api.Description, topicHeader: api.TopicHeader };
  out.listItem = itemOrUrl.httpStatus ? itemOrUrl : { canvasLength: (itemOrUrl.CanvasContent1 || '').length, sameCanvas: itemOrUrl.CanvasContent1 === api.CanvasContent1, promoted: itemOrUrl.PromotedState };
  // 3. The canvas: controls, web part IDs and the site-specific values inside their properties.
  let canvas = [];
  try { canvas = JSON.parse(api.CanvasContent1 || '[]'); } catch (e) { out.canvasParse = String(e); }
  const text = api.CanvasContent1 || '';
  const listIds = ((await get(`${web}/_api/web/lists?$select=Id,Title&$filter=Hidden eq false`)).value || []);
  const count = (needle) => (needle ? text.toLowerCase().split(needle.toLowerCase()).length - 1 : 0);
  out.controls = canvas.map((c) => ({
    controlType: c.controlType,
    position: c.position ? `${c.position.zoneIndex}/${c.position.sectionIndex}/${c.position.controlIndex} (${c.position.sectionFactor})` : null,
    webPartId: c.webPartId || (c.webPartData && c.webPartData.id) || null,
    title: c.webPartData ? c.webPartData.title : c.controlType === 4 ? 'Text' : null,
    propertyKeys: c.webPartData && c.webPartData.properties ? Object.keys(c.webPartData.properties).slice(0, 12) : [],
    hasServerProcessed: !!(c.webPartData && c.webPartData.serverProcessedContent),
    zoneEmphasis: c.zoneGroupMetadata || c.emphasis || null
  }));
  out.siteValues = {
    webId: count(webInfo.Id),
    siteUrlAbsolute: count(web),
    siteUrlRelative: count(rel + '/'),
    encodedRelative: count(encodeURIComponent(rel)),
    lists: listIds.map((l) => ({ title: l.Title, occurrences: count(l.Id) })).filter((l) => l.occurrences > 0)
  };
  // Image/document URLs in the canvas and the header.
  const urls = (text + ' ' + (typeof api.LayoutWebpartsContent === 'string' ? api.LayoutWebpartsContent : JSON.stringify(api.LayoutWebpartsContent || ''))).match(/(https?:\/\/[^"'\s\\]+|\/sites\/[^"'\s\\]+)/gi) || [];
  out.urls = urls.filter((u, i) => urls.indexOf(u) === i).slice(0, 20);
  // 4. Web parts available on this site: are the page's ones there, which are custom?
  const wps = await get(`${web}/_api/web/GetClientSideWebParts`);
  const list = wps.httpStatus ? [] : wps.value || wps || [];
  out.clientSideWebParts = wps.httpStatus ? wps : { count: list.length, sample: list.slice(0, 2).map((w) => Object.keys(w)) };
  out.pageWebParts = out.controls.filter((c) => c.webPartId).map((c) => {
    const found = list.filter((w) => (w.Id || '').toLowerCase() === String(c.webPartId).toLowerCase())[0];
    let manifest = null;
    try { manifest = found && found.Manifest ? JSON.parse(found.Manifest) : null; } catch { manifest = null; }
    return { id: c.webPartId, title: c.title, available: !!found, name: found ? found.Name : null, isInternal: manifest ? !!manifest.isInternal : null, componentType: found ? found.ComponentType : null };
  });
  const json = JSON.stringify(out, null, 2).split(location.host).join('contoso.sharepoint.com');
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:**
- a `controls` sorolja fel a lap szakaszait és webpartjait;
- a `siteValues` mutassa, hányszor fordul elő a vásznon a web azonosítója, a site URL-je (abszolút, relatív, kódolt) és melyik lista azonosítója;
- az `urls` a kép- és dokumentum-URL-eket;
- a `pageWebParts` a lapon használt webpartokat, hogy megvannak-e a site-on, és egyediek-e.

## Eredmény

### A – kiolvasás (2026-09-30, Forrás, `CopyJet-teszt.aspx`, Article)

- **Vászon:** a `_api/sitepages/pages(id)` `CanvasContent1` mezője **szöveg** (JSON-tömb, 10 389 karakter). A lap listaelemének (Webhelylapok) `CanvasContent1` mezője **más és nagyobb** (18 216). A sablonba a `sitepages` változat kerül.
- **Fejléc:** a `LayoutWebpartsContent` szöveg. A `BannerImageUrl` **abszolút** URL (`…/SiteAssets/SitePages/CopyJet-teszt/28452.jpeg`).
- **Vezérlők (`controlType`):**
  - 3 = webpart: Banner (`cbe7b0a9…`), Kép (`d1d91016…`), Lista (`f92bf067…`), Gyorshivatkozások (`c70391ea…`);
  - 4 = Szöveg;
  - 1 = üres oszlop;
  - 0 = lapbeállítás.

  A pozíció `zoneIndex/sectionIndex/controlIndex` alakú, `sectionFactor`-ral. A webpartoknak van `serverProcessedContent` részük is.
- **Site-függő értékek a vásznon:**
  - a web azonosítója 2×;
  - a site relatív URL-je 4× (pl. `/sites/Forras/Lists/Teszt%20lista/AllItems.aspx`, `/sites/Forras/Shared%20Documents`);
  - abszolút site URL 0×, kódolt relatív URL 0×;
  - lista-azonosítók: Dokumentumok 2×, Teszt lista 1×;
  - a Lista webpart a `selectedListId`, `selectedListUrl` és `selectedViewId` mezőben tárolja a listát és a nézetet.
- **Képek:** a feltöltött kép a `SiteAssets/SitePages/<lap>/` mappában van, ezt a csomagba kell tenni. A többi a Microsoft CDN-jéről jön (`cdn.hubblecontent.osi.office.net`), ezek változatlanul maradnak.
- **Webpartok:** a `GetClientSideWebParts` 75 elemet ad (`Id`, `Name`, `ComponentType`, `Manifest`). A lap mind a négy webpartja megvan, és `isInternal: true` (beépített).

→ **Terv:**
- tokenizálás: `{site}` / `{siterelative}` (URL), `{webid}` (új token), `{listkey:K}` (csak a sablonban szereplő listákra), `{viewid:K/Nézet}` (új token);
- a SiteAssets képek a csomagba kerülnek, és a célon ugyanarra az útvonalra kerülnek fel;
- `requiredWebParts`: a lap webpartjai, `isCustom` = nem `isInternal`.

## B. kísérlet – Cél site (**ír**): lap létrehozása a Forrás vásznából

Futtasd a **Cél** site-on. Ugyanabban a tenantban a Forrás oldalát is olvasni tudja. A kísérlet:
1. a Forrás `CopyJet-teszt.aspx` vásznában kicseréli a web azonosítóját, a site URL-jét és a listák azonosítóját a Cél értékeire (lista a cím alapján), ahogy a CopyJet tokenjei tennék;
2. **B1:** feltölti a fejlécképet a Cél `SiteAssets/SitePages/CopyJet-spike12-…/` mappájába;
3. **B2:** a `sitepages` API-val üres lapot hoz létre, kiveszi, elmenti vázlatként a vászonnal (`savepageasdraft`), majd közzéteszi (`publish`);
4. **B3:** a lap fájlját a kívánt névre nevezi (`FileLeafRef`);
5. visszaolvassa: név, elrendezés, vászonméret, maradt-e forrásérték, verzió, és hogy a webpartok megvannak-e a Célon.

A Cél „Teszt lista” listája legyen meg (a korábbi telepítésből), hogy a Lista webpart azonosítója leképezhető legyen. Új lapot hoz létre `CopyJet-spike12-….aspx` néven, amelyet utána törölhetsz.

```js
(async () => {
  // Run on the Cél site. Reads the Forrás page (same tenant) and recreates it here as a new page.
  const SOURCE = '/sites/Forras', SOURCE_PAGE = 'CopyJet-teszt.aspx';
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (url) => { const r = await fetch(url, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status, error: (await r.text()).slice(0, 200) }; };
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const target = await get(guess + '/_api/web?$select=Id,Url,ServerRelativeUrl');
  const web = target.Url, rel = target.ServerRelativeUrl.replace(/\/$/, '');
  const src = await get(location.origin + SOURCE + '/_api/web?$select=Id,Url,ServerRelativeUrl');
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const post = async (url, body) => { const r = await fetch(url, { method: 'POST', headers: W, body: body === undefined ? undefined : JSON.stringify(body) }); const t = await r.text(); let j = null; try { j = t ? JSON.parse(t) : null; } catch { j = null; } return { status: r.status, json: j, text: t.slice(0, 200) }; };
  const enc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const out = { steps: {} };

  // Source page: canvas, header, banner.
  const srcPages = (await get(`${src.Url}/_api/sitepages/pages?$select=Id,FileName`)).value || [];
  const srcPage = srcPages.filter((p) => p.FileName.toLowerCase() === SOURCE_PAGE.toLowerCase())[0];
  if (!srcPage) { console.log('Source page not found'); return; }
  const page = await get(`${src.Url}/_api/sitepages/pages(${srcPage.Id})?$select=Title,PageLayoutType,CanvasContent1,LayoutWebpartsContent,BannerImageUrl,Description`);

  // Source → target values, as CopyJet's tokens would resolve them: web ID, site URL, list IDs by list title.
  const lists = async (base) => ((await get(`${base}/_api/web/lists?$select=Id,Title&$filter=Hidden eq false`)).value || []);
  const srcLists = await lists(src.Url), tgtLists = await lists(web);
  const map = [[src.Id, target.Id], [src.ServerRelativeUrl, target.ServerRelativeUrl], [src.Url, web]];
  srcLists.forEach((l) => { const t = tgtLists.filter((x) => x.Title === l.Title)[0]; if (t) map.push([l.Id, t.Id]); });
  const swap = (s) => { let out2 = s || ''; map.forEach(([a, b]) => { out2 = out2.split(a).join(b).split(a.toUpperCase()).join(b); }); return out2; };
  out.replacements = map.length;

  // B1: banner image into the target SiteAssets (same relative path).
  const stamp = Date.now();
  const bannerSrc = page.BannerImageUrl ? page.BannerImageUrl.replace(/^https?:\/\/[^/]+/, '') : null;
  let banner = null;
  if (bannerSrc && bannerSrc.indexOf(src.ServerRelativeUrl + '/SiteAssets/') === 0) {
    const bytes = await (await fetch(`${src.Url}/_api/web/GetFileByServerRelativePath(decodedurl='${enc(bannerSrc)}')/$value`, { headers: { Accept: '*/*' } })).arrayBuffer();
    const folder = `${rel}/SiteAssets/SitePages/CopyJet-spike12-${stamp}`;
    await post(`${web}/_api/web/folders/addUsingPath(DecodedUrl='${enc(folder)}',overwrite=false)`);
    const name = bannerSrc.split('/').pop();
    const up = await fetch(`${web}/_api/web/GetFolderByServerRelativePath(decodedurl='${enc(folder)}')/Files/AddUsingPath(DecodedUrl='${enc(name)}',Overwrite=true)`, { method: 'POST', headers: { Accept: H.Accept, 'X-RequestDigest': digest }, body: bytes });
    banner = `${location.origin}${folder}/${name}`;
    out.steps.B1_bannerUpload = up.status;
  }

  // B2: create an empty page, then save the canvas as draft, then publish (the sitepages API flow).
  const created = await post(`${web}/_api/sitepages/pages`, { PageLayoutType: page.PageLayoutType, PromotedState: 0 });
  const id = created.json && created.json.Id;
  out.steps.B2_create = { status: created.status, id, fileName: created.json && created.json.FileName };
  if (!id) { console.log(JSON.stringify(out, null, 2)); return; }
  out.steps.B2_checkout = (await post(`${web}/_api/sitepages/pages(${id})/checkoutpage`)).status;
  const save = await post(`${web}/_api/sitepages/pages(${id})/savepageasdraft`, {
    Title: `${page.Title} (spike 12)`,
    CanvasContent1: swap(page.CanvasContent1),
    LayoutWebpartsContent: swap(page.LayoutWebpartsContent),
    BannerImageUrl: banner || swap(page.BannerImageUrl),
    Description: page.Description
  });
  out.steps.B2_save = { status: save.status, error: save.status >= 300 ? save.text : undefined };
  const pub = await post(`${web}/_api/sitepages/pages(${id})/publish`);
  out.steps.B2_publish = { status: pub.status, error: pub.status >= 300 ? pub.text : undefined };

  // B3: rename the page file to the wanted name (FileLeafRef of its SitePages item).
  const wanted = `CopyJet-spike12-${stamp}.aspx`;
  const pagesList = `${web}/_api/web/getList('${enc(rel + '/SitePages')}')`;
  const ren = await post(`${pagesList}/items(${id})/ValidateUpdateListItem`, { formValues: [{ FieldName: 'FileLeafRef', FieldValue: wanted.replace(/\.aspx$/, '') }], bNewDocumentUpdate: false });
  out.steps.B3_rename = { status: ren.status, errors: ren.json && (ren.json.value || []).filter((v) => v.HasException).map((v) => `${v.FieldName}: ${v.ErrorMessage}`) };

  // Read back: name, layout, canvas size, whether source values are gone, published state.
  const back = await get(`${web}/_api/sitepages/pages(${id})?$select=FileName,Title,PageLayoutType,CanvasContent1,BannerImageUrl,Url`);
  const text = back.CanvasContent1 || '';
  out.back = {
    fileName: back.FileName, title: back.Title, layout: back.PageLayoutType,
    canvasLength: text.length, sourceCanvasLength: (page.CanvasContent1 || '').length,
    sourceValuesLeft: { webId: text.split(src.Id).length - 1, siteUrl: text.split(src.ServerRelativeUrl + '/').length - 1, lists: srcLists.filter((l) => text.indexOf(l.Id) >= 0).map((l) => l.Title) },
    banner: (back.BannerImageUrl || '').replace(location.origin, ''),
    url: back.Url
  };
  const item = await get(`${pagesList}/items(${id})?$select=OData__UIVersionString,FileLeafRef,PromotedState`);
  out.back.version = item.OData__UIVersionString;
  out.back.fileLeafRef = item.FileLeafRef;
  // Web parts of the page available on the target.
  const wps = (await get(`${web}/_api/web/GetClientSideWebParts`)).value || [];
  let canvas = [];
  try { canvas = JSON.parse(page.CanvasContent1 || '[]'); } catch { canvas = []; }
  out.targetWebParts = canvas.filter((c) => c.webPartId).map((c) => ({ id: c.webPartId, available: wps.some((w) => (w.Id || '').toLowerCase() === c.webPartId.toLowerCase()) }));
  const json = JSON.stringify(out, null, 2).split(location.host).join('contoso.sharepoint.com');
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

### B – eredmény (2026-09-30, Cél, azonos tenant)

- ✅ **A folyamat működik:**
  1. `POST sitepages/pages {PageLayoutType, PromotedState}` → 201, `Id`, a fájlnév `Page.aspx`;
  2. `pages(id)/checkoutpage`;
  3. `pages(id)/savepageasdraft {Title, CanvasContent1, LayoutWebpartsContent, BannerImageUrl, Description}`;
  4. `pages(id)/publish`.
- ✅ **Átnevezés:** a Webhelylapok elemének `FileLeafRef` mezője `ValidateUpdateListItem`-mel (a `.aspx` nélkül). A lap `CopyJet-spike12-….aspx` lett. Mivel ez a közzététel **után** történt, a verzió **1.1** lett, egy közzé nem tett változással.
- ✅ **Fejléckép:** a Cél `SiteAssets/SitePages/<lap>/` mappájába feltöltve, és a `BannerImageUrl`-be téve jól jelenik meg.
- ✅ **Megjelenés (Zoli ellenőrizte):** a fejléckép, a Kép, a Lista (a Cél „Teszt listáját” mutatja) és a Gyorshivatkozások is helyes.
- A spike egyszerű szövegcseréje után a vásznon **megmaradt a web azonosítója (2×) és a „Dokumentumok” tár azonosítója.** A lap ennek ellenére jól működött: a gyorshivatkozás URL-je a Célra mutatott.
- ✅ A lap mind a négy webpartja megvan a Célon (`GetClientSideWebParts`).

→ **Döntések a `PageProvider`-hez:**
- sorrend: létrehozás → **átnevezés** → kivétel → mentés vázlatként → közzététel. Így 1.0 lesz;
- a vászon a CopyJet tokenizálójával készül: a GUID bármilyen alakban (kis- és nagybetű, kapcsos zárójel, `%7B`) felismerhető. Az alapértelmezett tár a `{listkey}`-en keresztül a cél alapértelmezett tárára képeződik;
- a feloldás engedékeny (`resolveKnown`): a sablonban nem szereplő listák azonosítója a forrásra mutat marad, erről a Setup figyelmeztet;
- a kezdőlap beállítása a navigációs lépésbe kerül, mert a cél kezdőlapjának cseréje külön döntés.

Állapot: **lezárva**.


## Utólagos megjegyzés (2026-10-05)

A provider eredetileg a vázlatmentés előtt nevezte át a lapot, így a telepítés `CopyJet-teszt(1).aspx`-et hozott létre. A még közzé nem tett lap első `savepageasdraft` mentése ugyanis a címből újranevezi a fájlt. Az 1.6.1.0 óta a sorrend: létrehozás → kivétel → vázlatmentés → átnevezés → közzététel (részletek: `13-navigation.md`, B).
