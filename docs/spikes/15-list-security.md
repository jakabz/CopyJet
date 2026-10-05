# Spike 15 – Listajogosultságok

Állapot: **lezárva** · Érintett kód (4. fázis): `ListSecurityExtractor`, `ListSecurityProvider`, séma `list.security`

## Kérdések

1. **Kiolvasás (A, Forrás):**
   - Mit ad a REST az egyedi jogosultságú listák szerepkör-hozzárendeléseiről?
     - tag: SharePoint-csoport (8), felhasználó (1), biztonsági csoport (4);
     - szintek: név, `RoleTypeKind`, rejtett-e.
   - Megjelenik-e a „Korlátozott hozzáférés” (Limited Access, `RoleTypeKind` 1, rejtett) és a megosztási linkek csoportja (`SharingLinks.…`)? Ezeket ki kell hagyni, mert a SharePoint maga kezeli őket.
   - Milyen login-alakúak a tagok? Ez kell a `{principal:key}` tokenhez.
     - felhasználó: `i:0#.f|membership|…`;
     - biztonsági csoport: `c:0t.c|tenant|…`;
     - M365-csoport: `c:0o.c|federateddirectoryclaimprovider|…`;
     - „Mindenki”: `c:0(.s|true`, illetve `c:0-.f|rolemanager|spo-grid-all-users/…`.
2. **Írás (B, Cél):**
   - `breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)` után a futtató marad egyedül (spike 14);
   - `addroleassignment` csoportnak és `ensureuser`-rel felvett felhasználónak;
   - **ugyanaz a hozzárendelés kétszer:** hibát ad, vagy csendben nem csinál semmit? Ebből derül ki, hogy az újrafuttatás idempotens-e;
   - **a futtató saját Teljes hozzáférésének levétele:** utána is eléri-e a listát? Site-gyűjtemény-adminisztrátorként igen, de egy egyszerű tulajdonosnál ez már a csoporttagságon múlik.

A második pont azért fontos, mert a célon a futtató mindig a lista jogosultságai közé kerül akkor is, ha a forrásban nem volt ott. Döntés kell róla: bent hagyjuk (a napló jelzi), vagy levesszük, ha a forrás hozzárendelései között nem szerepelt.

## Előkészítés (Forrás)

Az egyik tesztlistán, például a **„Teszt lista”**-n (Lista beállításai → *Engedélyek ehhez a listához*):
1. **Öröklés megszüntetése** (*Engedélyöröklés leállítása*).
2. A **Látogatók** csoportot vedd le.
3. Adj **Szerkesztés** jogot a **„Forrás Projektmenedzserek”** csoportnak (spike 07-ből).
4. Adj **Olvasás** jogot egy felhasználónak, például AdeleV-nek.
5. Ha van kéznél biztonsági vagy M365-csoport, annak is adj valamilyen jogot.
6. Ha szeretnéd: egy elemét oszd meg linkkel valakivel. Ettől jelenik meg a listán a „Korlátozott hozzáférés”.

## A. kísérlet – Forrás site (csak olvas)

A kód e-mail-címet nem ír ki, a login nevekből csak az előtagot (`i:0#.f`, `c:0t.c` …) és a felhasználónév első betűjét.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const get = async (p) => { const r = await fetch(guess + '/_api/' + p, { headers: H }); return r.ok ? r.json() : { httpStatus: r.status, error: (await r.text()).slice(0, 200) }; };
  const mask = (login) => { const m = /^(.*\|)([^|]*)$/.exec(login || ''); return m ? `${m[1].replace(/\|$/, '')} | ${m[2] ? m[2][0] + '…' : ''}` : login; };
  const owner = await get('web/AssociatedOwnerGroup?$select=Id,Title');
  const member = await get('web/AssociatedMemberGroup?$select=Id,Title');
  const visitor = await get('web/AssociatedVisitorGroup?$select=Id,Title');
  const assoc = { [owner.Id]: 'owners', [member.Id]: 'members', [visitor.Id]: 'visitors' };
  const lists = (await get(`web/lists?$select=Id,Title,BaseTemplate,HasUniqueRoleAssignments&$filter=Hidden eq false`)).value || [];
  const out = { uniqueLists: [], inheriting: lists.filter((l) => !l.HasUniqueRoleAssignments).map((l) => l.Title) };
  for (const l of lists.filter((x) => x.HasUniqueRoleAssignments)) {
    const ra = await get(`web/lists(guid'${l.Id}')/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Id,Member/Title,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind,RoleDefinitionBindings/Hidden`);
    out.uniqueLists.push({
      list: l.Title,
      template: l.BaseTemplate,
      assignments: ra.httpStatus ? ra : (ra.value || []).map((a) => ({
        member: `${a.Member.Title} (id ${a.Member.Id}, type ${a.Member.PrincipalType}${assoc[a.Member.Id] ? ', ' + assoc[a.Member.Id] : ''})`,
        login: a.Member.PrincipalType === 8 ? '(SharePoint group)' : mask(a.Member.LoginName),
        roles: a.RoleDefinitionBindings.map((b) => `${b.Name} (kind ${b.RoleTypeKind}${b.Hidden ? ', rejtett' : ''})`)
      }))
    });
  }
  // Is the web's member list the same principals? (A list can name principals the web does not.)
  const webRa = await get('web/roleassignments?$expand=Member&$select=Member/Id,Member/Title,Member/PrincipalType');
  out.webPrincipals = (webRa.value || []).map((a) => `${a.Member.Title} (id ${a.Member.Id}, type ${a.Member.PrincipalType})`);
  console.log(JSON.stringify(out, null, 2));
})();
```

**Futtatás:** a Forrás site egy lapján F12 → Console, másold be, Enter. A kimenő JSON-t kérem vissza.

## B. kísérlet – Cél site (ír, majd töröl)

A kísérlet létrehoz egy **„CopyJetSpike15”** listát, és megszünteti az öröklést. Ezután:
- a Látogatók csoportnak Olvasás jogot ad, kétszer egymás után;
- ha a `LOGIN` meg van adva, a felhasználót `ensureuser`-rel felveszi, és Olvasás jogot ad neki;
- leveszi a saját Teljes hozzáférésedet, és megnézi, olvasod-e még a listát.

A végén a listát a lomtárba teszi. A `LOGIN` értékét írd át egy Cél tenantbeli felhasználóra, például `i:0#.f|membership|adelev@…`, vagy hagyd üresen.

```js
(async () => {
  const LOGIN = '';
  const CLEANUP = true;
  const site = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const J = 'application/json;odata=nometadata';
  const digest = (await (await fetch(site + '/_api/contextinfo', { method: 'POST', headers: { Accept: J } })).json()).FormDigestValue;
  const call = async (url, method = 'GET', body) => {
    const headers = { Accept: J, 'X-RequestDigest': digest };
    if (body !== undefined) { headers['Content-Type'] = J; body = JSON.stringify(body); }
    const r = await fetch(site + '/_api/' + url, { method, headers, body });
    const text = await r.text();
    if (!r.ok) return { httpStatus: r.status, error: text.slice(0, 300) };
    try { return text ? JSON.parse(text) : 'ok'; } catch { return text || 'ok'; }
  };
  const out = {};
  const list = await call('web/lists', 'POST', { Title: 'CopyJetSpike15', BaseTemplate: 100 });
  if (list.httpStatus) { console.log(JSON.stringify({ createList: list }, null, 2)); return; }
  const L = `web/lists(guid'${list.Id}')`;
  const ras = async () => ((await call(`${L}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=Member/Id,Member/Title,RoleDefinitionBindings/Name`)).value || []).map((a) => `${a.Member.Title} (${a.Member.Id}): ${a.RoleDefinitionBindings.map((b) => b.Name).join(', ')}`);
  const me = await call('web/currentuser?$select=Id,IsSiteAdmin');
  out.me = { id: me.Id, siteAdmin: me.IsSiteAdmin };
  out.break = await call(`${L}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`, 'POST');
  out.afterBreak = await ras();
  const read = await call('web/roledefinitions/getbytype(2)?$select=Id,Name');
  const full = await call('web/roledefinitions/getbytype(5)?$select=Id,Name');
  const visitors = await call('web/AssociatedVisitorGroup?$select=Id,Title');
  out.addVisitorsRead = await call(`${L}/roleassignments/addroleassignment(principalid=${visitors.Id},roledefid=${read.Id})`, 'POST');
  out.addVisitorsReadAgain = await call(`${L}/roleassignments/addroleassignment(principalid=${visitors.Id},roledefid=${read.Id})`, 'POST');
  if (LOGIN) {
    const u = await call('web/ensureuser', 'POST', { logonName: LOGIN });
    out.ensureUser = u.httpStatus ? u : `ok (id ${u.Id})`;
    if (!u.httpStatus) out.addUserRead = await call(`${L}/roleassignments/addroleassignment(principalid=${u.Id},roledefid=${read.Id})`, 'POST');
  }
  out.afterAdds = await ras();
  out.removeMyFullControl = await call(`${L}/roleassignments/removeroleassignment(principalid=${me.Id},roledefid=${full.Id})`, 'POST');
  out.afterRemove = await ras();
  const still = await call(`${L}?$select=Title`);
  out.canStillReadList = !still.httpStatus;
  const perms = await call(`${L}/EffectiveBasePermissions`);
  out.myEffectivePermsOnList = perms.httpStatus ? perms : perms;
  if (CLEANUP) { const r = await call(`${L}/recycle`, 'POST'); out.cleanup = r.httpStatus ? r : 'recycled'; }
  console.log(JSON.stringify(out, null, 2));
})();
```

**Futtatás:** a Cél site egy lapján F12 → Console, másold be, Enter. A kimenő JSON-t kérem vissza.

## Eredmények

### A. kísérlet (Forrás, 2026-10-05)

- Egyedi jogosultsága csak a „Teszt lista” listának van. A többi nem rejtett lista örököl, köztük a Dokumentumok, a Teszt lookup forrás és a TesztDoktár.
- A „Teszt lista” hozzárendelései:

  | Tag | Típus | Szintek |
  | --- | --- | --- |
  | Forrás Owners (alapcsoport) | 8 | Teljes hozzáférés (5) |
  | Forrás Members (alapcsoport) | 8 | Szerkesztés (6) |
  | Forrás Projektmenedzserek | 8 | Szerkesztés (6), Munkatárs (3), CopyJet olvasó (0, egyedi) |
  | Adele Vance | 1, `i:0#.f\|membership` | Olvasó (2) |
  | KerVezAdmins (biztonsági csoport) | 4, `c:0t.c\|tenant` | Korlátozott nézet (8, rejtett) |
  | KerVezPortal Members (M365-csoport) | 4, `c:0o.c\|federateddirectoryclaimprovider` | Korlátozott hozzáférés (1, rejtett) |

- **Egy tagnak több szintje is lehet** (Projektmenedzserek), köztük egyedi szint (`RoleTypeKind` 0).
- **A rejtett szintek kétfélék:**
  - a „Korlátozott hozzáférés” (1) a megosztásból jön, ezt a SharePoint kezeli;
  - a „Korlátozott nézet” (8) ellenben kiosztható szint.
- **Biztonsági és M365-csoport** egyaránt 4-es típusú tag, a login előtagjuk különbözik.
- **A lista tagjai mind szerepelnek a web hozzárendelései között is.** Ez itt véletlen, egy lista nevesíthet olyan tagot is, aki a weben nincs.

### B. kísérlet (Cél, site-gyűjtemény-adminisztrátorként)

- `breakroleinheritance(false, true)` után csak a futtató marad a listán, Teljes hozzáféréssel (mint a spike 14-ben).
- **Ugyanaz a hozzárendelés kétszer** (`addroleassignment`, Látogatók → Olvasás): mindkétszer sikeres, és nem lesz belőle két bejegyzés. Az újrafuttatás tehát nem okoz hibát.
- **A futtató saját Teljes hozzáférése** (`removeroleassignment`) levehető. Utána a listán csak a Látogatók maradnak, a futtató site-adminként továbbra is olvassa a listát, és teljes jogai vannak rajta.

## Döntések

- **Mi kerül át:**
  - csak az egyedi jogosultságú listák (`HasUniqueRoleAssignments`);
  - a Setupban a lista alatt „Jogosultságok → Egyedi jogosultságok (N hozzárendelés)” sorként jelennek meg, a listával együtt kijelölődnek;
  - a sablonban: `list.security = { breakInheritance: true, copyRoleAssignments: false, roleAssignments }`. A séma ezt már tartalmazta, nem változott.
- **Tagok:**
  - az alapcsoportok `{associated…group}` tokenként;
  - a site saját csoportjai `{groupkey:K}` tokenként, a csoport-extractor kulcsával. A Setup a hiányzó csoportot függőségként ajánlja fel;
  - a felhasználók, a biztonsági és az M365-csoportok `{principal:key}` tokenként, a felhasználó-leképezésen át;
  - kimarad az a csoport, amelynek a weben nincs szerepe (a CopyJet nem másolja): `LIST_SECURITY_GROUP_OUTSIDE_TEMPLATE`.
- **Szintek:**
  - a beépítettek angol néven, a célon `RoleTypeKind` szerint;
  - új beépített szint: Review (7) és Restricted View (8);
  - az egyedi szint név szerint, és ha a célon nincs ilyen: `ROLE_NOT_FOUND`;
  - a „Korlátozott hozzáférés” (1) soha nem kerül át, és kimarad az a tag is, akinek csak ilyen szintje van (megosztás).
- **Telepítés** (`ListSecurityProvider`, a nézetek után, a lista zárolásával):
  - **öröklő céllista:**
    - `breakRoleInheritance(false, false)`: a lista elemeinek saját jogosultsága megmarad;
    - utána a sablon hozzárendelései kerülnek fel;
  - **már egyedi céllista:** csak a hiányzó hozzárendelések kerülnek fel, a meglévőket nem veszi el;
  - **„Kihagyás” mód:** egy előre meglévő lista jogosultságához nem nyúl (`LIST_SECURITY_SKIPPED`). Az ebben a futásban létrehozott lista „frissítés” módot kap, mint a nézetei.
- **A futtató saját Teljes hozzáférése:**
  - az öröklés megszüntetése után akkor veszi le, ha a sablon nem nevesíti a futtatót, és a futtató site-gyűjtemény-adminisztrátor (`LIST_SECURITY_INSTALLER_REMOVED`);
  - különben bent hagyja, hogy ne zárja ki magát (`LIST_SECURITY_INSTALLER_KEPT`);
  - más hozzárendelést a program soha nem vesz el.

## Utólagos megjegyzés (1.9.1.0)

Az első telepítésen a „Forrás Projektmenedzserek” csoport lépése elbukott, mert a Célon nincs „CopyJet olvasó” szint. A `roledefinitions/getbyname('…')` hiányzó szintre HTTP 500-at ad (`-2146232832`, „Nem található a jogosultsági szint”), nem 404-et. Emiatt a listajogosultság lépése is letiltott lett. Javítás: az egyedi szintet a web szintjeinek listájából keresem ki név szerint, így a hiány `ROLE_NOT_FOUND` figyelmeztetés lesz, a csoport és a lista jogosultsága pedig a többi szinttel települ.

## Utólagos megjegyzés (1.9.2.0)

A második telepítésen Adele és a KerVezAdmins nem került a listára (`LIST_SECURITY_PRINCIPAL_MISSING`). A `{principal:…}` tokeneket a felhasználó-leképezés lustán, a lépés kérésére tölti ki. Eddig ezt csak az elem- és fájllépések kérték, a listajogosultság nem. Javítás: a `ListSecurityProvider` maga kéri a leképezést. Akit csak a helyettesítő felhasználóra lehetett leképezni, az kimarad, mert a helyettesítő nem kaphatja meg a jogait.


Ellenőrizve (1.9.2.0, Cél): a lista öt tagja a forrás szerint megkapta a jogát (Owners: Teljes hozzáférés; Members: Szerkesztés; Projektmenedzserek: Szerkesztés, Munkatárs, CopyJet olvasó; Adele: Olvasó; KerVezAdmins: Korlátozott nézet). A telepítő saját Teljes hozzáférését a program levette. A „CopyJet olvasó” szint a Célon kézzel létrehozva, frissítés módban a csoport webszintű szerepeit is pótolta.
