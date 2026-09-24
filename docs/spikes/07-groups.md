# Spike 07 – SharePoint-csoportok: kiválasztás, alapcsoportok, jogosultsági szintek, tulajdonos

Állapot: **lezárva** (2026-09-24) · Érintett kód: `GroupExtractor`, `GroupProvider`

## Kérdések

1. **Kiválasztás:** a `web.siteGroups` a site collection összes csoportját adja, a rendszercsoportokat is (pl. megosztási linkek „Limited Access” csoportjai). Mely csoportok tartoznak a site-hoz?
   - **Feltevés:** azok, amelyeknek a weben szerepkör-hozzárendelésük van.
2. **Alapcsoportok:** a web `AssociatedOwnerGroup`, `AssociatedMemberGroup` és `AssociatedVisitorGroup` csoportja a célon is megvan. Ezeket nem hozzuk létre, hanem a `{associatedownergroup}` stb. tokenekkel párosítjuk.
3. **Jogosultsági szintek:** a nevük nyelvfüggő („Teljes hozzáférés” / „Full Control”). A beépítettek `RoleTypeKind` értéke nyelvfüggetlen (2 = Read, 3 = Contribute, 4 = Design, 5 = Full Control, 6 = Edit). Van-e egyedi szint (`RoleTypeKind` = 0)?
4. **Létrehozás REST-ből:**
   - **B1:** `POST /_api/web/sitegroups` a beállításokkal;
   - **B2:** tulajdonos beállítása: `sitegroups(<id>)/SetUserAsOwner(<principal ID>)`, ahol a tulajdonos egy csoport, a web Tulajdonosok csoportja;
   - **B3:** jogosultság hozzárendelése: `roleassignments/addroleassignment(principalid=…, roledefid=…)`, ahol a szint azonosítója a `roledefinitions/getByType(3)` hívásból jön.

## A. kísérlet – Forrás site (csak olvas)

**Előkészítés a Forrás site-on:**
1. *Webhelyengedélyek* → *Speciális engedélybeállítások* → *Csoport létrehozása*: név **„Forrás Projektmenedzserek”**, jogosultsági szint **Közreműködés** (Contribute), tulajdonos: te magad.
2. Ha szeretnéd: *Engedélyszintek* → hozz létre egy egyedi szintet (pl. „CopyJet olvasó+”), és add hozzá a fenti csoporthoz is.

Utána: F12 → Console → beillesztés → Enter. A kiírt JSON-t küldd vissza. A kód felhasználói e-mail-címet nem ír ki, csak a tulajdonos típusát.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const webInfo = await (await fetch(guess + '/_api/web?$select=Url,Title', { headers: H })).json();
  const web = webInfo.Url;
  const get = async (p) => (await fetch(web + p, { headers: H })).json();
  const assoc = {};
  for (const k of ['AssociatedOwnerGroup', 'AssociatedMemberGroup', 'AssociatedVisitorGroup']) {
    const g = await get(`/_api/web/${k}?$select=Id,Title`);
    assoc[k] = `${g.Id}: ${g.Title}`;
  }
  const groups = (await get('/_api/web/sitegroups?$select=Id,Title,Description,IsHiddenInUI,OnlyAllowMembersViewMembership,AllowMembersEditMembership,AllowRequestToJoinLeave,AutoAcceptRequestToJoinLeave,Owner/PrincipalType,Owner/Title&$expand=Owner')).value || [];
  const ras = (await get('/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind,RoleDefinitionBindings/Hidden')).value || [];
  const defs = (await get('/_api/web/roledefinitions?$select=Id,Name,RoleTypeKind,Hidden,BasePermissions')).value || [];
  const roleOf = {};
  ras.forEach((ra) => (roleOf[ra.PrincipalId] = ra.RoleDefinitionBindings.map((b) => `${b.Name} (kind ${b.RoleTypeKind}${b.Hidden ? ', rejtett' : ''})`)));
  const out = {
    web, siteTitle: webInfo.Title, associated: assoc,
    groups: groups.map((g) => `${g.Id} | ${g.Title} | hiddenUI=${g.IsHiddenInUI} | owner=${g.Owner ? (g.Owner.PrincipalType === 8 ? 'csoport: ' + g.Owner.Title : 'felhasználó (típus ' + g.Owner.PrincipalType + ')') : '-'} | onlyMembersView=${g.OnlyAllowMembersViewMembership} membersEdit=${g.AllowMembersEditMembership} requestJoin=${g.AllowRequestToJoinLeave} | webRoles=${(roleOf[g.Id] || []).join('; ') || '-'}`),
    otherWebRoleAssignments: ras.filter((ra) => ra.Member.PrincipalType !== 8).map((ra) => `típus ${ra.Member.PrincipalType} | ${ra.RoleDefinitionBindings.map((b) => b.Name).join(', ')}`),
    roleDefinitions: defs.map((d) => `${d.Id} | ${d.Name} | kind ${d.RoleTypeKind}${d.Hidden ? ' | rejtett' : ''}`)
  };
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## B. kísérlet – Cél site (**ír**: 1 tesztcsoport)

Csak a Cél **teszt** site-on futtasd. Semmit nem töröl. Létrehozza a **„CopyJet Spike csoport”** csoportot, tulajdonosnak a web Tulajdonosok csoportját állítja be, és a weben **Közreműködés** (Contribute) jogot ad neki. A csoport a teszt után a *Webhelyengedélyek* oldalon kézzel törölhető.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const W = { ...H, 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest };
  const short = async (r) => { try { return (await r.text()).slice(0, 300); } catch { return ''; } };
  const TITLE = 'CopyJet Spike csoport';
  const out = { web };

  const owners = await (await fetch(`${web}/_api/web/AssociatedOwnerGroup?$select=Id,Title`, { headers: H })).json();
  out.ownersGroup = `${owners.Id}: ${owners.Title}`;

  // B1: create the group
  let g = ((await (await fetch(`${web}/_api/web/sitegroups?$select=Id,Title&$filter=Title eq '${TITLE}'`, { headers: H })).json()).value || [])[0];
  if (!g) {
    const r = await fetch(`${web}/_api/web/sitegroups`, { method: 'POST', headers: W,
      body: JSON.stringify({ Title: TITLE, Description: 'Törölhető teszt', OnlyAllowMembersViewMembership: false, AllowMembersEditMembership: true, AllowRequestToJoinLeave: false }) });
    out.b1_status = r.status; out.b1_body = r.ok ? '' : await short(r);
    g = r.ok ? await r.json() : null;
  }
  out.group = g && `${g.Id}: ${g.Title}`;

  if (g) {
    // B2: owner = the web's Owners group
    let r = await fetch(`${web}/_api/web/sitegroups(${g.Id})/SetUserAsOwner(${owners.Id})`, { method: 'POST', headers: W });
    out.b2_status = r.status; out.b2_body = r.ok ? '' : await short(r);
    const after = await (await fetch(`${web}/_api/web/sitegroups(${g.Id})?$select=AllowMembersEditMembership,Owner/Id,Owner/Title,Owner/PrincipalType&$expand=Owner`, { headers: H })).json();
    out.b2_owner = after.Owner ? `${after.Owner.Id}: ${after.Owner.Title} (típus ${after.Owner.PrincipalType})` : null;
    out.b1_membersEdit = after.AllowMembersEditMembership;

    // B3: Contribute on the web, found by RoleTypeKind 3
    const def = await (await fetch(`${web}/_api/web/roledefinitions/getByType(3)?$select=Id,Name`, { headers: H })).json();
    out.b3_roleDef = `${def.Id}: ${def.Name}`;
    r = await fetch(`${web}/_api/web/roleassignments/addroleassignment(principalid=${g.Id},roledefid=${def.Id})`, { method: 'POST', headers: W });
    out.b3_status = r.status; out.b3_body = r.ok ? '' : await short(r);
    const ra = await (await fetch(`${web}/_api/web/roleassignments/getByPrincipalId(${g.Id})?$expand=RoleDefinitionBindings&$select=RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind`, { headers: H })).json();
    out.b3_roles = (ra.RoleDefinitionBindings || []).map((b) => `${b.Name} (kind ${b.RoleTypeKind})`);
  }

  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

## Mit várunk?

- **A:**
  - a „Forrás Projektmenedzserek” csoportnak van webes szerepköre (Közreműködés, kind 3);
  - a rendszercsoportoknak nincs, vagy `hiddenUI` = `true`;
  - az alapcsoportok azonosítói megjelennek;
  - egy esetleges egyedi szint kind 0.
- **B:**
  - `b1_status` = 201;
  - `b2_owner` a Tulajdonosok csoport (típus 8);
  - `b3_roles` tartalmazza a Közreműködést (kind 3).

## Eredmény

### A – kiolvasás (2026-09-24, Forrás site, magyar nyelvű)

| ID | Csoport | Tulajdonos | Webes jogosultság |
| --- | --- | --- | --- |
| 3 | Forrás Owners (alapcsoport) | Forrás Owners | Teljes hozzáférés (kind 5) |
| 5 | Forrás Members (alapcsoport) | Forrás Owners | Szerkesztés (kind 6) |
| 4 | Forrás Visitors (alapcsoport) | Forrás Owners | Olvasó (kind 2) |
| 10 | **Forrás Projektmenedzserek** | felhasználó (típus 1) | Munkatárs (kind 3), **CopyJet olvasó (kind 0, egyedi)** |

- **Kiválasztás:** a site collection mind a 4 csoportjának van webes jogosultsága, felhasználóhoz közvetlenül nincs webes hozzárendelés. A feltevés (webes jogosultsággal rendelkező csoport) itt nem szűrt ki semmit, de rendszercsoportot sem vett fel.
- **Alapcsoportok:** a nevük a site címéből és angol utótagból áll (a magyar site-on is „Forrás Owners”). Ezeket nem visszük át, a célon a `{associated…group}` tokenek oldják fel őket.
- **Szintek:** a beépített szintek neve nyelvfüggő (Munkatárs, Szerkesztés, Olvasó, Teljes hozzáférés, Tervezés), a `RoleTypeKind` nyelvfüggetlen. A „Korlátozott hozzáférés” (kind 1) rejtett. Az egyedi szint kind 0.

→ **Döntés (`GroupExtractor`):**
- Az alapcsoportokon kívüli, webes jogosultsággal rendelkező csoportokat viszi át.
- **Cím:** ha a site címével kezdődik, `{sitename} …` lesz („Forrás Projektmenedzserek” → `{sitename} Projektmenedzserek`, ahogy a sablonmintában).
- **Szintek:** a beépítettek egységes angol névvel kerülnek a sablonba (2 = Read, 3 = Contribute, 4 = Design, 5 = Full Control, 6 = Edit); a célon a `RoleTypeKind` szerint keressük őket. Az egyedi szint a saját nevével kerül a sablonba; a célon név szerint keressük, ha nincs, figyelmeztetéssel kimarad.
  - A séma nem írja le a szintek jogosultság-készletét (`BasePermissions`), ezért egyedi szintet a célon nem hozunk létre. Ez a séma 1.1-es bővítésével lehetne pótolni.
- **Tulajdonos:** ha a tulajdonos csoport, `{associatedownergroup}` vagy `{groupkey:X}`. Ha felhasználó, akkor a 2. fázisig (felhasználó-leképezés) `{associatedownergroup}` kerül a helyére, figyelmeztetéssel (`GROUP_OWNER_USER`).
- **Tagok:** az alapértelmezés szerint nem kerülnek át (`includeMembers: false`), ez 2. fázisos funkció.

### B – létrehozás (2026-09-24, Cél site)

- **B1 – csoport létrehozása: ✅.** A `POST /_api/web/sitegroups` 201-et adott, a beállítások (pl. `AllowMembersEditMembership: true`) megmaradtak.
- **B2 – tulajdonos REST-ből: ❌ csendes hiba.** A `sitegroups(11)/SetUserAsOwner(3)` HTTP 200-at adott, de a tulajdonos **a létrehozó felhasználó maradt**, nem a „Cél Owners” csoport (3). A hívás csak felhasználót kezel, csoport ID-ra nem jelez hibát.
  → A tulajdonost **CSOM**-mal állítjuk be (`Group.Owner = SiteGroups.GetById(<tulajdonos>)`, `Update()`, `csomXml.setGroupOwnerBody`), és beállítás után visszaolvasva ellenőrizzük. Eltérésnél figyelmeztetés: `GROUP_OWNER_NOT_SET`.
- **B3 – jogosultság a `RoleTypeKind` alapján: ✅.** A `roledefinitions/getByType(3)` a „Munkatárs” szintet adta, a `roleassignments/addroleassignment(principalid=11, roledefid=…)` 200-at adott, és a csoport megkapta a szintet.

## C. kísérlet – Cél site (**ír**): tulajdonos CSOM-mal

Csak a Cél **teszt** site-on futtasd, a B-ben létrehozott „CopyJet Spike csoport” csoporton. **Pontosan azt a kérést** küldi el, amelyet a provider küldene: a csoport tulajdonosa a web Tulajdonosok csoportja lesz. Semmit nem töröl.

```js
(async () => {
  const guess = location.origin + ((location.pathname.match(/^\/(?:sites|teams)\/[^\/]+/i) || [''])[0]);
  const H = { Accept: 'application/json;odata=nometadata' };
  const web = (await (await fetch(guess + '/_api/web?$select=Url', { headers: H })).json()).Url;
  const digest = (await (await fetch(web + '/_api/contextinfo', { method: 'POST', headers: H })).json()).FormDigestValue;
  const owners = await (await fetch(`${web}/_api/web/AssociatedOwnerGroup?$select=Id,Title`, { headers: H })).json();
  const g = ((await (await fetch(`${web}/_api/web/sitegroups?$select=Id,Title&$filter=Title eq 'CopyJet Spike csoport'`, { headers: H })).json()).value || [])[0];
  const ownerOf = async () => { const x = await (await fetch(`${web}/_api/web/sitegroups(${g.Id})?$select=Owner/Id,Owner/Title,Owner/PrincipalType&$expand=Owner`, { headers: H })).json(); return x.Owner ? `${x.Owner.Id}: ${x.Owner.Title} (típus ${x.Owner.PrincipalType})` : null; };
  const out = { web, group: g ? `${g.Id}: ${g.Title}` : 'nincs meg a CopyJet Spike csoport', ownersGroup: `${owners.Id}: ${owners.Title}` };
  if (g) {
    out.ownerBefore = await ownerOf();
    // Exactly what GroupProvider sends (generated by src/core/http/csomXml.ts setGroupOwnerBody):
    const xml = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CopyJet"><Actions><ObjectPath Id="10" ObjectPathId="4" /><ObjectPath Id="11" ObjectPathId="5" /><SetProperty Id="12" ObjectPathId="4" Name="Owner"><Parameter ObjectPathId="5" /></SetProperty><Method Name="Update" Id="13" ObjectPathId="4" /></Actions><ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="1" Name="Web" /><Property Id="3" ParentId="2" Name="SiteGroups" /><Method Id="4" ParentId="3" Name="GetById"><Parameters><Parameter Type="Int32">${g.Id}</Parameter></Parameters></Method><Method Id="5" ParentId="3" Name="GetById"><Parameters><Parameter Type="Int32">${owners.Id}</Parameter></Parameters></Method></ObjectPaths></Request>`;
    const r = await fetch(web + '/_vti_bin/client.svc/ProcessQuery', { method: 'POST', headers: { 'Content-Type': 'text/xml', Accept: 'application/json', 'X-RequestDigest': digest }, body: xml });
    const body = await r.text();
    out.c_status = r.status;
    out.c_error = (body.match(/"ErrorMessage":"([^"]*)"/) || [])[1] || null;
    out.ownerAfter = await ownerOf();
  }
  const json = JSON.stringify(out, null, 2);
  console.log(json);
  try { await navigator.clipboard.writeText(json); console.log('Copied to clipboard.'); } catch { console.log('Copy the JSON above manually.'); }
})();
```

**Mit várunk:** `c_status` = 200, `c_error` = `null`, az `ownerAfter` pedig a Tulajdonosok csoport (típus 8).

### C – eredmény (2026-09-24, Cél site): ✅

A `setGroupOwnerBody` által előállított kérés 200-at adott, `ErrorInfo` nélkül. A tulajdonos a „Zoltán Jakab (típus 1)” felhasználóról **„Cél Owners (típus 8)”**-ra változott. A `GroupProvider` ezt használja, és a beállítás után visszaolvassa a tulajdonost.

**Takarítás:** a Cél site-on a „CopyJet Spike csoport” csoport kézzel törölhető (*Webhelyengedélyek*).
