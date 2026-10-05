# CopyJet – felhasználói leírás

A CopyJet egy SharePoint Online-site részeit – listákat, tárakat, oszlopokat, nézeteket, tartalomtípusokat, SharePoint-csoportokat, modern lapokat és a navigációt, kérésre a tartalommal együtt – egy hordozható sablonfájlba menti, majd egy másik site-on (akár másik tenantban) újra létrehozza.

Két webpartból áll:
- **CopyJet Setup** – a *forrás* site egy lapján: kiválasztod, mi kerüljön a sablonba, és letöltöd a sablonfájlt.
- **CopyJet Install** – a *cél* site egy lapján: betöltöd a sablont, megnézed, mi fog történni, és telepíted.

Mindkettő a bejelentkezett felhasználó saját jogaival dolgozik: nem fér hozzá semmihez, amihez te sem.

## 1. Előkészületek

1. **Az app feltöltése (rendszergazda):** a `copyjet.sppkg` fájlt a tenant (vagy a site-gyűjtemény) alkalmazáskatalógusába kell feltölteni és engedélyezni. Tenantok közötti másoláshoz mindkét tenantban.
2. **Az app hozzáadása a site-okhoz:** a forrás és a cél site-on *Beállítások → Alkalmazás hozzáadása → CopyJet*.
3. **A webpartok elhelyezése:** a forráson egy lapra a „CopyJet Setup”, a célon egy lapra a „CopyJet Install” webpart. Érdemes olyan lapot választani, amelyet csak a tulajdonosok látnak.
4. **Jogok:**
   - Setup: olvasási jog a másolandó listákra, tárakra, lapokra.
   - Install: tulajdonosi (Teljes hozzáférés) jog a cél site-on. Az Install indulás előtt ellenőrzi, és ha hiányzik, nem enged tovább.
   - Egyedi listajogosultságok másolásához a cél site-gyűjtemény adminisztrátorának lenni ajánlott (lásd 5. pont).
   - Hivatkozásos fájlmásolásnál (3.2) a telepítőnek a forrás tárat is olvasnia kell.

## 2. Sablon készítése (Setup)

### 2.1 Kiválasztás

A webpart betöltéskor feltérképezi a site-ot, és fában mutatja: csoportok, site-oszlopok, tartalomtípusok, listák és tárak (alattuk az oszlopok, nézetek, egyedi jogosultságok), lapok, navigáció. A **Teljes site** mindent kijelöl. Ami nem másolható (pl. nem támogatott listatípus), szürkén, indoklással jelenik meg.

### 2.2 Opciók

Listánként és táranként, a sorra kattintva:
- **Csak szerkezet** (alapértelmezés) vagy **Szerkezet + tartalom** – listánál az elemek a mellékletekkel, tárnál a fájlok.
- **Verzióelőzmények másolása** – csak tárnál, tartalommal; alapértelmezetten ki.

Globális beállítások:
- **Sablon neve, leírása.**
- **Szerző, módosító és dátumok megőrzése** – a tartalom a forrás szerinti Létrehozta / Módosította értékekkel és dátumokkal jön létre (alapértelmezetten be).
- **Max. fájlméret** – ennél nagyobb fájl nem kerül a sablonba (figyelmeztetéssel kimarad).
- **Fájlok hivatkozással (azonos tenanton)** – lásd 2.4.
- **Formátum** – automatikus: tartalom nélkül `.json`, tartalommal `.zip`.

### 2.3 Összesítő és export

Az összesítő megmutatja, mi kerül a sablonba, és felajánlja a **hiányzó függőségeket** (pl. a lookup céllistáját, a lista által használt site-oszlopot vagy tartalomtípust, a lapon szereplő listát) – ezeket érdemes hozzáadni, különben a telepítésnél kimaradnak. Személyes adatot (szerzők, felhasználói mezők, tartalom) tartalmazó sablonnál figyelmeztet.

A **Sablon létrehozása** után a haladásjelző és a napló látszik; a végén a **Letöltés** gombbal mented a fájlt. A napló is letölthető.

### 2.4 Beágyazott vagy hivatkozásos fájlok

| | Beágyazott (alapértelmezés) | Hivatkozásos |
| --- | --- | --- |
| Hol vannak a fájlok | A `.zip`-ben | A forrás tárban; a sablon csak a listájukat viszi |
| Hova telepíthető | Bárhova, másik tenantba is | Csak a forrással azonos tenantba |
| Méret | A `.zip` a fájlokkal együtt nő; nagyon nagy tárnál a böngésző memóriája szab határt | A sablon kicsi; a másolást a SharePoint szerveroldalon végzi |
| Feltétel telepítéskor | Nincs | A forrásfájloknak még meg kell lenniük, és a telepítőnek olvasnia kell őket |

## 3. Telepítés (Install)

### 3.1 Betöltés

Húzd a sablonfájlt a webpartra, tallózd ki, vagy add meg az URL-jét. Tartalommal készült sablonnál a teljes `.zip`-et töltsd be. Az Install ellenőrzi a sablon szerkezetét és az ellenőrzőösszeget; ha a fájl az elkészítése óta megváltozott, figyelmeztet.

Ha ugyanennek a sablonnak egy korábbi telepítése ezen a site-on félbemaradt, egy sáv jelenik meg: **Folytatás** (a kész lépések nem futnak újra) vagy **Új telepítés**.

### 3.2 Leképezés

- **Felhasználók:** a sablonban szereplő személyeket (szerzők, személymezők, csoporttulajdonosok, jogosultságok) a cél site felhasználóihoz párosítja. Sorrend: kézi választás, CSV-tábla (`forrás;cél` soronként), azonos fiók, azonos e-mail, domaincsere. Akit egyik sem talál, ott választhatsz: **üresen hagyás + napló** vagy **helyettesítő felhasználó**.
- **Managed Metadata termek:** azonos tenantban azonosító szerint, másik tenantban a termcsoport / termkészlet / címke-útvonal szerint. A célon nem található termek értéke üres marad.

### 3.3 Előnézet

Minden elemnél látszik, mi történne: **új**, **azonos**, **eltérő** vagy **probléma** (indoklással). Eltérő elemnél az ütközéskezelés választható, egyenként vagy mindenre:
- **Kihagy** (alapértelmezés) – a meglévőhöz nem nyúl.
- **Frissít** – a hiányzó részeket pótolja, az eltérőket a sablon szerint állítja; **törölni soha nem töröl**.
- **Átnevez** – ahol támogatott, új néven hoz létre egy másikat.

### 3.4 Telepítés

A lépések függőségi sorrendben futnak (csoportok, oszlopok, tartalomtípusok → listák → oszlopok, nézetek, jogosultságok → elemek, fájlok → lookup-értékek → lapok → navigáció). A haladásjelző és az élő napló mutatja, hol tart.

- **A lapot ne zárd be** telepítés közben; ha mégis megtörténik, vagy a **Leállítás** gombot nyomod meg, a telepítés a következő betöltéskor folytatható (3.1). Az állapotot a cél site rejtett „CopyJetLog” listája tárolja, csak a tulajdonosok számára látható.
- Egy hibás lépés csak a tőle függő lépéseket állítja meg; a többi fut tovább.

### 3.5 Eredmény

Összesítés (létrehozott, frissített, kihagyott, hibás, folytatásból kész), a napló szűrhető nézete, és letöltés CSV-ben vagy JSON-ban. Hiba vagy figyelmeztetés esetén a napló kódja alapján a 6. pont segít.

Ugyanaz a sablon nyugodtan telepíthető többször ugyanoda: a meglévő elemeket felismeri, nem duplikál és nem ír felül (Kihagy módban).

## 4. Mit visz át a CopyJet

| Elem | Megjegyzés |
| --- | --- |
| Site-oszlopok, tartalomtípusok | Belső név és azonosító szerint; a Managed Metadata oszlop a termkészlethez kötve |
| Listák és dokumentumtárak | Egyéni lista, dokumentumtár; beállítások, mappák, oszlopok, nézetek. Az azonos nevű meglévő tárba (pl. „Documents” ↔ „Dokumentumok”) telepít |
| Listaelemek | Mellékletekkel; lookup, személy, Managed Metadata értékek leképezve; a számított oszlop értéke a célon számolódik újra |
| Fájlok | Metaadatokkal, kérésre verzióelőzménnyel; 10 MB felett darabolva |
| SharePoint-csoportok | Név, leírás, tulajdonos, beállítások, site-szintű jogosultsági szint; tagok nélkül |
| Egyedi listajogosultság | Csoportok és leképezett személyek a forrás szerinti szintekkel |
| Modern lapok | Szakaszok, webpartok, fejléc és képek; a site-ra és listákra mutató hivatkozások a célra átírva |
| Navigáció | Bal oldali és felső menü, célközönség, kezdőlap; meglévő menühöz hozzáfűz |

## 5. Ismert korlátok

- **Verzióelőzmény szerzője:** a fájl jelenlegi Létrehozta / Módosította adatai a forrás szerintiek, de a **korábbi verziók szerzője a telepítő** lesz (a SharePoint nem engedi utólag átírni). A verziók dátuma beágyazott módban a forrás szerinti.
- **Egyedi listajogosultságnál a telepítő:** az öröklés megszüntetésekor a SharePoint a telepítőt is felveszi Teljes hozzáféréssel. Ha a telepítő site-gyűjtemény-adminisztrátor és a forrás nem nevezte meg, a CopyJet leveszi; különben bent hagyja, nehogy kizárja magát (a napló jelzi).
- **Nem másolódik:** csoporttagság, elemszintű egyedi jogosultság, megosztási hivatkozások, klasszikus lapok, Power Automate folyamatok, a site maga (a cél site-nak léteznie kell), hub-navigáció.
- **Listatípusok:** egyéni lista, dokumentumtár és a Webhely lapjai; a többi (naptár, fórum, wiki, felmérés) kimarad, figyelmeztetéssel.
- **Egyedi webpartok:** a lapon lévő egyedi (nem Microsoft) webpart csak akkor működik a célon, ha a megoldása ott is telepítve van; az Install figyelmeztet.
- **Hivatkozásos fájlok:** csak azonos tenantban; másik tenantban a fájllépés kimarad.
- **Böngészőben fut:** nagy tartalomnál a telepítés hosszú lehet; a lapnak nyitva kell maradnia (megszakadás után folytatható).

## 6. Gyakori naplóüzenetek

| Kód | Jelentés | Teendő |
| --- | --- | --- |
| `LIST_MATCHED` | A tár egy meglévő, más nyelvű/nevű tárba került (pl. „Megosztott dokumentumok”) | Nincs |
| `FILES_KEPT`, „already present; skipped” | A célon már megvolt, nem írta felül | Nincs; frissítéshez Frissít mód |
| `ITEMS_TARGET_NOT_EMPTY` | A céllistában már vannak elemek; nem ír bele, hogy ne duplikáljon | Üres listába telepíts, vagy töröld a cél elemeit |
| `PRINCIPAL_NOT_FOUND`, `PRINCIPAL_FALLBACK` | Egy felhasználó nem található a célon; üres maradt vagy a helyettesítő került a helyére | Leképezés lépésben párosítsd (CSV / kézi) |
| `TERM_NOT_FOUND`, `TERM_SET_NOT_FOUND` | A term vagy termkészlet nincs meg a cél termtárban | Hozd létre a célon, majd telepíts újra Frissít módban |
| `ROLE_NOT_FOUND` | Egy egyedi jogosultsági szint nincs meg a célon | Hozd létre ugyanazzal a névvel, majd telepíts újra Frissít módban |
| `LIST_SECURITY_PRINCIPAL_MISSING` | Egy személy vagy csoport nincs a célon / nincs leképezve; a jogosultsága kimaradt | Leképezés, majd újratelepítés |
| `LIST_SECURITY_INSTALLER_KEPT` | A telepítő Teljes hozzáférése a listán maradt | Ha nem kell, kézzel vedd le |
| `FILE_TOO_LARGE`, `FILE_CHECKED_OUT` | A fájl a méretkorlát felett volt, vagy ki volt véve | Nagyobb korlát, illetve beadás után új sablon |
| `FILES_UNSUPPORTED` (`otherTenant`) | Hivatkozásos fájlok másik tenantba | Beágyazott módban készíts sablont |
| `PAGE_WEBPART_MISSING` | A lapon lévő egyedi webpart nincs telepítve a célon | Telepítsd a megoldását a célra |
| `ITEM_LOOKUP_UNRESOLVED` | A lookup céllista elemei nem ebben a futásban készültek, az értékek kimaradtak | Vedd fel a céllista tartalmát is a sablonba |
| `STEP_BLOCKED` | Egy lépés kimaradt, mert egy előfeltétele hibás volt | Javítsd az előző hibát, majd Folytatás |
| `STATE_SAVE_FAILED` | A futás állapota nem menthető; a telepítés megy tovább, de nem folytatható | Ellenőrizd a tulajdonosi jogot |

## 7. Adatvédelem

A tartalommal készült sablon személyes adatot tartalmazhat (nevek, e-mail-címek, szerzők, dokumentumok). Úgy kezeld, mint magát a tartalmat: csak az kapja meg, aki a forrást is láthatja, és a felhasználás után töröld.
