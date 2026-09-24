# CopyJet – projektútmutató Claude Code-hoz

Ez a fájl minden munkamenet elején betöltődik. Rövid, kötelező szabályok; a részletek a `docs/` mappában vannak.

## Mi ez a projekt

SharePoint Online SPFx megoldás, amely egy site egészét vagy részeit (listák, tárak, SP-csoportok, modern lapok, navigáció – opcionálisan tartalommal) hordozható sablonba menti, és egy másik (akár másik tenantban lévő) site-on újra létrehozza.

- **Setup webpart** – a forrás site lapján: feltérképezés → kiválasztás → opciók → sablon (`manifest.json` vagy `.zip`) exportja.
- **Install webpart** – a cél site lapján: sablon betöltése → validálás → felhasználó/term-leképezés → előnézet (diff) → telepítés függőségi sorrendben → napló.

Olvasd el feladat előtt a releváns részt:

| Fájl | Mikor |
| --- | --- |
| `docs/01-rendszerterv.md` | Követelmények, architektúra, telepítési sorrend, döntések (13. fejezet) |
| `docs/02-fejlesztoi-leiras.md` | Modulok, interfészek, API-k, fejlesztési sorrend (13. fejezet = feladatlista) |
| `schema/copyjet.v1.schema.json` | A sablon szerződése – extractor és provider is ehhez igazodik |
| `schema/examples/` | Kitöltött minták, tesztfixture-nek is |
| `docs/ui/index.html` | Képernyővázlatok (Setup 4 + Install 5 lépés), böngészőben megnyitható |

## Rögzített technológiai döntések (ne térj el engedély nélkül)

- SPFx **1.23.2**, Heft-alapú build (nincs gulp). Node.js **22 LTS**. TypeScript ≤ 5.8.
- React **17.0.1** pontos verzióval (`--save-exact`), Fluent UI az SPFx-szel szállított verzióban.
- PnPjs **4.21.0** (`@pnp/sp`, `@pnp/queryable`; opcionálisan `@pnp/graph` 4.21.0). Minden PnPjs csomag azonos verzión.
- AJV 8 (`ajv/dist/2020`), `ajv-formats`, `ajv-errors`; JSZip; Jest + ts-jest.
- `package-solution.json`: `skipFeatureDeployment: false` (site-onkénti app-hozzáadás), `includeClientSideAssets: true`, nincs `webApiPermissionRequests` amíg a Graph-funkció nincs bekapcsolva.
- Nincs backend; minden a bejelentkezett felhasználó delegated jogaival fut.
- A hosted workbench 2026-12-01-jén megszűnik: helyi tesztelés `npm start` + SPFx Debug Toolbar egy valódi SharePoint lapon.

## Architektúra-szabályok

- `src/core/**` UI-független: **tilos** React, Fluent UI vagy `@microsoft/sp-webpart-base` import. A `WebPartContext`-et csak a `core/http` kapja meg.
- Minden artefaktumtípus = egy `*Extractor` + egy `*Provider` + tesztek, azonos névtővel (`ListExtractor` / `ListProvider`).
- A sablon csak adat: site-függő érték (URL, GUID, user ID) **csak tokenként** kerülhet bele (`{site}`, `{listkey:X}`, `{fieldid:X}`, `{principal:key}` …). Új tokent a `tokenizer` regiszterében vegyél fel.
- Providerek: `apply` előtt mindig `diff`; idempotensek; **soha nem törölnek** a cél site-on.
- Minden SharePoint-hívás a `core/http`-n megy át (retry 429/503-ra `Retry-After` szerint, max. 5 próba, párhuzamosság 4, batch ≤ 100).
- Minden aszinkron core-függvény fogad `AbortSignal`-t.
- A séma a szerződés: ha egy extractor új mezőt ad, először a sémát bővítsd (alverzió `1.x`, migráció a `core/schema`-ban), a TypeScript típusokat a sémából generáld (`json-schema-to-typescript`), és frissítsd a `schema/examples/` mintákat.
- UI-szövegek a `loc/` fájlokban (hu-hu, en-us), a kódban nincs beégetett szöveg. Kódazonosítók és kommentek angolul, dokumentáció magyarul.
- Alapértelmezések (felhasználói döntés): tartalommásolás ki, csoporttagság ki, verzióelőzmény ki, navigáció ütközésnél hozzáfűzés.

## Munkamód

- Egyszerre egy fázis, a `docs/02-fejlesztoi-leiras.md` 13. fejezetének feladatlistája szerint; a kész tételeket ott pipáld ki (`- [x]`).
- Minden logikai lépés után: `npm run build` és `npx jest` fusson hibátlanul, mielőtt továbblépsz.
- A felhasználó (Zoli) végzi a SharePoint-oldali lépéseket: bejelentkezés, `.sppkg` feltöltése az app catalogba, webpart kihelyezése, kattintásos tesztek. Fázis végén írj rövid, lépésenkénti **tesztforgatókönyvet**, amit ő végigkattint, és kérd vissza az eredményt / a napló exportját.
- Ne futtass tenantot módosító parancsot (pl. `m365 spo app add/deploy`) a felhasználó kifejezett kérése nélkül. Ne tárolj jelszót, tokent vagy tenant-URL-t a repóban (helyi értékek: `.env.local`, gitignore-olva).
- Commit csak kérésre; Conventional Commits formátum (`feat(core): …`, `fix(install): …`).
- Kétes API-viselkedésnél (pl. `CanvasContent1` betöltése, verziók szerzője, `createCopyJobs`) előbb kis, izolált kísérletet javasolj, és az eredményt írd a `docs/spikes/` mappába.

## Parancsok

```bash
node -v                      # v22.x kell
npm ci                       # függőségek
npm run build                # Heft build
npm start                    # helyi szerver a Debug Toolbarhoz
npx jest                     # egység- és szerződéstesztek
npx heft package-solution    # → sharepoint/solution/copyjet.sppkg (a scaffold package.json scriptjei az irányadók; élesre production builddel)
```

## Első feladat (0. fázis indítása)

1. Ellenőrizd: `node -v` (22.x), `npm -v`, git. Ha a Node nem 22, szólj, ne folytasd.
2. Scaffold a repó gyökerébe, a meglévő `docs/`, `schema/`, `CLAUDE.md` megtartásával: SPFx 1.23.2, solution neve `copyjet`, React webpart `CopyJetSetup`, majd második webpart `CopyJetInstall` ugyanabba a solutionbe (Yeoman `@microsoft/generator-sharepoint@1.23.2` vagy `@microsoft/spfx-cli`). Ha a generátor nem enged nem üres mappába, generálj ideiglenes mappába és mozgasd át.
3. Rögzítsd a verziókat: `react@17.0.1 react-dom@17.0.1 --save-exact`, `@pnp/sp@4.21.0 @pnp/queryable@4.21.0 --save-exact`, `ajv ajv-formats ajv-errors jszip`, dev: `jest ts-jest @types/jest json-schema-to-typescript`.
4. Hozd létre a `src/core/*` mappastruktúrát (lásd fejlesztői leírás 1. fejezet), `src/shared/components/`, `tests/`.
5. `core/schema`: AJV validátor a `schema/copyjet.v1.schema.json`-ra, `$defs/itemsFile|filesMetaFile|pageFile` külön validátorokkal; teszt: a `schema/examples/` fájlok érvényesek, 3 szándékosan hibás minta elutasítva.
6. `core/model`: típusgenerálás a sémából (`npm run gen:types` script) + kézi kiegészítő típusok (`ArtifactKind`, `IArtifactRef`, `IExtractor`, `IProvider`).
7. `core/http` (`createSp`, `CopyJetBehavior`, `runBatched`, `limitConcurrency`), `core/logger`, `core/tokenizer` (alap tokenek + tesztek), `core/packager` (`.json` írás/olvasás).
8. A két webpart jelenjen meg egy egyszerű „CopyJet Setup / Install – 0. fázis” felülettel, és a Setup gombnyomásra írja ki a site címét és a listák számát (PnPjs-kapcsolat próbája).
9. Zárásként: build + tesztek zöldek, a feladatlista kipipálva, és egy tesztforgatókönyv Zolinak (sppkg feltöltése, app hozzáadása, webpart kihelyezése, mit kell látnia).
