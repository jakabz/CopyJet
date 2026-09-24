# CopyJet

SharePoint Online SPFx megoldás site-ok (vagy részeik) hordozható sablonba mentésére és újra létrehozására – két webpart: **Setup** (forrás site) és **Install** (cél site).

- Rendszerterv: [docs/01-rendszerterv.md](docs/01-rendszerterv.md)
- Fejlesztői leírás: [docs/02-fejlesztoi-leiras.md](docs/02-fejlesztoi-leiras.md)
- Sablon séma: [schema/copyjet.v1.schema.json](schema/copyjet.v1.schema.json) · minták: [schema/examples/](schema/examples/)
- UI-vázlatok: [docs/ui/index.html](docs/ui/index.html)
- Claude Code útmutató: [CLAUDE.md](CLAUDE.md)

**Követelmények:** Node.js 22 LTS, SPFx 1.23.2 (Heft), PnPjs 4.21.0, egy SharePoint Online tenant app catalogkal és egy forrás- és egy cél-site.
