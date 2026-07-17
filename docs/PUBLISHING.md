# Publishing runbook — get PowerChart live in PowerPoint

Instructions for a Claude (Opus 4.8) session working with the repo owner to
take PowerChart from "feature-complete on main" to "usable inside
PowerPoint, with the Claude skill active". Read `CLAUDE.md` first — the
working conventions there (lockstep rule, branch flow, auto-merge policy,
visual QA) apply to every change you make here.

Legend: **[agent]** = you can do it with repo access; **[owner]** = needs
the owner's click (GitHub settings, PowerPoint UI, claude.ai account).
Do the phases in order — later phases depend on the hosted URLs.

---

## Phase 0 — Preconditions ✅ done

1. **[owner] Make the repo public** — ✅ done.
2. **[agent] Pre-publication sweep** — ✅ done: no secrets/keys/tokens
   (`git grep` clean; every "token" hit is benign code), no env/credential
   files, all sample/showcase data is invented dummy data, `npm test` green.
3. **[owner] Post-public hygiene** — ✅ done (Dependabot + CodeQL, description
   + topics).
4. **[agent] Branch protection** — ✅ done: ruleset "main: require CI green"
   (active) requires the `test` check and blocks force-push and deletion of
   `main`. Only `test` is required because it is the only check that runs on
   `pull_request` — `build`/`deploy` (Pages) and `release` fire on push/tag
   only, so requiring them would deadlock every PR. Repo admins keep an
   `always` bypass, so the owner can never be locked out of his own default
   branch.

## Phase 1 — Host the add-in on GitHub Pages ✅ agent work landed

Office add-ins load from an HTTPS URL; the dev manifests point at
`https://localhost:3000`. The site is hosted on GitHub Pages under a **custom
domain**, `https://powerchart.struktureretsundfornuft.dk/`. Because a custom
domain serves the project site from its **root**, the bundle base is `/`
(no `/PowerChart/` path segment) — the prod-manifest URLs are just
`https://powerchart.struktureretsundfornuft.dk/…`.

1. **[agent] Build for Pages** — ✅ `npm run build:pages`
   (`scripts/pages-postbuild.mjs`): runs the prod-manifest gen, `tsc`, a
   root-base `vite build`, then copies the manifest-referenced ribbon icons
   into `dist/assets/`. Emits `index.html`, `src/taskpane/taskpane.html`,
   `src/excel/excel.html`, `assets/icon-*.png`, and the static `public/` files
   (`CNAME`, `privacy.html`, `terms.html`) which Vite copies verbatim.
   > Gotcha found & fixed: Vite doesn't bundle `assets/icon-*.png` (they're
   > referenced only by the manifests), so without the copy step the hosted
   > icon URLs 404. `pages-postbuild.mjs` copies them; the `CNAME` and legal
   > pages ride along from `public/`.
2. **[agent] Deploy workflow** — ✅ `.github/workflows/pages.yml`: on push to
   `main`, `npm ci` → `npm run build:pages` → `upload-pages-artifact` (path
   `dist`) → `deploy-pages`, with `pages: write` / `id-token: write`.
3. **[owner] Enable Pages + custom domain** — ✅ done (Source: GitHub Actions;
   domain `powerchart.struktureretsundfornuft.dk`). Confirm **Enforce HTTPS**
   is checked once the cert provisions.
4. **[agent] Production manifests** — ✅ `scripts/build-manifest.mjs` rewrites
   `https://localhost:3000` → the custom-domain origin into
   `manifest-prod.xml` / `manifest-excel-prod.xml` (committed; `--check` mode
   gates staleness in `ci.yml`; regenerated + attached to releases in
   `release.yml`). Both GUIDs (`b7f6d3a2…`, `c8a7e4b3…`) preserved; 0 localhost
   URLs survive. (`office-addin-manifest validate` couldn't run in the sandbox
   — no network for the install — so validate once locally when convenient.)
5. **[agent/owner] Smoke-test the deployment**: after the first Pages run,
   `curl -sI https://powerchart.struktureretsundfornuft.dk/src/taskpane/taskpane.html`
   → 200, and the icons under `/assets/icon-*.png`. Load the demo gallery URL
   in a browser to confirm assets render.

## Phase 2 — Sideload in PowerPoint ([owner], agent assists)

Pick the platform(s); the manifest file is `manifest-prod.xml` from Phase 1
(attached to the latest release, or in the repo).

- **PowerPoint on the web** (fastest validation): open a deck on
  office.com → Home ▸ Add-ins → **More add-ins** → **My Add-ins → Upload My
  Add-in** → pick `manifest-prod.xml`.
- **Windows**: easiest supported route is the same Upload dialog (newer
  builds), else the shared-folder catalog: put the manifest in a folder,
  share it (`\\machine\manifests`), add it under File → Options → Trust
  Center → Trusted Add-in Catalogs, restart PowerPoint, Insert → My
  Add-ins → Shared Folder.
- **Mac**: copy the manifest to
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` and
  restart PowerPoint; it appears under Insert → My Add-ins.
- **Excel companion**: same procedure in Excel with
  `manifest-excel-prod.xml`.

**Validation checklist** (owner drives PowerPoint, agent fixes fallout —
expect the first real-host run to surface issues the mocked tests can't):

1. Ribbon shows the PowerChart menu; pane opens; gallery renders.
2. Insert a stacked chart → native grouped shapes appear; move/resize one
   bar by hand (proof of editability).
3. Select the chart → "Edit it" banner → change data → **Update chart**
   replaces in place.
4. Pie chart on a 1.10+ host (triangle-fan rotation), grouping on 1.8+.
5. **Use deck theme** on a 1.10+ host pulls the template's accent colors.
6. Elements (harvey ball, table with a total row) and Agenda insert.
7. Excel: select a range → Generate → paste JSON into PowerPoint pane →
   Import → chart matches.

Record anything broken as issues; fix per the lockstep rules. Real-host
degradation paths that are *expected* (not bugs): radar fills are
outline-only, pattern fills render solid.

## Phase 3 — Activate the Claude skill ([owner])

1. Download `powerchart-charts.zip` from the latest release (the rolling
   [`skill-latest`](../../releases/tag/skill-latest) is rebuilt on every
   merge).
2. claude.ai → Settings → Capabilities → **Skills** → upload the zip.
3. Test from any Claude surface: *"Make me an EBITDA bridge: FY23 86,
   Volume +14, Price +9, Cost −12, FX −4, FY24 total"* → expect a .pptx
   with native shapes. Then test inside **Claude for PowerPoint** (the
   add-in from AppSource) — skills enabled in settings are available there,
   which closes the loop: Claude builds PowerChart charts directly in the
   user's deck.

## Phase 4 — Cut the release ✅ v0.2.0 done

1. **[agent]** ✅ done — `gh workflow run release.yml -f version=v0.2.0`
   (the workflow creates the tag; the git proxy in remote sessions can't push
   tags). Trigger it from a green `main`: the release job runs `npm test` but
   **not** typecheck or the coverage thresholds, so it is a weaker gate than
   `ci.yml` and trusts the branch it builds from.
2. **[agent]** ✅ done — README carries the live gallery link and an install
   section pointing at the prod manifest download.
3. **[agent]** ✅ done — CLAUDE.md's "Pending / user-gated" list now names
   only what is genuinely still owner-gated.

**v0.1.0 shipped the DEV manifests** (`manifest.xml`, pointing at
`https://localhost:3000`) because it predates the Phase 1 change that attaches
the prod pair — so the README's install path was broken for anyone who tried
it: the file it names, `manifest-prod.xml`, was not in the release at all.
v0.2.0 fixes that. If a future release ever ships a manifest again, check the
asset list, not just the workflow file — the workflow was right for 12 days
while the published release stayed wrong.

## Distribution beyond sideloading (later, optional)

- **Org-wide (BESTSELLER)**: a Microsoft 365 admin deploys the manifest
  centrally via Admin Center → Settings → Integrated apps → Upload custom
  app. No store review; appears for chosen users automatically. Fastest path
  for internal use — recommended before attempting the public store.
- **AppSource** (public store): requires a Partner Center account (free for
  Office Store apps) and Microsoft validation (works on every claimed platform,
  WCAG, privacy + terms + support URLs). Substantial process; only worth it if
  PowerChart should be publicly installable. Prep is staged:
  - **[agent, done]** Hosted **privacy** (`/privacy.html`) + **terms**
    (`/terms.html`) pages (in `public/`, built to the site root), a
    trademark-clean store listing in `docs/STORE-LISTING.md`, and the
    store-facing manifest `<Description>` reworded off the "think-cell" mark.
  - **[owner]** Create the Partner Center account, produce the listing images
    (300×300 logo + screenshots), run `office-addin-manifest validate`, then
    submit. Full checklist in `docs/STORE-LISTING.md`.
  - ⚠️ **Trademark:** keep everything store-facing (name, description,
    screenshots) free of the "think-cell" mark — internal docs may keep it.

## Known constraints to keep in mind

- Requirement sets: shapes need PowerPointApi **1.4+** (Win 2207+, Mac
  16.62+, web; not iPad); grouping 1.8, re-edit tags 1.3, pie rotation 1.10,
  theme colors 1.10. The pane degrades gracefully below each.
- Pages is static HTTPS — exactly what an add-in needs; no server code, no
  auth, no cost. If the repo must stay private instead, any static HTTPS
  host works (Azure Static Web Apps free tier, Cloudflare Pages) — only the
  base URL in the prod manifests changes.
