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

## Phase 0 — Preconditions

1. **[owner] Make the repo public** (Settings → General → Danger Zone →
   Change visibility). GitHub Pages on the free tier requires it, and Pages
   is the zero-cost HTTPS host the rest of this runbook assumes.
2. **[agent] Pre-publication sweep** (should already hold, verify anyway):
   - `git grep -iE "secret|api[_-]?key|token" -- ':!package-lock.json'` — nothing sensitive.
   - All sample/showcase data is invented dummy data (see CLAUDE.md).
   - `npm test` green on main.
3. **[agent] Post-public hygiene** (was deferred until this moment; see
   CLAUDE.md "Pending"): enable Dependabot alerts + security updates and
   CodeQL default setup (Settings → Security, owner may need to click),
   propose a branch-protection ruleset for `main` (require CI green), and
   give the repo a description + topics (`powerpoint`, `office-addin`,
   `think-cell`, `charts`, `claude-skill`).

## Phase 1 — Host the add-in on GitHub Pages

Office add-ins load from an HTTPS URL; the manifests currently point at
`https://localhost:3000` (dev server). Replace that with Pages hosting.

1. **[agent] Build for Pages.** The site must be served from
   `https://<owner>.github.io/PowerChart/`, so the bundle needs that base
   path: build with `vite build --base=/PowerChart/` (add a
   `build:pages` npm script). Verify `dist/` contains `index.html` (demo
   gallery), `src/taskpane/taskpane.html`, `src/excel/excel.html`, and
   `assets/icon-*.png` — the taskpane path segment matters because the
   manifests reference it.
2. **[agent] Add a deploy workflow** `.github/workflows/pages.yml`:
   on push to `main` → checkout, `npm ci`, `npm run build:pages`,
   `actions/upload-pages-artifact` (path `dist`), `actions/deploy-pages`.
   Permissions: `pages: write`, `id-token: write`. Keep it separate from
   `ci.yml`.
3. **[owner] Enable Pages**: Settings → Pages → Source: **GitHub Actions**.
4. **[agent] Produce production manifests.** Don't edit the dev manifests —
   generate `manifest-prod.xml` / `manifest-excel-prod.xml` by replacing
   every `https://localhost:3000` with
   `https://<owner>.github.io/PowerChart` (a small
   `scripts/build-manifest.mjs` keeps them in lockstep with the dev ones;
   wire it into `npm run build:pages` and attach the prod manifests to
   releases in `release.yml`). Rules:
   - **Keep the GUID** (`<Id>`) stable across updates — changing it makes
     PowerPoint treat it as a different add-in.
   - The Excel manifest has a different GUID; keep that one too.
   - Validate both: `npx office-addin-manifest validate manifest-prod.xml`.
5. **[agent] Smoke-test the deployment**: after the Pages run,
   `curl -sI https://<owner>.github.io/PowerChart/src/taskpane/taskpane.html`
   → 200, and the icons under `/assets/`. Load the demo gallery URL in a
   browser screenshot (Playwright is preinstalled) to confirm assets render.

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
4. Pie chart on a 1.9+ host (triangle-fan rotation), grouping on 1.8+.
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

## Phase 4 — Cut the release

1. **[agent]** Confirm main is green, then trigger the **Release** workflow
   via `workflow_dispatch` with input `version: v0.2.0` (the git proxy in
   remote sessions can't push tags; the workflow creates the tag). It
   attaches the skill zip, both manifests (add the prod manifests, Phase 1
   step 4), and the showcase deck.
2. **[agent]** Update README with the Pages URLs: live demo gallery link
   and "install" section pointing at the prod manifest download.
3. **[agent]** Move the finished items out of CLAUDE.md's "Pending /
   user-gated" list and out of this runbook's open questions.

## Distribution beyond sideloading (later, optional)

- **Org-wide (BESTSELLER)**: a Microsoft 365 admin deploys the manifest
  centrally via Admin Center → Settings → Integrated apps → Upload custom
  app. No store review; appears for chosen users automatically.
- **AppSource** (public store): requires a Partner Center account and
  Microsoft validation (WCAG, privacy URL, support URL). Substantial
  process; only worth it if PowerChart should be publicly installable.

## Known constraints to keep in mind

- Requirement sets: shapes need PowerPointApi **1.4+** (Win 2207+, Mac
  16.62+, web; not iPad); grouping 1.8, re-edit tags 1.3, pie rotation 1.9,
  theme colors 1.10. The pane degrades gracefully below each.
- Pages is static HTTPS — exactly what an add-in needs; no server code, no
  auth, no cost. If the repo must stay private instead, any static HTTPS
  host works (Azure Static Web Apps free tier, Cloudflare Pages) — only the
  base URL in the prod manifests changes.
