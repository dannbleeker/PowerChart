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

### The standing test run

What to do on **every** build that lands a real-host fix. Four items, ordered
by risk. Two are one click each and run themselves; the other two are a scroll
and a single click-and-edit — about three minutes of attention in total.

This used to be seven tests. Tests 1, 2, 3 and 7 were absorbed into the
self-test battery once `Slide.setSelectedShapes` turned out to have been
available since PowerPointApi 1.5, and the old test 4 (the `=SUM(A1:ZZ999)`
formula crash) was deleted outright: it is pure browser JavaScript, it never
needed PowerPoint, and it has been covered in CI all along by
`test/formula-blanks.test.ts`. The owner drives PowerPoint, the agent
fixes fallout — expect a real host to surface things the mocked tests can't,
because every Office.js assertion in this repo is against a fake.

**Before you start.** Wait ~2 minutes after the merge for the Pages deploy.
Open the pane and check the **build stamp** under the title is the commit you
mean to test — PowerPoint caches the pane aggressively, and a whole session can
otherwise go into testing code the host never fetched. Hard-reload if it is
older. Then tick **Verbose trace** in the Testing section and leave it on.

**If a run dies, the evidence is no longer lost.** The run log is downloadable
only once a run ends, and two rounds were lost to runs that did not end: one
wedged at 1819 seconds, one killed by PowerPoint's own *"Sorry, we ran into a
problem"* at 108 seconds. Neither produced a log. Two things cover that now:

- **Live steps** is on screen the whole time. Copy or screenshot it before
  reloading — it is the only thing that catches the very last moments.
- Every step is also written to **browser storage** as it happens. Reopen the
  pane after a crash and a red note offers **Download the crashed run**. Send
  that file; `npm run triage` reads it on its own, with no deck
  (`node scripts/triage.mjs powerchart-crashed-run.json`).

The storage copy survives a closed tab and a reloaded pane. It does not survive
clearing site data, and it keeps the last ~2000 steps — a longer run loses its
beginning, never its end.

| # | test | what it catches |
| --- | --- | --- |
| 1 | **Run host self-test.** One click, nine scenarios (below), and the Live steps list names each one as it starts. Chasing one failure? Set **Scenario** to it first — it runs that plus the two inserts it needs, in seconds rather than minutes. Read the verdicts, and if the run does not finish, read the last line. | Nearly everything that used to be tests 1, 2, 3 and 7 of this table. The battery now selects shapes itself (`Slide.setSelectedShapes`, PowerPointApi 1.5), stops its own run, and asks the host to render a slide before and after drawing so it can tell a chart that is *there* from a chart that is *visible*. A verdict of **skipped** is not a failure. |
| 2 | **Demo deck — both paths.** Path → **Both, one after the other** → **Insert demo deck**. ~6 s for the file half, then 1–2 minutes for the shape half. | The file half must report **38 of 38 complete** — anything less is a regression. The shape half being slower with some items short is the *measurement* of what the everyday path costs at 38× scale, not a defect. This is the one thing no battery can stand in for: it is the only test at real scale, and scale is what crashes the tab. |
| 3 | **Look at the deck.** Scroll the 38 slides the demo run added. | The judgement a machine does not have. The battery's visibility check answers "did anything render"; it does not answer "is this the right chart, drawn well". Look for charts off the slide edge, overlapping labels, and anything that is visibly not what the gallery shows. |
| 4 | **One selection round trip, by hand.** Click a chart. The pane says "A PowerChart is selected." Press **Edit it**, change a number, **Update chart**. Drag the chart, then edit it again. | The battery drives the same machinery through `setSelectedShapes`, which is Office.js selecting a shape — not a human clicking one. Those are the same call in theory. This is the check that they are the same in practice, plus the drag-delta round trip (`POWERCHART_ORIGIN`), which needs a real drag and so cannot be scripted at all. |

**Run it on a 4:3 deck at least once**, and on desktop PowerPoint at least once
per release — everything above is normally run on the web, which is where the
bugs have been, but it is not the only host and the battery is not a substitute
for a different one.

Then send back two files: **Download run log** (Testing section) and the deck
itself (File → Download a copy). The log carries the run's identity token, so
`npm run triage` joins it to the deck exactly rather than by guesswork.

What to read in the result: the `tagging failed` count (was 28 on the last slow
run; should be near zero), any line annotated `^ known host bug: office-js#…`
(Microsoft's, not ours — annotated automatically, don't chase it), the nine
self-test verdicts, and — new — the **`phases an error escaped`** block.

That block is the half of a failure that used to be missing. Office.js reports
what it refused (`errorLocation`, `code`); it cannot report what the add-in was
doing at the time, and reconstructing that from timestamps and call order is
what made the last three real-host failures cost a session each. Every error
now carries the phase it came out of — `at=drawing the chart's shapes`,
`at=writing the chart's config tag`, `at=reading slides 20-38 for charts` — in
its own message, in the trace as it happens, and tallied by `npm run triage`.
An error with no `at=` came from somewhere with no `step` around it, which is
itself worth reporting.

Record anything broken as issues; fix per the lockstep rules. Real-host
degradation paths that are *expected* (not bugs): radar fills are
outline-only, pattern fills render solid.

**A programmatic `setSelectedShapes` wedges the web host's selection
subsystem.** Measured on PowerPoint on the web, twice, on build `55011a3`:

- `Slide.setSelectedShapes([id])` is GA at PowerPointApi 1.5 and **takes the
  call** — no error, no refusal, `context.sync()` resolves.
- Every selection call after it goes silent. `getSelectedShapes` (the read
  behind *Edit it*) ran out its full 90-second budget; so did the
  `setSelectedSlides` that follows it. Nothing throws — the host simply never
  answers, and a run that would otherwise take seconds took 159.

Two selection bugs are already open against the web host — office-js#3083
(`setSelectedShapes([])` does not clear on web) and #3698 (no picture insert
while a shape is selected) — and this is the same family: the web host's
selection layer accepts writes it cannot then serve. It is Microsoft's, not
ours, and it is now **gated**: the *edit the chart the user selected* scenario
waits ten seconds rather than ninety, and reports **skipped** with the reason
rather than red. `npm run triage` annotates it as a known host bug
automatically.

**Nothing in the add-in is affected by it**, which is why it is a gate and not
a fix: the pane never selects a shape programmatically. It reads the selection
the *user* made, and a user's own click leaves the subsystem answering
normally. That is exactly what test 4 of the standing run checks by hand, and
on the web it is the only thing that can.

**Two more things that are expected and look alarming.** Neither is a defect:

- **A blank slide appears at the end of the deck and disappears again** while a
  chart redraws. That is deliberate. Updating a chart is the add-in's worst
  case on the web, and it is only bad because the slide is on screen — so the
  view is parked elsewhere for the redraw. A one-slide deck has nowhere to
  park, so one is made and removed afterwards. If a blank slide is ever left
  behind, *that* is the bug; the pane warns when it cannot remove it.
- **The view jumps to another slide and back** during an edit, for the same
  reason. Your own navigation wins: click away mid-redraw and you stay where
  you clicked.

**More on slide size.** It is read from the host now
(`PageSetup` at 1.10, an exported slide at 1.8, `getFileAsync` below that) and
drives both chart placement and the generated deck's declared size. Every
earlier session tested 16:9 only, where a wrong answer is invisible — on 4:3 it
shows up as charts off the right edge or a demo deck rescaled on insert. The
run log's `slide size` line records the value *and which rung produced it*;
`assumed` there means nothing answered.

**Wider feature sweep** — once per release, rather than per build:

1. Ribbon shows the PowerChart menu; pane opens; gallery renders.
2. Pie chart on a 1.10+ host (triangle-fan rotation), grouping on 1.8+.
3. **Use deck theme** on a 1.10+ host pulls the template's accent colors.
4. Elements (harvey ball, table with a total row) and Agenda insert.
5. Excel: select a range → Generate → paste JSON into PowerPoint pane →
   Import → chart matches.

### The host self-test

Six paths existed only as items on this list for a human to remember to try,
which in practice meant six separate sessions. They are now one button:

| scenario | what it proves |
| --- | --- |
| insert on top of an earlier run | the run token keeps two runs' slides apart, instead of one being deleted as the other's duplicate |
| two slides claiming one slot | the repair pass drops one copy and keeps a working one — not both, not the wrong one |
| edit a chart on the visible slide | the live-canvas redraw survives with the slide genuinely on screen |
| insert onto a slide that already has content | the everyday action — a chart drawn onto a slide that is not blank stays grouped and re-editable, and does not swallow what was already there |
| same scale across the deck | a deck-wide rescale empties nothing |
| explode a degraded picture | a picture keeps its config and can become native shapes again |
| edit the chart the user selected | the read behind *Edit it* — the only entry point a real user travels on |
| stop a run part-way | a stopped run adds nothing and leaves nothing behind claiming to be a chart |
| the chart is actually visible | the host's own render changed where the chart was drawn — not just that shapes exist |

**One more, and a full run deliberately leaves it out.** Pick **which selection
call wedges the host** from the Scenario menu to run it:

| scenario | what it proves |
| --- | --- |
| which selection call wedges the host | *which* call stops the web host answering — the one measurement that settles it |

It climbs a ladder from the least invasive selection call to the most — read the
selection, `setSelectedSlides`, read, `setSelectedShapes([id])`, read,
`setSelectedShapes([])`, read — and **stops at the first one that goes silent**,
reporting that call and the last one the host answered.

Two things about it are deliberate. It **stops** because after a wedge every
later call is silent too, so climbing on would report four timeouts and name
nothing — which is the ambiguity it exists to remove. And it is **picked only**
because its answer is meaningful solely on a host nothing else has touched:
`edit the chart the user selected` provokes the same wedge six scenarios
earlier, so inside a full run this would have reported silence on its own first
rung. (That was the first draft. A test caught it; a real host would have cost a
round.)

It reports **ok** whatever it finds. It is an experiment, not an assertion — a
host that answers every rung is a real result, and one worth knowing.

Each verdict says what was observed, not just pass/fail, and a scenario that
throws is recorded and the rest still run — a battery that stopped at the first
error would spend a whole session to learn one thing. A scenario the host
cannot run is reported as **skipped**, kept apart from a failure: "we did not
check" and "we checked and it is broken" send a diagnosis in different
directions.

It leaves its slides in the deck on purpose — save the file and hand it to
`npm run triage` with the log.

**What it cannot cover.** For a long time this paragraph said Office.js had no
way to select a *shape*. That was wrong — `Slide.setSelectedShapes` has been GA
since PowerPointApi **1.5** — and the belief cost this project four manual tests
that could have been buttons. It is scripted now.

What it still cannot cover is a **human** selecting: a real click, and the drag
that precedes an edit (`POWERCHART_ORIGIN`). Test 4 of the standing run is the
only coverage those get, which is why it is still manual. And on PowerPoint on
the **web** the scripted version does not work at all — see below — so on that
host the manual round trip is the *only* coverage the selection path has.

### Reading the demo-deck self-check (post-#212–#216)

The **Insert demo deck** button runs every chart kind, appends a results
slide, and posts a summary in the pane note. Ten harness-reliability PRs
since v0.2.0 mean the note now carries strictly more signal than "N of M
rendered":

- **`rendered`** — chart landed with every expected shape and was grouped.
- **`late-settled`** — a sync timed out but the shapes committed anyway;
  the harness read back the slide and trusted the count (no dup slide, no
  NOT COMPLETE stamp). Counted as rendered.
- **`rendered-partial`** — a sync threw with ≥85 % of the expected shapes
  on the slide. Counted as rendered; the fresh-context rescue groups what
  landed so the chart is still re-editable.
- **`ungrouped`** — chart shapes are on the slide but not one group. Not
  re-editable via the `POWERCHART_CONFIG` tag; a rescue attempt already
  ran. Flag for investigation.
- **`failed`** — under the 85 % gate after both attempts AND the
  unstamp+rescue path could not group what landed; slide carries the red
  NOT COMPLETE banner.
- **`BLANK: <title>`** — slide committed but readback showed zero shapes,
  and the slot tag names which item was on it (host lost the content).
- **`N of M results pages added`** — the run's own results slide
  paginated; each page is attempted independently, so a partial landing
  no longer drops the record.
- **`addsLostAtCommit=N`** — `addSlides` confirmed N `slides.add()` calls
  never landed even after its own retry. Correlates with the office-js
  bug documented in `OFFICE_JS_LOST_ADDS.md`. `addsIssued − slidesAdded`
  is the wider gap.

A clean run reports every chart rendered + grouped, no ungrouped/blank/
`addsLostAtCommit`. The full console.table dump under **F12** carries
`shapes`, `status`, `grouped`, `lateOutcome` and `ms` per item. Better
still, **Download run log** writes the whole run to JSON — both insert
paths, the settled repair verdicts, and the activity trace when Verbose
trace was on. That file is the right attachment for a Phase-2 regression.

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
