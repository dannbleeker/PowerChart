# PowerChart — project memory

Open-source think-cell clone: a PowerPoint charting add-in whose charts are
**native, editable shapes** (never pictures), plus a Claude Agent Skill that
renders the same charts headlessly. 25 chart kinds, think-cell's signature
decorations, and "the good chart" design formalia are all implemented — see
the README feature table for the authoritative list.

## Architecture in one paragraph

`ChartConfig` (pure JSON) → `buildChart()` in `src/core` (pure TypeScript, no
Office imports) → renderer-agnostic scene graph (`rect | line | text | ellipse
| wedge | chevron | arrowhead | polygon | symbol` nodes, coordinates in points) → three
renderers: SVG (`src/render/svg.ts`, preview + tests), Office.js
(`src/render/powerpoint.ts`, live add-in; config persisted in a
`POWERCHART_CONFIG` shape tag for re-editing), and PptxgenJS
(`skill/scripts/render-pptx.mjs`, the skill's headless output). Details:
`docs/ARCHITECTURE.md`; the research that shaped it: `docs/RESEARCH.md`.

## The lockstep rule (CI-enforced — do not skip)

Any feature change must update, in the same PR:

1. **Skill docs** — `skill/SKILL.md` + `skill/reference.md`
   (`test/skill-docs.test.ts` fails on missing kinds/rows/keys).
2. **Showcase deck** — extend `scripts/build-showcase.mjs`, run
   `npm run showcase`, commit the regenerated `examples/showcase.json` +
   `showcase.pptx` (`test/showcase.test.ts` + a CI byte-diff staleness gate).
3. **User manual** — `docs/MANUAL.md` (`test/manual.test.ts` fails on
   missing kinds/rows/controls/elements).
4. **README feature table.**

## Commands

```bash
npm run dev        # gallery + pane at localhost:3000
npm test           # full suite (1500+); npm run coverage enforces thresholds
npm run typecheck
npm run showcase   # regenerate the showcase deck (required after feature work)
npm run skill      # build skill-dist/powerchart-charts.zip
```

## Working conventions (established with the repo owner)

- **Answer in caveman style** (the `caveman` skill, `full` level) from the first
  reply of every session in this repo, without being asked. Terse: drop
  articles and filler, fragments fine, keep every technical fact. Code,
  commits, PR bodies and user-facing docs stay normal prose — the style is for
  chat only. Drop it for security warnings, destructive-action confirmations,
  and any multi-step sequence where dropping conjunctions could be misread.
  "stop caveman" turns it off for that session.
- **Branch flow**: develop on the session's designated `claude/*` branch (the
  task prompt names it); after each merge, reset it onto `origin/main`
  (`git checkout -B <branch> origin/main`) — never stack on merged history.
  One PR per increment.
- **Auto-merge is authorized**: once CI is green on the exact pushed commit
  (verify `head_sha` matches local HEAD), merge the PR to main without asking.
- **Snapshots** (`test/snapshots.test.ts`) freeze every sample chart's SVG.
  Only update (`vitest -u`) after reviewing renders visually — screenshot via
  Playwright (`/opt/pw-browsers/chromium`) and inspect before accepting. CI now
  does the part a machine can: `npm run visible-charts` rasterises every sample
  in a real browser and fails on a chart that is drawn but not visible
  (white-on-white, collapsed to zero, off-canvas, one flat block, or missing over
  half its usual ink). Not a pixel diff — those numbers are measurements, and
  `--update` re-records `test/fixtures/chart-ink.json` after a deliberate change.
- **Visual QA is part of done**: render new features to SVG → PNG and look at
  them; several real bugs were only caught this way.
- **A regression test must be proven to fail without its fix.** Stash the
  source file, re-run, confirm the new test goes red, restore. A guard that
  passes against the pre-fix file is not a guard — it is decoration. Check
  _which_ assertion went red, too: a guard can fail for the wrong reason and
  still look proven (one here failed on a slide-swap rung that records nothing
  either way; another compared a `min()` both sides answered 0 for).
- **When two explanations fit the evidence, ask — do not reason.** A probe
  answer, a verdict, a log line can be about the host or about the thing that
  asked. Reasoning about which cost this project two full sheets and a session;
  adding the variant that separates them settled it in one run
  (`shape-add-fresh-slide-proxy` / `-held-` / `-positional-`, and again with
  `shapes-items-via-positional-slide`). Any answer that could be about the probe
  gets a partner question, and `KNOWN_DIVERGENCES` is where one waits while its
  re-run is pending. A probe can now carry its own partner: `Probe.follow` fires
  a second question in the same run when the first answer admits two readings,
  which is the rule automated — the reasoning was never the expensive part, the
  round trip was.
- **The fake is gated against a real host in CI** — `test/host-contract.test.ts`
  diffs `FAKE_BASELINE` against the committed sheet in
  `test/fixtures/host-answers-web.json`. A new divergence fails there unless it
  is declared in `KNOWN_DIVERGENCES` with a reason. When a fresh sheet arrives:
  replace the fixture, run the suite, deal with what goes red. Do not edit the
  fixture by hand — it is a recording, not a preference.
- **The office-js tracker is swept weekly** by
  `.github/workflows/office-js-watch.yml`, which reports only issues touching
  APIs this repo calls that are not yet in `KNOWN_ISSUES`
  (`scripts/office-js-watch.mjs`). When one is triaged, add it to that table
  **with what was done about it** — including "no exposure", which records that
  somebody checked. Anything left out comes back next Monday. Five of the guards
  in this repo came from one manual sweep; this is so the next one is not luck.
- **Stale documentation is a defect — fix it when you find it, in that turn.**
  Don't file it, don't mention it and move on. This applies to comments that
  justify a design with a claim that is no longer true, which is the expensive
  kind: `placement.ts` asserted that no requirement set exposes slide
  dimensions, and that single false sentence is why charts could only ever be
  placed vertically. If the fix is genuinely out of scope, say so explicitly
  rather than leaving the claim standing.
- **Test files are named by topic, never by increment** (`test/README.md` has
  the map). No `batch-N` / `bug-hunt-N` / `coverage-*` grab-bags: a test
  belongs with the thing it tests, not with the reason it was written.
- **Three sweeps run on a schedule, none of them gating a PR** — the office-js
  tracker (above), plus `.github/workflows/quality-sweep.yml`: a **flake hunt**
  (the suite three times under CPU load, reporting any test that disagreed with
  itself — `scripts/flaky.mjs` tells that apart from a suite that is red the
  same way every time) and a **mutation run** over `src/core` (`npx stryker run`,
  scoped by `vitest.mutation.config.ts`). Mutation is the rule below, automated:
  it answers "which assertions are decorative" for the whole engine at once,
  where the stash-and-re-run answers it for one. Not required checks on purpose —
  minutes-long jobs in front of every merge get switched off after the first bad
  week.
- **When moving tests between files, pin the total first and check it after.**
  A reorg once silently deleted 43 tests — the suite still went green because
  the count was never compared. `npx vitest run | grep "Tests "` before and
  after; if the number moved and you did not add or remove a case, stop. CI now
  holds the floor: `scripts/test-count.mjs` fails when the total drops below the
  recorded mark, which rises on its own. A deliberate drop is re-recorded with
  `--update`, so it lands in the diff where a reviewer sees it.
- **Releases**: merges to main refresh the rolling `skill-latest` prerelease.
  Versioned releases via the Release workflow's manual dispatch (the git proxy
  rejects tag pushes — dispatch creates the tag server-side).
- **Flag manifest re-installs to the owner**: the add-in is hosted on GitHub
  Pages (`powerchart.struktureretsundfornuft.dk`) and the sideloaded
  `manifest-prod.xml` only points at stable URLs — so code/pane/chart changes
  ship through `main` → Pages with **no** re-install. But when a change touches
  the **manifest itself** (ribbon buttons/menu items, `Permissions`,
  requirement sets, `DisplayName`/`Description`, icon references, or — never do
  this — the `<Id>` GUID), the owner MUST re-sideload the updated manifest in
  PowerPoint for it to take effect. Always tell the owner explicitly in that
  PR/turn: "⚠️ this needs a manifest re-install in PowerPoint" and say why.
  CI validates all four manifests with Microsoft's own tool (the `manifest` job,
  kept out of `test` because it calls a Microsoft service and `test` is the only
  required check). It found `<Version>0.1.0` on its first run — Office rejects
  anything below 1.0 — so the version is `1.0.0.0` and independent of the npm
  package version; `test/manifest.test.ts` pins that offline, along with the
  `<Id>` GUIDs and "no localhost in a prod manifest".

## Gotchas

- Office.js has **no freeform paths**: pies are triangle fans, radar/polygon
  fills degrade to outlines in the live add-in (the skill's pptx output gets
  real filled `custGeom` polygons), pattern fills are SVG-only (solid in PPT).
- **A shape the user selected must be let go of before anything is drawn.** On
  the web, `addTextBox` DELETES the selected shape (office-js#2775) and a picture
  cannot be inserted while one is selected (#3698) — and the insert path reads
  the selection's bounds precisely so it can place the chart there. Reading the
  selection and holding it are different things; `dropShapeSelection` runs
  between them. Ordering is the whole protection, and the guard asserts it.
- **Anything that calls `setSelectedShapes` must run after the ladder.** That
  call wedges the web host's whole selection subsystem, so the ladder
  (`which selection call wedges the host`) has to be the FIRST such call in a
  run — that, and not being alone in a run, is what lets it be routine. Two
  orderings were tried and are wrong, both recorded in `docs/PUBLISHING.md`:
  last (the wedge happens before the ladder asks) and third (the ladder wedges
  six scenarios instead of two). The test states it as the property, because
  adjacency to one named scenario was only ever a proxy for it.
- **`getItemOrNullObject` is not the last word on whether a slide exists.**
  PowerPoint on the web resolved a freshly-added slide's id once and refused it
  ever after, while still listing that id in `slides.load("items/id")`. So
  "the host will not resolve it" ≠ "it is gone": `deleteSlideById` read it that
  way and reported clean-ups it had not done, and a probe run left fourteen
  blank slides in the deck. The deck's own id list is the stronger question —
  use it to identify a slide (`addScratchSlide` diffs it) and to confirm a
  delete. Any new code that acts on a slide id needs the same doubt.
- **A slide handle is good for ONE sync when the slide was just added.**
  Resolving a `getItemOrNullObject(id)` proxy is what makes Office.js rewrite
  its object path to `getItem(id)`, and a freshly-added slide's id does not
  round-trip through `getItem` on the web — so the handle that just passed its
  liveness check answers `GeneralException`
  (`errorLocation: SlideCollection.getItem`) the next time it is used. The
  `SlideThunk` comment says this for `getItemAt` and it is just as true by id:
  re-acquire per sync-batch, never hold one across a `context.sync()`. It cost
  the host probe eight of its fourteen questions (each one recorded as a host
  divergence that had never been asked) and `slideImageBase64` the self-test's
  whole visibility scenario. Pre-existing slides are unaffected — their ids
  round-trip — which is why editing a chart in place has always worked.
- **A chart the drawing context could not tag is not finished.** On the web the
  tag write goes through a shape proxy several syncs old and the host refuses it
  (`InvalidParam passed to GetItem(id)`, 46 times in one 38-item run), leaving a
  chart on the slide with no config — visibly a chart, and not re-editable. The
  demo path always survived it because `insertDemoDeck` re-reads the settled
  deck and plans a `retag`; the ordinary insert and update paths had no such
  pass, which is what `same scale across the deck` was reporting as "3 of 8
  charts carry the shared scale". `settleAndTagChart` gives them the same
  second chance from a fresh context. A settled write is not the "retry against
  a host that just dropped a sync" that gave this project duplicate slides —
  the distinction is a re-read, and it is the one recovery this host honours.
- The showcase build is **byte-deterministic**; CI diffs slide XML, so always
  commit the regenerated deck with the code that changed it.
- The pane rebuilds `ChartConfig` from UI state: new **decoration** keys
  round-trip automatically; new **top-level** config keys need a state field
  or the `state.extras` passthrough in `src/taskpane/app.ts`.
- All sample/showcase data is invented dummy data (`src/core/samples.ts`,
  `scripts/build-showcase.mjs`) — keep it that way; the repo will go public.
- GitHub MCP `actions_list` and `list_pull_requests` responses exceed the token
  cap — parse the saved JSON file with python instead of reading it. In
  `list_pull_requests` the `merged` field is unreliable; read `merged_at`.
- **Squash-merge hides ancestry.** A merged `claude/*` branch is NOT an
  ancestor of `main`, so `git branch --merged` reports nothing and
  `--contains` proves nothing. Confirm a branch is safe to delete from its
  PR's `merged_at`, not from git.
- **Prune after a merge** (`git fetch --prune`, or set `fetch.prune=true`).
  GitHub deletes the branch on merge, but the local `origin/claude/*` tracking
  ref survives pointing at the pre-merge head. The stop hook picks that stale
  ref as its baseline and reports the squash-merge commit as an unverified
  commit "on your branch" — which it is not, and amending it as the hook
  suggests would force-push a rewrite of `main`.
- Object lookups keyed by a config string must use
  `Object.prototype.hasOwnProperty.call` — a pattern/colour/marker named
  `__proto__` or `constructor` otherwise reaches `Object.prototype` and either
  crashes the renderer or gets CALLED. Guarded in `svg.ts`, `pptx-paint.mjs`,
  `geometry.ts`, `i18n.ts`; apply it to any new table. The saved-templates table
  in `app.ts` was a fifth and was missed for months, so **check the write side
  too**: `all[name] = value` where name is `__proto__` hits the inherited
  setter and re-parents the object instead of storing, and the entry then
  vanishes with nothing said. That one is guarded by a null prototype, which
  fixes both directions at the root instead of at each call site.
- `Date.parse` is far looser than a date cell: `parseDateToken` therefore
  gates on shape (date punctuation + month/weekday words only) before parsing.
  Don't route new cell input around it.
- **`dist-lib/` is a build artifact, and `skill-scripts.test.ts` runs against
  it.** A local full-suite run after a change to `src/core` tests the OLD core
  through that path while CI builds fresh — so the suite goes green locally and
  red in CI, for a real regression. Run `npm run build:lib` before trusting a
  local run that touched core.
- **There are THREE colour sinks, and they are separate code on purpose:**
  `src/core/color.ts` (preview), `skill/scripts/pptx-paint.mjs` (headless
  pptx), `officeHex` in `powerpoint.ts` (the live add-in). The same bug has now
  been found in all three independently — each was fixed when a sweep aimed at
  _that_ renderer found it. Change one, check the other two.
- **A `string` in the types is not a string in the file someone pasted.** A
  config arrives from the JSON box, a saved template, a shape tag written in
  another deck, and the skill's caller. `categories: [2023, 2024]` and
  `title: 2024` are ordinary things to write and both used to crash. Coerce at
  the boundary (`normalizeConfig` / `normalizeData` for config,
  `paintText` / `xmlText` for text) rather than at each consumer.

## Out of scope (decided, don't revisit without the owner)

- think-cell's "impossible trio" for a sandboxed add-in: live Excel data
  links, in-canvas drag manipulation, and the slide-layout engine — think-cell
  memory-patches Office binaries for these.
- True geographic maps and 3D surface charts (no freeform paths; the deck this
  project follows argues against 3D anyway). Built instead: tile-grid
  cartograms (`tilemap`) and heatmaps.

## Backlog

`docs/BACKLOG.md` is the single curated backlog (researched candidates with
feasibility/priority, plus a rejected list — don't re-propose those). Items
graduate from there into PRs and are removed when shipped.

Do **not** restate the backlog's contents here — a count in this file and a
list in that one drift the moment either changes, and this paragraph spent
several releases claiming the open candidate was image-output render mode
after that shipped. Read `docs/BACKLOG.md` for what is open; read the README
feature table and git for what has shipped.

## Pending / user-gated

`docs/PUBLISHING.md` is the go-live runbook. Phases 0, 1 and 4 are done: the
repo is public and protected, the add-in is hosted at
<https://powerchart.struktureretsundfornuft.dk/>, and **v0.2.0** ships the
prod manifests, the skill zip and the showcase deck.

What is left needs the owner, not the agent:

- **Phase 2 — keep validating in real PowerPoint.** It HAS run now: three
  answer sheets, several run logs and decks, and the fixes they produced are in
  `git`. `test/fixtures/host-answers-web.json` is one of those recordings and is
  what the CI contract gate diffs against. So the standing question is no longer
  "does any of this work" but "what does the newest build answer".

  What has still never run on a real host: the degradation experiment
  (_what makes a long run slow down_), _edit the chart YOU click_, and every
  probe question added since the fixture was recorded — `PENDING_QUESTIONS` in
  `scripts/host-baseline.mjs` is the authoritative list of those, and it shrinks
  by itself when a newer sheet lands.

  **Owed to the owner right now: a manifest re-install** — `<Version>` changed,
  so the sideloaded copy is stale. What to ask him to click is written down:
  "The standing test run" in `docs/PUBLISHING.md`. Don't improvise a new one per
  session, and don't ask for the deck or a screenshot — the round's own file has
  carried both since the deck-evidence change.

- **Phase 3 — activate the Claude skill** (upload the zip on claude.ai).

Follow it phase by phase; retire items from it and from this list as they
complete.
