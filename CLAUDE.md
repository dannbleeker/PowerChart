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
- **One answer sheet is not evidence about this host — it is evidence about
  this host in that minute.** Two probe questions SWAPPED answers between two
  runs of the same build, minutes apart: `shape-add-held-slide-proxy` went
  `threw`→`yes` and `shape-add-positional-slide-proxy` went `yes`→`threw`. The
  run log says why, and it is not subtle — three `scratch slide landed but its
id will not resolve` lines mid-run, two replacement scratch slides taken, and
  every question inside that window answering `no-scratch-slide` before the host
  came back. The host's ability to resolve a freshly added slide's id comes and
  goes within a single 37-second run, which is the same reversible bimodality
  the draw times show (~10s or ~41s per chart, recovering mid-run and going
  again). So a question that has been asked ONCE has not been answered; it has
  been sampled. Both are declared in `KNOWN_DIVERGENCES` as UNSTABLE with both
  observations, because the dangerous move is building on the convenient one —
  `shape-add-positional-slide-proxy: yes` is exactly what would make a positional
  slide handle look like the safe way out of the by-id refusals.
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
  **The second chance is not a guarantee**, and the deck from 2026-08-06 is the
  counter-example: four charts redrawn to a new scale, ungrouped and carrying no
  config at all, after the settle had had its turn. The fake had no host on
  which the settle also fails — `faults.refuseTagWrites` is now that host, and
  it had to be armed on BOTH tag writers, because the refreshed-proxy path is
  the one the settle uses and a fault on the other only ever refuses the write
  the settle repairs. Which of the two failed on the real host is not yet
  known: `settleUntaggedCharts` traces `{charts, settled, lost}`, so the run log
  says, and the deck alone does not. **Answered on 2026-08-06: neither — the
  settle was never invoked.** A run with FIVE `tagging failed` events carried no
  `settled the config tag` trace at all, and that trace is unconditional, so its
  absence proves the pass never ran. The hole was `updateChartsInSlides`
  returning the caller's old target _bare_ when group-and-tag produced no target
  at all: `lost` stayed undefined, the `no-config` filter matched nothing, and
  the settle returned before it could trace. "Never asked" and "asked and
  failed" look identical from a deck and want different fixes — the trace is
  what separates them, so never reason from a deck about which one happened.
  `targetWithNoTagResult` is that decision, extracted so it can be checked
  without a PowerPoint.
- **Do NOT wait after adding a slide — it was tried and it cost 18 of 19 probe
  answers.** office-js#2903 says a slide added on Online is unusable for a
  couple of seconds and its reporter's fix is to wait; `addScratchSlide` did
  that on 2026-08-07 and the next round answered **1 of 25** questions against
  19 of 26 the build before. The add landed, the wait ran, and the liveness
  check after it found nothing, so every question came back `no-scratch-slide`.
  This host is not the host that issue describes: it resolves a fresh slide's
  id ONCE and refuses it ever after, so waiting spends the one resolution later
  rather than buying time. `web-host.test.ts` guards against reintroducing it.
- **Only an ID may cross a sync — and a proxy's PARENT counts.** A shape proxy
  carries its parent's object path, so members from a re-read collection and
  members from `created` are equally poisoned once Office.js has rewritten that
  slide handle to `slides.getItem(id)`. Handing either to `addGroup` does not
  merely fail to group: it THROWS, and the throw takes the batch's tagging with
  it, so the chart loses its group AND its config. Five charts in one run.
  `chooseGroupMembers` is that decision — group by re-resolved ids, or group
  nothing — and "group nothing" is strictly better, because an ungrouped chart
  that carries its config is still re-editable. **It worked on 2026-08-07**
  (`not grouping: no member handle this host will accept index=0 refreshed=0`)
  and the chart lost its config anyway — for the reason in the next bullet, not
  this one.
- **The host may refuse to name a SHAPE by id, and the recovery is a collection
  read.** Sixty-six errors in the 2026-08-07 run log, every one `InvalidParam
passed to GetItem(id)`, code 5010, at `errorLocation: ShapeCollection.getItem`
  — the slide answered, the shape did not. It took the drawing context's tag
  write, the readback, `ungroupedFallback`'s id read, **and
  `settleAndTagChart`'s own fresh-context write**, which is exactly what "the
  update reported 5×no-config" was. The settle then gave up, on the reasoning
  that a collection search "would only find a DIFFERENT shape to put this
  chart's config on" — sound with no id, wrong with one, because the read loads
  `items/id` and the caller's id picks its own shape out of the answer.
  `settleByCollectionRead` is that fall-through, and a collection read is the
  pattern this host DOES honour: 23 retags landed that way in the run that lost
  46 tag writes. `faults.refuseShapeById` is the fake being this host — and it
  must be paired with a refused in-context tag write, or the settle never runs
  and the guard passes against the unfixed file, which is how the first version
  of that test proved nothing.
- **A printed `getItem` does NOT mean a held handle — ASKED AND ANSWERED.**
  Office.js rewrites a resolved `getItemOrNullObject` proxy's path to
  `getItem(id)`, so for months every log read as though this code were holding
  handles across syncs. It was not. With `extendedErrorLogging` on, the host
  annotates the call each path was CREATED by, and the answer is unambiguous:
  `var slide = slides.getItem("282#543504795") /* originally
getItemOrNullObject("282#543504795") */`. The rewrite is just how a resolved
  null-object proxy prints. Do not re-derive this from an excerpt; the
  annotation is the evidence, and it took one build to get after two sessions
  of reasoning that got nowhere. The same annotation is what identifies a
  genuinely poisoned proxy: `shapes.getItem("27") /* originally addTextBox(...)
*/` is a `created` proxy, and those the host really does refuse.
- **When the excerpt cannot settle it, make the host say more.** Every
  `debugInfo` before 2026-08-07 ends `"fullStatements":["Please enable
config.extendedErrorLogging to see full statements."]`, so all a reader got was
  `surroundingStatements`, an excerpt — and the excerpt could not tell a held
  handle from a rewritten one, which is the question above. Turning the flag on
  answered it in a single round. `enableExtendedErrorLogging` is the
  two-explanations rule applied to the tooling: stop reasoning, make the host
  say. `trimDebugInfo` keeps BOTH ends of the statement list so the file stays
  sendable — tail-only was the first version and it was wrong, because
  `surroundingStatements` centres its `>>>>>` marker on the failing statement
  and in that round the marker sat on the batch's FIRST line while the log said
  "… 37 earlier statement(s) dropped".
- **`load("items")` does not load the items' properties — Microsoft says so.**
  "You must explicitly specify each property you need from collection items, as
  they won't be loaded by default, including scalar properties." Every
  collection load in this file names its properties; the re-read before grouping
  was the one that did not, and `id` is the only thing it reads. Changed to
  `items/id` on 2026-08-07 — after being changed and REVERTED once the same day,
  which is the part worth keeping: the first attempt had no source behind it and
  the suite went red, so it was traded back rather than argued for. The red test
  turned out to be asserting the wrong property (which writer landed the tag,
  rather than whether the chart is re-editable), and it only surfaced because
  the fake's `hollowReads` keys on the projection string. Guarded by a source
  scan in `web-host.test.ts`.
- **The web host does not LIST the shapes a run just added.** The finding the
  extended log produced, and the one everything else downstream hangs off:
  `the re-read before grouping came back empty index=0 drew=24`, four times in
  one round, on slides that had just taken 24 shapes each. The probe says the
  same thing three more ways in the same run — `shapes-items-count-honest:
unreadable`, `shape-proxy-survives-one-sync: unreadable`,
  `shapes-items-via-positional-slide: not-listed`. So `refreshed=0` is the EMPTY
  case, not a throw and not a failed match, which is exactly why that trace was
  split three ways. The consequence chain is the whole of `same scale across the
deck`: nothing to group with, no fresh tag target, so the tag goes through a
  `created` proxy, the host refuses it, and the settle has nothing better to
  offer. Slides are fine — `getitem-durable-slide: yes`,
  `shape-add-positional-slide-proxy: yes`. It is the shape COLLECTION that will
  not answer.
- **A self-test scenario that ends the run costs the whole report, even last.**
  `the chart is actually visible` ran dead last precisely so its crash could not
  take other scenarios' verdicts — and it still cost every round, because the
  report is written when the battery RETURNS. Four rounds, four builds, always
  within a step or two of `adding a scratch slide`, never once a verdict. It is
  `pickedOnly` now. "Last" protects the verdicts; only "out of the routine list"
  protects the report. Whether the scenario or the ten minutes in front of it is
  the killer has never been separated — running it alone is that experiment.
- **A rasterise answers fast or not at all — never wait a readback's budget for
  one.** `getImageAsBase64` on a freshly-added slide has now failed on the web
  three different ways in three rounds: `GeneralException` at
  `SlideCollection.getItem`, then taking the call and silently producing
  nothing, then never answering the sync. The third cost a whole round —
  `the chart is actually visible` sat on the full ninety-second readback budget
  and the tab died on the delete that followed, taking the run's report with it,
  for a scenario whose honest verdict is `skipped`. `rasteriseTimeoutMs` is
  twenty seconds, capped by the readback budget so a test can still shorten it.
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
  (_what makes a long run slow down_), _edit the chart YOU click_, _the chart is
  actually visible_ — which is `pickedOnly` as of 2026-08-07 precisely because it
  killed the tab four rounds running without ever returning a verdict — and every
  probe question added since the fixture was recorded. `PENDING_QUESTIONS` in
  `scripts/host-baseline.mjs` is the authoritative list of those, and it shrinks
  by itself when a newer sheet lands.

  **Nothing is owed to the owner right now.** The manifest re-install he was
  asked for was done on 2026-08-06, and nothing since has touched a manifest —
  do not ask again unless a PR actually changes one. What to ask him to click is
  written down: "The standing test run" in `docs/PUBLISHING.md`. Don't improvise
  a new one per session, and don't ask for the deck or a screenshot — the round's
  own file has carried both since the deck-evidence change.

  **A round is in flight.** `48e9a00` was merged and deployed 2026-08-07 17:49Z,
  and the owner has been asked for two things: a normal full round, and — on a
  fresh deck — a second short round with _the chart is actually visible_ picked
  alone. That second one is an experiment, not a regression check: every crash so
  far arrived ten minutes and nine scenarios into a run, so "the scenario kills
  the host" and "ten minutes of drawing kills the host, and this is what happened
  to be running" both fit, and running it alone is what separates them. Whatever
  comes back — including another crash — is the answer.

  Three things in `48e9a00` have never met a real host: the settle's fall-through
  to a collection read, `enableExtendedErrorLogging`, and the routine round
  finishing at all. **The extended statements are the one to read first** — every
  log before this build says `"Please enable config.extendedErrorLogging"`, and
  the question it was turned on to settle is whether a batch printing
  `slides.getItem(...)` is a held handle or just how Office.js prints a fresh
  `getItemOrNullObject`. Those two readings disagree about whether the settle is
  repairable, and no amount of reasoning has separated them.

- **Phase 3 — activate the Claude skill** (upload the zip on claude.ai).

Follow it phase by phase; retire items from it and from this list as they
complete.
