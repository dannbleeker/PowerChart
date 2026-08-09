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

## Where things live

`test/README.md` maps the test suite; this is the source side.

| directory                       | what it owns                                                                                                                                                                                                                                                                                                                           |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/core/`                     | the engine, **zero Office imports**: `chart.ts` (`buildChart`), `layout/<kind>.ts` one per chart kind, `decor.ts`, `scene.ts` (the node contract **and** the three-renderer parity rules), `geometry.ts`, `format.ts`, `color.ts`, `collide.ts`, `samples.ts`                                                                          |
| `src/core/`, run-time reasoning | `placement.ts` (where a chart goes on a slide that already has content), `reconcile.ts` (what a run ACTUALLY produced, once the host stopped moving), `trace.ts` (the ordered record a run nobody watched leaves behind)                                                                                                               |
| `src/render/`                   | `svg.ts` (the reference renderer), `powerpoint.ts` (Office.js, the live add-in), `ooxml.ts` (post-processes a `.pptx` to carry the slot tags and groups pptxgenjs cannot write), `pptx-deck.ts` (builds a whole deck in the browser, handed over in one call), `host-probe.ts` (the answer sheet — what THIS PowerPoint actually does) |
| `src/taskpane/`                 | `app.ts`, `datasheet.ts`, `selftest.ts` (the in-host battery), `crashlog.ts` (the record that survives a run which never ends), `i18n.ts`, `templates.ts`                                                                                                                                                                              |
| `skill/scripts/`                | `render-pptx.mjs` + `pptx-paint.mjs`, the headless renderer. **Outside `tsconfig.include` (`["src", "test"]`), so never typechecked**                                                                                                                                                                                                  |
| `scripts/`                      | `triage.mjs`, `verify-deck.mjs`, `validate-ooxml.mjs`, `host-baseline.mjs`, `visible-charts.mjs`, `office-js-watch.mjs`, and the `build-*` set                                                                                                                                                                                         |

**Adding a `SceneNode` kind: four seams fail loudly, three do not.** Loud, as
compile errors — `nodeToSvg` (`src/render/svg.ts`), `addNode`
(`src/render/powerpoint.ts`), the coordinate chain in `test/fuzz.test.ts`, and
`translateNodes` in `src/core/chart.ts`, which carries an explicit `never`
guard for exactly this. Silent:

- `shiftNode` in `src/demo/demo.ts` — a duck-typed `Record<string, number>`
  cast over a fixed key list, so it can never fail to compile; a new coordinate
  field is simply left unshifted.
- `src/core/collide.ts` — text-only, matched by `MOVABLE` name prefixes, so a
  new label-like node is invisible to the de-collision pass.
- `skill/scripts/render-pptx.mjs` — not typechecked at all (see the table).

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
npm test           # full suite; npm run coverage enforces thresholds
npm run typecheck
npm run showcase   # regenerate the showcase deck (required after feature work)
npm run skill      # build skill-dist/powerchart-charts.zip
npm run triage     # join a real run's deck + run log, and say where they differ
npm run host-diff  # line a probe answer sheet up against the fake
```

The suite's size lives in `test/fixtures/test-count.json` and is gated by
`scripts/test-count.mjs` — read it there. This line used to restate the number,
and drifted to claiming 1500 against a recorded floor of 2224: the same failure
the backlog and `UNSTABLE_ANSWERS` paragraphs below already warn about.

### Running the gate on the owner's Windows box

Environment only — CI on ubuntu is unaffected — but locally the commands above
fail in ways that read as code bugs:

- **`npx` is dead under PowerShell Constrained Language Mode.** `npx.ps1` does
  method invocation and throws
  `MethodInvocationNotSupportedInConstrainedLanguage` before reaching the tool.
  Call the entry point with `node`, or run it from git-bash.
- **AppLocker blocks an npm script that NESTS `npm run`** ("blocked by group
  policy"). That kills `showcase`, `skill`, `build:pages` and `visible-charts`
  — every one starts with `npm run build:lib`. Split them into their two flat
  halves. **Any new npm script must stay flat** or it is unrunnable there.

```bash
node ./node_modules/typescript/bin/tsc --noEmit                 # typecheck
node ./node_modules/vitest/vitest.mjs run                       # whole suite, nothing excluded
npm run build:lib && node scripts/build-showcase.mjs            # = npm run showcase
npm run build:lib && node scripts/build-skill.mjs               # = npm run skill
```

**The whole suite runs there now — do not exclude anything.** This paragraph
used to say `test/skill.test.ts` could not pass and to run with
`--exclude '**/skill.test.ts'`, which was true and cost more than it looked: the
file that checks what the SKILL ships was gated by CI alone. Both reasons are
gone. It read a `.pptx` by interpolating an OS temp path into a `python3 -c`
string, where `\rings.pptx` is a carriage return, and `build-skill.mjs` zipped
through `python3 -m zipfile`, which on Windows is the Microsoft Store alias stub
and simply fails. Both now use `jszip`, already a dependency — so the zipper and
the reader are one library, no interpreter is involved, and the zip is
byte-identical run to run.

`node scripts/build-showcase.mjs` **alone renders with a stale engine** — it
imports `dist-lib/`, so without `build:lib` first you are diffing the showcase
against your own uncompiled change, and it reads as non-determinism.

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
  this host in that minute.** Several probe questions give different answers on
  different runs of the SAME build, minutes apart —
  `shape-add-held-slide-proxy` and `shape-add-positional-slide-proxy` most
  sharply, and for a while in lockstep and always opposite. That is not a
  sequence of corrections where the newest value is the true one; it is a coin,
  and a majority across samples is not a mechanism either. A question that has
  been asked ONCE has not been answered, it has been sampled.

  The run log says why, and it is not subtle: `scratch slide landed but its id
will not resolve` several times mid-run, replacement scratch slides taken, and
  every question inside that window answering `no-scratch-slide` before the host
  came back. The host's ability to resolve a freshly added slide's id comes and
  goes within a single run.

  **`UNSTABLE_ANSWERS` in `scripts/host-baseline.mjs` is the authoritative list,
  with every observation on every entry — do not restate its counts here.** A
  tally in this file and a list in that one drift the moment either changes, and
  this paragraph spent several rounds claiming "five sheets" while the table had
  six and eight; the same mistake the backlog paragraph further down already
  warns about. Read the table.

  What the table is FOR is stopping the next reader building on whichever answer
  a sheet happens to carry. `shape-add-positional-slide-proxy: yes` is exactly
  what would make a positional slide handle look like the safe way out of the
  by-id refusals, and it is what most of its samples say. It would not help
  anyway: `shapes-items-via-positional-slide` answers as short as the by-id form
  does in every run that put both, so a positional handle reads a shape
  collection no better.

- **The fake is gated against a real host in CI** — `test/host-contract.test.ts`
  diffs `FAKE_BASELINE` against the committed sheet in
  `test/fixtures/host-answers-web.json`. A new divergence fails there unless it
  is declared in `KNOWN_DIVERGENCES` with a reason. When a fresh sheet arrives:
  replace the fixture, run the suite, deal with what goes red. Do not edit the
  fixture by hand — it is a recording, not a preference.
  **A newer sheet is not automatically a better one.** The `a546897` round put
  only 18 of 27 questions, and swapping it in would have turned eight committed
  ANSWERS — `delete-then-lookup: reports-gone`, `group-reports-its-children:
threw`, `tags-on-fresh-shape: yes`, `untrack-available: no` among them — into
  `no-scratch-slide`, which the gate correctly reads as "unknown". That is
  deleting knowledge on the strength of a bad ten minutes, the same mistake the
  gate itself used to make one level up. Replace when the new sheet answers at
  least as much; otherwise keep the fixture, and put what genuinely moved into
  `UNSTABLE_ANSWERS` where it belongs.
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
  **Confirmed head-on 2026-08-08, and it is not about the handle at all.** A
  FRESH `slides.getItem(id)` — nothing held, nothing resolved earlier — answered
  `threw` (`GeneralException`) for a slide added moments before, while its
  follow-up partner asked the same call about a pre-existing slide in the same
  run and got `yes`. So the rule is not "do not hold a slide handle"; it is
  **`getItem(id)` does not work on a NEW slide, by any route**. That is the call
  `getTargetSlide` makes. The partner is `Probe.follow` doing its job: one run,
  two readings separated, no reasoning.
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
  without a PowerPoint. **And on 2026-08-08 the next question was answered too:
  the settle now RUNS and still loses.** Five `settle pass:` lines in one round,
  every one of them `settled: 0, lost: 1`, each preceded by `the settle's
re-read came back empty`. So the fall-through to a collection read does not
  rescue a chart the run has just drawn — it meets the same empty collection
  everything else does. That is not in tension with the 23 retags an earlier run
  landed by collection read: those were pre-existing slides read by the repair
  pass, and this is a slide the run drew seconds ago. The distinction to hold on
  to is not by-id versus collection, it is **fresh versus settled**.

  **Partly overturned on 2026-08-09 by round 8: the settle CAN rescue a
  freshly-drawn chart, and did.** One of five settles in that round reported
  `settled: 1, lost: 0` — the first success on record. What separates it from
  the four that failed is not freshness: it is whether the chart was GROUPED.
  A grouped chart gives the settle a shape id to write through; an ungrouped one
  leaves it a collection read, and the collection is what this host will not
  answer. So the rule is narrower than "the settle does not rescue fresh
  charts": it rescues the ones grouping survived for, which on a degrading host
  is the chart or two either side of the flip. The
  messages carry their outcome now — they all begin `settle pass:` so absence
  still reads as "never invoked" — because for one round the log said
  "settled the config tag…" five times while its own numbers said nothing was
  settled.

- **Do NOT wait after adding a slide — it was tried and it cost 18 of 19 probe
  answers.** office-js#2903 says a slide added on Online is unusable for a
  couple of seconds and its reporter's fix is to wait; `addScratchSlide` did
  that on 2026-08-07 and the next round answered **1 of 25** questions against
  19 of 26 the build before. The add landed, the wait ran, and the liveness
  check after it found nothing, so every question came back `no-scratch-slide`.
  This host is not the host that issue describes: it resolves a fresh slide's
  id ONCE and refuses it ever after, so waiting spends the one resolution later
  rather than buying time. `web-host.test.ts` guards against reintroducing it.
- **Two `slides.add()` calls in quick succession kill PowerPoint on the web.**
  Named on 2026-08-08 by the step tracing added the same day, in one round:
  `adding the first scratch slide` at 33.2s, `adding the second scratch slide`
  at 33.6s, and the tab was gone. The first add SURVIVED — the second step is
  only written if it did — and nothing after the second was reached, which
  exonerates the ninety-six shapes the scenario was about to draw. Four tenths
  of a second apart is all it took.
  Read with the rasterise gotcha below, the pattern is one thing: whatever this
  host does after `slides.add()` it does slowly and badly, and a second
  operation arriving before it settles takes the process with it. Rasterising a
  fresh slide killed five rounds; adding a second slide killed two more.
  **Nothing in this repo takes a scratch slide any more.** Both scenarios that
  did have stopped, and `test/selftest.test.ts` asserts it for the whole battery
  by slide count, so a new one cannot quietly reintroduce it. The host probe
  still takes them — it has no alternative, every question is about a slide it
  owns — which is why it churns through so many and why its rounds are the ones
  that see `no-scratch-slide` windows.
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
- **Drawing cost grows with the shapes ALREADY on the slide — and the request
  context has nothing to do with it.** `what makes a long run slow down` finally
  measured something on 2026-08-09, after two outings that died in setup. Eight
  rounds of twelve rectangles per arm, one arm holding a single context open and
  the other taking a fresh one each round:

      one context     2339 3177 3704 4363 5010 5775 6372 6769 ms   +101%
      fresh contexts  2852 3397 3936 4636 5165 5635 6736 7352 ms   +105%

  Both arms grew, within four points of each other, so the CONTEXT is not the
  variable — which is what the experiment was built to decide, and it is now
  decided. What each round shares is that it adds twelve shapes to a slide that
  already holds twelve times the round number: the per-round increments are
  roughly constant (+838, +527, +659, +647, +765, +597, +397), so a round costs
  about `2339 + 630(n-1)` ms and the TOTAL is quadratic in shapes per slide.
  Ninety-six shapes on one slide cost 37.5 seconds; the same ninety-six over
  eight slides would cost about nineteen. That is the arithmetic behind the
  17-second batches, and behind dense charts being the ones that fail.
  **REPRODUCED on a second round (`87dc418`), and the second round corrects the
  first's reading of it.** The curves land on top of each other:

      one context     2395 3095 3713 4400 5072 5734 6398 7455 ms   +107%
      fresh contexts  2428 3056 3676 4356 5557 5619 6345 7264 ms   +106%

  Round for round within a few percent of the run above, on a different build,
  with the same `deckBefore=7`. The per-round increment is ~660ms both times.
  That part is solid.

  **And it is PER-SLIDE, not per-deck — the two arms already separated that.**
  This file said a third arm on a fresh slide per round would be needed and that
  the host would not give one out. Wrong: the arms run in SEQUENCE on DIFFERENT
  slides, so the second arm's first round is the measurement. It starts after
  the first arm has put ninety-six shapes on another slide in the same deck, and
  if the cost were deck-wide it would start slow. It does not: 2395 against 2428,
  a 1.4% gap. The 22% gap in the first round (2339 against 2852) that this file
  called "the only hint" of a deck effect did not reproduce, and one sample of a
  22% gap next to one of 1.4% is noise, not a mechanism.

  So: adding shapes to a slide makes THAT slide expensive to draw on, and costs
  the next slide nothing. Which is why spreading a deck's charts one per slide
  is not merely tidier — it is the difference between linear and quadratic.

  **The routine rounds add a second hint, and it points the other way.** Two
  consecutive rounds stalled in the same two scenarios, on the same call, at
  much the same times — `drawing shapes 1-10 of 24` giving up after 45 seconds
  at 481s and 591s on `1fd6aa3`, and at 456s and 567s on `40b5e44`, the latter
  on a deck that started with ONE slide (`round starting deckSlides: 1`). Those
  draws go onto slides holding a couple of charts, not ninety-six shapes, so
  per-slide cost does not explain them.

  And per-deck SHAPE COUNT does not either, now that the second experiment round
  has ruled it out — ninety-six shapes on a neighbouring slide cost the next
  slide 1.4%. What is left is the TAB: elapsed time, memory, or whatever else a
  long-lived Office.js session accumulates that a shape count cannot see. That
  is a narrower answer than "the deck or the tab", and it is narrower because a
  second sample was taken rather than argued for.

  **But elapsed time is not the whole of it either, and the 2026-08-09 sweep of
  seven rounds says why.** Read as a distribution rather than as anecdotes, the
  stalls have two properties nothing here had noticed:

  - **Every stall is the FIRST batch of a scenario's draw. Eight for eight,
    `drawing shapes 1-10 of 24`, never batch 2 or 3.** If the host were simply
    fatigued, stalls would land anywhere in a chart's three batches. They land
    on the first sync after some other operation and nowhere else.
  - **Elapsed time cannot be the variable on its own**, because in every one of
    those rounds the scenario immediately AFTER the stalled one — older tab,
    more elapsed time — drew its full twenty-four shapes and passed. r6 is the
    clean case: `a selected shape survives an insert` stalls at 410s, `edit the
chart the user selected` passes at 459-514s, `the chart is actually visible`
    passes at 517-586s.

  What every stalled scenario had in common was its PREDECESSOR: the selection
  ladder (×4), `stop a run part-way`, which aborts a draw mid-flight (×3), and
  a scenario that itself selected and stalled (×1). That was 8 of 8.

  **RETRACTED on 2026-08-09 by round 7, which is why it was written down as a
  correlation.** `a selected shape survives an insert` PASSED — in the routine
  round, at 407s, eighth in the list and immediately after the ladder, which is
  the exact position and the exact predecessor that had stalled it four times
  out of four. Same build family, same battery, same order. A clean
  counter-example at the identical position kills the predecessor account
  outright; nothing about "what ran before it" survives.

  The picked round ran the same day and agrees. Alone on the deck, with only the
  two head-of-round inserts in front of it, it drew all three batches (8.2s,
  12.2s, 12.4s) and passed — at 750s of tab life, thirteen minutes in, which
  takes elapsed time down with it for that scenario.

  So of the four candidates, three are now out for `a selected shape survives an
insert`: the scenario, its predecessor, and the tab's age. What is left is
  that this host stalls INTERMITTENTLY, and the run of four was a run.

  **Both things that were built to settle it have now reported, and both did.**

  - **The call-level record fired on its first outing.** The log used to have
    nothing at all between a scenario announcing itself and its first `batch
committed` — three to five seconds of probe reads, deck inventories and
    selection calls, all invisible — which is the only reason the account was
    ever at scenario level. Round 7's one draw stall says:

        gave up waiting  what="drawing shapes 1-10 of 24"
                         afterAnswering="rasterising a slide"  idleMs=1

    A rasterise answered, the draw's first sync went out a millisecond later,
    and it never came back. Every future round carries this for free.

    **`idleMs` IS NOT THE VARIABLE — asked and answered on the next round.**
    The gap was recorded for the first batch of every draw, stalled or not, and
    round 8 (`d812d0c`) reported four batches that all SURVIVED:

        82.6s  idleMs=2182   edit a chart on the visible slide

    100.1s idleMs=2 insert onto a slide that already has content
    129.9s idleMs=1 insert onto a slide that already has content
    188.3s idleMs=854 same scale across the deck

    A surviving first batch went out **one millisecond** after the previous
    answer, which is exactly what the stall reported. A value that occurs in
    both populations cannot separate them, so the gap is dead as a lead and
    nothing should be built on it — least of all a wait before drawing, which
    this host has already punished once (see the scratch-slide gotcha).

    The instrumentation stays. It cost one trace field, it killed a plausible
    hypothesis in a single round, and the half that is still live rides along
    with it: `afterAnswering` names the CALL, and the identity of the call is
    the lead the gap turned out not to be.

  - **The picked round ran and answered.** The criterion was set in advance:
    stalls → the scenario itself; passes → predecessor or position. It PASSED,
    so the scenario is out, and the routine round passing it at the same
    position takes the predecessor with it.

  **Round 9 removed the last of that too.** It produced TWO stalls, and they
  named two DIFFERENT predecessors:

      496.6s  drawing shapes 1-10 of 24   afterAnswering="selecting a shape"     idleMs=1
      608.2s  drawing shapes 1-10 of 24   afterAnswering="rasterising a slide"   idleMs=1

  So "the sync after a rasterise" is not the pattern — `a selected shape survives
an insert` stalled after a selection call, having passed the two rounds before.
  And `edit the chart the user selected` ALSO selects a shape and then draws, in
  the same round, and survived.

  **And the predecessor's name is in exactly the condition the gap was in: no
  baseline.** It is recorded on stalls and nowhere else, so thirteen draws
  survived in that round without saying what they followed. That is the same
  mistake `idleMs` cost two rounds to kill, made a second time with a different
  field — which is why the first batch of every draw now records BOTH. One round
  decides whether the identity of the preceding call discriminates or goes the
  way the gap did.

  **Round 10 gave that field its baseline, and unlike the gap it SURVIVED.**
  With `afterAnswering` on every first batch rather than only on stalls, the two
  populations separate cleanly:

      survivors followed   moving the view to a slide, counting the deck's slides,
                           writing the chart's origin tag (x5), re-reading a slide
                           to tag the chart it would not tag (x4), selecting a
                           shape, reading the selected chart
      the stall followed   rasterising a slide

  Twenty-nine surviving first batches across rounds 9 and 10, not one of them
  after a rasterise; two stalls after a rasterise, in consecutive rounds. And
  `selecting a shape` — round 9's other stall — turns up among round 10's
  survivors, which is what a non-cause looks like.

  **And round 11 killed it, the same way every other candidate died.** That
  round contains two draws after a rasterise, 150 seconds apart:

      584.5s  pass   after="rasterising a slide"    (the chart is actually visible — PASSED)
      734.8s  STALL  after="rasterising a slide"    (the control's rasterise arm)

  A value that occurs in both populations cannot separate them. So
  `afterAnswering` has gone the way of `idleMs`: **no call before the draw
  predicts the stall.** Every candidate is now eliminated — the scenario, the
  preceding scenario, the tab's age, the idle gap, and the identity of the
  preceding call. What is left is a host that stalls the first sync of a draw
  intermittently, at roughly one or two draws in fifteen, and nothing yet
  distinguishes which.

  **The control's own verdict that round was WRONG, and that is the more useful
  lesson.** It reported `the draw after a RASTERISE did not land` — a claim the
  same log contradicts — because the two arms ran in a fixed order, cheap first
  and rasterise second, so the rasterise arm was always the later one. A
  diagnostic that manufactures a finding is worse than no diagnostic, and this
  one would have manufactured that finding every round it stalled.

  It is counterbalanced now: rasterise, cheap, cheap, rasterise, with each call
  type running once early and once late. `rasteriseArmVerdict` names the CALL
  only when every rasterise arm fails and every cheap arm draws; it names
  POSITION when both late arms fail and both early ones draw; and it says "no
  pattern" for anything else, which on eleven rounds of evidence is the answer
  to expect. The arms draw eight-shape charts rather than the battery's
  twenty-four, so four of them cost about what two of the old ones did.

  This is the same shape as the two experiments that already paid: `the chart is
actually visible` and `what makes a long run slow down` each spent four rounds
  saying "it crashed again" and one picked round saying which call.

- **A value recorded only on FAILURES cannot be compared against anything, and
  this project keeps building them.** Four in one session: `idleMs` and
  `afterAnswering` were written on stalls but not on the draws that survived;
  the settle's two writes shared one label so a refusal named neither; and
  `listChartsInDeck` traced only when a scan came back SHORT, which made every
  deck scan in every round invisible. That last one has a price tag: round 10's
  `stop a run part-way` took 39.4 seconds against 2.6-3.2s in the eight rounds
  before it, and the log had a 39-second hole where the scan was. The stop
  itself was instant — the verification after it was not, and nothing said so.
  The scan is also the operation the quadratic per-slide cost predicts should
  grow worst, since it reads every slide's shapes on a deck the battery keeps
  adding to, and it was the one operation never measured. Before adding a
  diagnostic field, ask what its value is on the runs that WORK; if the answer is
  "it is not written", the field cannot discriminate and is not yet a
  measurement.

- **A stall is DEATH, not slowness — do not raise `BATCH_TIMEOUT_MS` hoping for
  an answer.** Seventeen abandoned calls across nine rounds, and not one of them
  ever came back: `a call we gave up on finally answered` appears zero times in
  any round file, including the 100-200 seconds each round keeps running
  afterwards. Round 9 looked like a counter-example and is not — two probe
  questions there answered at 8429ms and 8399ms against an 8000ms budget, which
  reads as a late answer and is the probe's RETRY: it replaced the scratch slide
  and re-asked, and the elapsed time is measured from the first attempt. And the population is bimodal with an empty band — of 327 batches
  that DID answer the slowest took **29.2s**, against a 45-second budget, so
  nothing has ever landed between 29s and 45s. A batch answers within ~29s or
  never.

  Two things follow. Raising the budget buys nothing, and lowering it to ~35s
  would save ten seconds a round for a 1.2× margin instead of 1.5× — not worth
  it either. Leave it alone; the number is fine and the finding is what it means.

  **That was read out of a trace line's ABSENCE, which is this project's most
  expensive habit** (`settleUntaggedCharts` was "ran and failed" for two
  sessions when it had never run). So it is measured now rather than inferred:
  a stalled scenario waits `LATE_ANSWER_WAIT_MS` and its verdict says which
  happened, in words, every time (`stallDetail`). The plumbing behind that had
  never been tested on the DRAW path either — every existing late-answer test
  goes through `insertDemoDeck` — so there is now one that stalls a real draw
  batch and checks the report carries the late answer.

- **FAST IS THE BROKEN MODE, not the healthy one.** The draw times are bimodal —
  ~17s per batch or ~3-5s — and this file said until 2026-08-08 that the host
  "recovers mid-run and goes again". It does not. Within ONE run of `same scale
across the deck`, eight updates of the same chart at the same size: charts 1-4
  took ~17s per batch with no grouping failures at all, and charts 5-8 took
  ~3-5s and every one of them carries `the re-read before grouping came back
empty` and `not grouping`. The boundary is exact and the correlation is
  8-for-8. Batches got cheap because the host stopped answering for the shape
  collection, so there was nothing left to group or read back — it was not
  recovering, it was degrading further. Do not read a fast stretch as health,
  and do not "fix" the slow one: `what makes a long run slow down` was written
  on the assumption that slow means degraded, and it has that backwards.
  (AutoSave was the obvious candidate for a periodic slowdown and is ruled out:
  the regimes are long contiguous blocks — eight slow, then ten fast — not
  anything alternating on a timer.)

  **Reproduced exactly on 2026-08-09, and it makes the scenario's SCORE
  readable.** In the `40b5e44` round, `same scale across the deck` redrew its
  eight charts between 134.7s and 342.4s. Charts 1-3 ran at 16-18s per batch;
  chart 4 straddled the boundary; charts 5-8 ran at 0.6-5s. The five grouping
  failures in that window land at 291.8, 303.8, 316.5, 328.9 and 341.1s — one
  per fast chart, none for a slow one, each the full chain (`the re-read before
grouping came back empty` → `not grouping` → `tagging failed` → `settle pass:
could not repair any`). The verdict was **3 of 8**, which is the three slow
  charts.

  So this scenario's score is not a variable defect. It is a measurement of WHEN
  the host flips regime: 3 of 8 and 4 of 8 across rounds is the flip landing one
  chart earlier or later. Read it that way, and stop treating a move from 4 to 3
  as a regression.

  **Round 8 (`d812d0c`) decomposes the score exactly, and the degradation turns
  out to have THREE stages rather than two.** Per-chart wall clock:

      charts 1-3   ~35s each   SLOW    clean: grouped, tagged, done
      chart  4       8.1s      fast    grouped; tag refused; SETTLE REPAIRED IT
      chart  5       8.3s      fast    grouped; tag refused; settle refused too
      charts 6-8   ~8.2s each  fast    NOT grouped; tag refused; settle empty

  4 of 8 = the three slow charts plus the one the settle rescued.

  **Round 9 (`448ffc6`) narrows that: the three stages were one round's shape,
  not a structure.** Same scenario, same build family, and the two boundaries
  land together — charts 1-3 slow and grouped, charts 4-8 fast and NOT grouped,
  five settles all `settled: 0, lost: 1`, score 3 of 8:

      charts 1-3   ~35s each   SLOW    grouped, clean
      charts 4-8   ~8.4s each  fast    NOT grouped, all lost

  So the speed flip and the grouping death are INDEPENDENT and may or may not
  coincide: a chart apart in round 8, together in round 9. What survives both is
  narrower and more useful than either — **the score counts the charts that
  GROUPED and kept a tag, and grouping is what the collection read decides.**
  Speed is a symptom sitting near that boundary, not the thing that sets it.

  **And the settle repaired a chart for the first time on record** —
  `settle pass: repaired every config tag the drawing context lost, charts=1
settled=1 lost=0`. Every previous observation was `settled: 0, lost: 1`, and
  this file said so. The recovery works; it needs the chart to have been grouped,
  because that is what gives it an id to write through.

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
  not answer. **Settled 2026-08-08, twice over.** The parent handle was never
  the problem: `shapes-items-count-honest` and its positional partner
  `shapes-items-via-positional-slide` both answered `short-0` in the same run,
  so renaming a readback positionally buys nothing. And the collection is not
  merely quiet, it is inconsistent with itself —
  `getcount-populates-same-sync: yes, value=8` next to `items=0`, same host,
  same minute. The count is right and the list is empty. That is why every
  collection read in this repo has to be corroborated against the slide's own
  count (`slideShapeNames`) rather than believed.
- **A self-test scenario that ends the run costs the whole report, even last.**
  `the chart is actually visible` ran dead last precisely so its crash could not
  take other scenarios' verdicts — and it still cost every round, because the
  report is written when the battery RETURNS. Four rounds, four builds, always
  within a step or two of `adding a scratch slide`, never once a verdict. It is
  `pickedOnly` now. "Last" protects the verdicts; only "out of the routine list"
  protects the report. Whether the scenario or the ten minutes in front of it is
  the killer has never been separated — running it alone is that experiment.
- **Do not rasterise a slide the run just added. It kills PowerPoint.**
  `getImageAsBase64` on a freshly-added slide has now failed on the web FIVE
  distinct ways: `GeneralException` at `SlideCollection.getItem`; taking the
  call and silently producing nothing; never answering the sync; sitting on the
  full ninety-second readback budget; and — 2026-08-08 — killing the tab
  outright. `rasteriseTimeoutMs` is twenty seconds, capped by the readback
  budget so a test can still shorten it, and it does not help against the fifth:
  a timeout cannot save a process that is gone.

  **The isolating experiment ran on 2026-08-08 and it is the scenario, not the
  load.** `the chart is actually visible`, picked alone on a fresh deck, was
  reached at 61.5s with only its two inserts in front of it — against ten
  minutes and nine scenarios in the four rounds before. It added a scratch slide
  at 61.5s, logged `rasterising the empty slide` at 61.8s, and the tab died.
  Those same two inserts head every routine round and kill nothing, so elapsed
  time and volume of drawing are both out. Four rounds of "it crashed again"
  said nothing; one picked round said this.

  `chartIsVisible` no longer takes a scratch slide at all — it does its
  before/after on a slide the run added EARLIER, which drops the scratch add,
  the fresh-slide rasterise and the delete (that delete had killed a round of
  its own).

  **And the re-run closed the last reading.** On `e49cca8` the step named
  `rasterising a slide that already existed` answered without incident, the
  scenario drew its 24 shapes, rasterised again and returned a verdict — the
  first in six rounds. So this host rasterises perfectly well; it was the FRESH
  slide, both directions now measured rather than inferred. A slide the run just
  added remains the worst surface this host offers, and nothing may rasterise
  one.

  **It PASSED on `c7d91d5` and is routine again** — `drawing the chart changed
what the slide looks like (10064 → 15652 bytes)`, through PowerPoint's own
  rasteriser. That is the first time this project has confirmed a chart it drew
  is VISIBLE anywhere but in a human's eyes. Sharing a slide is the price of
  never rasterising a fresh one, and the bill came the same round: two full-size
  charts drawn over each other, which a rasteriser reads fine and a human calls
  broken. The scenario's chart is 30% of the slide in the bottom-right corner
  now — this battery leaves its slides in the deck for someone to look at, so
  "the measurement still works" is not the bar.

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
  **Done deliberately once, on 2026-08-09, and it paid immediately**: probing
  all three with the same six inputs found the same defect in all three in one
  pass, rather than one at a time over three sessions. The defect: `parseFloat`
  is looser than the regex feeding it — `[\d.]+` matches a bare `.` and
  `parseFloat(".")` is NaN — so `hsl(., 50%, 50%)` **threw** (the hue sector
  table has no NaN entry, so destructuring `undefined` blew up) and
  `rgb(., ., .)` returned `#NaNNaNNaN`, which in the pptx sink is not the six
  hex digits that path's own security note requires. Both are now enforced at
  the root — `rgbToHex` for the preview, an exit check in `hex()` for pptx —
  rather than per branch, and `officeHex` inherits the first.
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

  Both experiments have now run. _the chart is actually visible_ PASSED on
  `c7d91d5` and is routine again. _what makes a long run slow down_ ran on
  `25407ed`, picked alone, and KILLED THE TAB at 26.9 seconds with two scenarios
  behind it — announcing itself and then nothing, which is exactly the state
  `chartIsVisible` spent four rounds in. Two calls fit and the log cannot choose
  between them: taking a scratch slide, or drawing ninety-six shapes onto one
  the run had just added. It is instrumented now (`degradation step`), so the
  next picked round names the call instead of the scenario. Do not reason about
  which; that is what the steps are for. `PENDING_QUESTIONS` in
  `scripts/host-baseline.mjs` is the authoritative list of unasked probe
  questions — **read the table rather than a count here**, which is the rule
  the `UNSTABLE_ANSWERS` paragraph above already states and which this sentence
  broke within a day of being written: it said "nine down to one" and a probe
  added on 2026-08-09 (`binding-names-shape-later`) made it two.
  _edit the chart YOU click_ came off this list on 2026-08-08 — it PASSED.

  **A never-asked answer is not an answer, and the gate used to think it was.**
  `no-scratch-slide` and `no-scratch-shape` mean the run could not put the
  question. The contract gate read their absence from `differ` as agreement and
  their presence in the sheet as an answer, so the first sheet carrying them
  demanded that two `KNOWN_DIVERGENCES` entries be deleted and a
  `PENDING_QUESTIONS` entry retired — deleting knowledge on the strength of a
  setup failure. Both directions now treat `notAsked` as unknown.

  **Nothing is owed to the owner right now.** The manifest re-install he was
  asked for was done on 2026-08-06, and nothing since has touched a manifest —
  do not ask again unless a PR actually changes one. What to ask him to click is
  written down: "The standing test run" in `docs/PUBLISHING.md`. Don't improvise
  a new one per session, and don't ask for the deck or a screenshot — the round's
  own file has carried both since the deck-evidence change.

  **Eleven rounds have now run, and the battery reports.** The committed fixture
  is whichever sheet last answered at least as much as the one before it — read
  its `build` field rather than trusting a commit named here, because this
  paragraph has already been wrong about that once. `layouts-readable` and
  `slide-layout-readable` both answer `yes`, which retires office-js#4906 and
  #3826 as exposures on this host and this deck.

  **Both picked-alone experiments ran and answered, and neither is pending.**
  _the chart is actually visible_ is routine again and has PASSED twice
  (`c7d91d5`, and again on `1cd7ea3` at 15704 → 16580 bytes). _what makes a long
  run slow down_ has MEASURED — twice, on two builds — and the quadratic
  per-slide cost written up above is its result; it is `pickedOnly` and does not
  need running again unless that number is in doubt.

  **The stall investigation is closed as far as candidates go.** Eleven rounds
  eliminated the scenario, the preceding scenario, the tab's age, the idle gap
  before the sync, and the identity of the preceding call — each by a later
  round putting the candidate in both populations. What remains is a host that
  stalls the first sync of a draw intermittently, one or two draws in fifteen,
  with no known discriminator. `does a rasterise poison the next draw` is the
  counterbalanced control that will say "no pattern" until something changes;
  a few rounds of that IS the finding, and is the point at which to stop
  instrumenting and live with it.

- **Phase 3 — activate the Claude skill** (upload the zip on claude.ai).

Follow it phase by phase; retire items from it and from this list as they
complete.
