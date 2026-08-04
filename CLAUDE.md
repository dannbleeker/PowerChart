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
  Playwright (`/opt/pw-browsers/chromium`) and inspect before accepting.
- **Visual QA is part of done**: render new features to SVG → PNG and look at
  them; several real bugs were only caught this way.
- **A regression test must be proven to fail without its fix.** Stash the
  source file, re-run, confirm the new test goes red, restore. A guard that
  passes against the pre-fix file is not a guard — it is decoration. Check
  _which_ assertion went red, too: a guard can fail for the wrong reason and
  still look proven (one here failed on a slide-swap rung that records nothing
  either way; another compared a `min()` both sides answered 0 for).
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
- **When moving tests between files, pin the total first and check it after.**
  A reorg once silently deleted 43 tests — the suite still went green because
  the count was never compared. `npx vitest run | grep "Tests "` before and
  after; if the number moved and you did not add or remove a case, stop.
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

## Gotchas

- Office.js has **no freeform paths**: pies are triangle fans, radar/polygon
  fills degrade to outlines in the live add-in (the skill's pptx output gets
  real filled `custGeom` polygons), pattern fills are SVG-only (solid in PPT).
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

- **Phase 2 — sideload + validate in real PowerPoint.** Nothing has ever run
  in a real host: every Office.js assertion in this repo is against a fake.
  Expect the first real run to surface things the mocked tests cannot. The two
  areas with the least mock fidelity are chart **positioning** (the
  `POWERCHART_ORIGIN` drag-delta round trip) and **grouping** on hosts that
  gate it behind `supports("1.8")`. What to actually ask the owner to click is
  written down — "The standing test run" in `docs/PUBLISHING.md`, six tests
  ordered by risk, about five minutes of manual work plus two one-click
  batteries. Don't improvise a new one per session.
- **Phase 3 — activate the Claude skill** (upload the zip on claude.ai).

Follow it phase by phase; retire items from it and from this list as they
complete.
