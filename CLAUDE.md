# PowerChart — project memory

Open-source think-cell clone: a PowerPoint charting add-in whose charts are
**native, editable shapes** (never pictures), plus a Claude Agent Skill that
renders the same charts headlessly. 18 chart kinds, think-cell's signature
decorations, and "the good chart" design formalia are all implemented — see
the README feature table for the authoritative list.

## Architecture in one paragraph

`ChartConfig` (pure JSON) → `buildChart()` in `src/core` (pure TypeScript, no
Office imports) → renderer-agnostic scene graph (`rect | line | text | ellipse
| wedge | chevron | arrowhead | polygon` nodes, coordinates in points) → three
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
npm test           # full suite (380+); npm run coverage enforces thresholds
npm run typecheck
npm run showcase   # regenerate the showcase deck (required after feature work)
npm run skill      # build skill-dist/powerchart-charts.zip
```

## Working conventions (established with the repo owner)

- **Branch flow**: develop on `claude/think-cell-clone-research-5tgx3v`; after
  each merge, reset it onto `origin/main` (`git checkout -B <branch>
  origin/main`) — never stack on merged history. One PR per increment.
- **Auto-merge is authorized**: once CI is green on the exact pushed commit
  (verify `head_sha` matches local HEAD), merge the PR to main without asking.
- **Snapshots** (`test/snapshots.test.ts`) freeze every sample chart's SVG.
  Only update (`vitest -u`) after reviewing renders visually — screenshot via
  Playwright (`/opt/pw-browsers/chromium`) and inspect before accepting.
- **Visual QA is part of done**: render new features to SVG → PNG and look at
  them; several real bugs were only caught this way.
- **Releases**: merges to main refresh the rolling `skill-latest` prerelease.
  Versioned releases via the Release workflow's manual dispatch (the git proxy
  rejects tag pushes — dispatch creates the tag server-side).

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
- GitHub MCP `actions_list` responses exceed the token cap — parse the saved
  JSON file with python instead of reading it.

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

## Pending / user-gated

`docs/PUBLISHING.md` is the go-live runbook (written for an Opus 4.8
session + the owner): make the repo public, host on GitHub Pages, generate
production manifests, sideload + validate in real PowerPoint, activate the
Claude skill, cut v0.2.0. Follow it phase by phase; retire items from it
and from this list as they complete.
