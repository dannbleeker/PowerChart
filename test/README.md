# Tests

Vitest discovers every `test/*.test.ts` automatically — there is no manifest to
update. Run the suite with `npm test`, thresholds with `npm run coverage`.

## Naming convention

**Name a test file by the topic it covers, never by the increment that shipped
it.** A chart kind's tests live in `<kind>.test.ts`; a cross-cutting feature's
tests live in `<feature>.test.ts`. If you add a kind or feature, add or extend
the matching file — don't start a `batch-N` / `backlog-x` grab-bag. (The old
`backlog-a…t`, `bug-hunt`, `hunt-*`, `r2-*` files were exactly that, and were
split back out into the topic files below. So were `coverage-core` /
`coverage-branches`, which named a _mechanism_ rather than a topic — a test
belongs with the thing it tests, not with the reason it was written.)

## Where things live

- **Chart kinds** — one file each: `waterfall`, `column`, `line`, `combo`,
  `scatter`, `pie`, `sunburst`, `radar`, `radial-bar`, `boxplot`, `violin`,
  `candlestick`, `heatmap`, `tilemap`, `treemap`, `gantt`, `funnel`, `butterfly`,
  `waffle`, `gauge`, `bump`, `pareto`, `sparkline`, `bullet`, `cascade`,
  `elements`.
- **Cross-cutting features** — `axis-features` / `axis-scale` (axes, scales, log
  floors), `format` / `format-edge` (number & label formatting), `dates`,
  `palette`, `data-sorting`, `legend-layout`, `decor-guards` /
  `decoration-layout` (decoration clipping & anchoring), `value-extent`
  (cross-kind extent/auto-scale invariants), `geometry`, `color` (paint parsing
  & contrast ink), `collide` (label collision resolution), `good-chart*`.
- **Renderers & app** — `office-render` (Office.js against the fake host),
  `web-host` (the same renderer against that host at its WORST — every
  misbehaviour a real PowerPoint on the web has shown us, on at once), `pptx-paint`
  (headless pptx node mapping), `svg-render` (SVG node emission — paths,
  polygons, options), `pane-state` / `pane-host-actions` / `pane-widgets` /
  `dom-pane` (task pane), `skill*`, `parity`, `snapshots`, `a11y-svg`,
  `security-*`, `dark-theme`, `fuzz`, `hardening`, `degenerate-inputs`.

## The fake PowerPoint host

`helpers/office-host.ts` is the Office.js double: recording proxies for shapes,
slides, tags and groups, plus a `faults` object of misbehaviours. Every fault in
it was added AFTER a real host taught us the behaviour — a stale shape proxy
refused by `getItem(id)`, a shape collection reading back shorter than it is
without throwing, a refused `addGroup`, a sync that answers minutes late.

Faults are opt-in per test. `applyWebProfile()` turns on the set a real web host
shows at once; call it AFTER `installHost`, which resets every fault. It is not
the default, because applied everywhere it would fail hundreds of tests for
reasons that have nothing to do with what they assert.

**`shapes.items` hands back fresh handles and leaves earlier ones stale**, the
way real Office.js does. The fake used to refresh the shape objects themselves,
so a re-fetch anywhere healed a stale proxy held anywhere — and that one
kindness is why a whole class of stale-proxy bug could only be found by a human
running the add-in in a real PowerPoint. Do not "simplify" it back.

## Lockstep-gated files — do not rename

These enforce the feature-set lockstep (see `CONTRIBUTING.md`) and are referenced
by name from `CONTRIBUTING.md`, the PR template, and build scripts:
`skill-docs.test.ts`, `showcase.test.ts`, `manual.test.ts`, `snapshots.test.ts`.
