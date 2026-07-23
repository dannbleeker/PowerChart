# Tests

Vitest discovers every `test/*.test.ts` automatically — there is no manifest to
update. Run the suite with `npm test`, thresholds with `npm run coverage`.

## Naming convention

**Name a test file by the topic it covers, never by the increment that shipped
it.** A chart kind's tests live in `<kind>.test.ts`; a cross-cutting feature's
tests live in `<feature>.test.ts`. If you add a kind or feature, add or extend
the matching file — don't start a `batch-N` / `backlog-x` grab-bag. (The old
`backlog-a…t`, `bug-hunt`, `hunt-*`, `r2-*` files were exactly that, and were
split back out into the topic files below.)

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
  (cross-kind extent/auto-scale invariants), `geometry`, `good-chart*`.
- **Renderers & app** — `office-render` (Office.js fake host), `pptx-paint`
  (headless pptx node mapping), `pane-state` / `pane-host-actions` /
  `pane-widgets` / `dom-pane` (task pane), `skill*`, `parity`, `snapshots`,
  `a11y-svg`, `security-*`, `dark-theme`, `fuzz`, `hardening`, `degenerate-inputs`.

## Lockstep-gated files — do not rename

These enforce the feature-set lockstep (see `CONTRIBUTING.md`) and are referenced
by name from `CONTRIBUTING.md`, the PR template, and build scripts:
`skill-docs.test.ts`, `showcase.test.ts`, `manual.test.ts`, `snapshots.test.ts`.
