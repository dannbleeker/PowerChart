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
- **Hostile input** — `chart-hostile-input` (every chart kind against every
  value a cell can hold: huge, subnormal, NaN, infinite, empty, degenerate
  sizes). Its bar is _termination_, not output — both bugs it found were loops
  whose bound came from the data.
- **Repair planner** — `reconcile` (the rules), `reconcile-fuzz` (four thousand
  generated decks against the invariants that cost a user their work: never
  delete another run's slide, never delete every copy of an item, never act on
  a slide that is not there).
- **Renderers & app** — `office-render` (Office.js against the fake host),
  `web-host` (the same renderer against that host at its WORST — every
  misbehaviour a real PowerPoint on the web has shown us, on at once), `pptx-paint`
  (headless pptx node mapping), `svg-render` (SVG node emission — paths,
  polygons, options), `pane-state` / `pane-host-actions` / `pane-widgets` /
  `dom-pane` (task pane), `skill*`, `parity`, `snapshots`, `a11y-svg`,
  `security-*`, `dark-theme`, `fuzz`, `hardening`, `degenerate-inputs`.
- **The deck a run produces, audited from its bytes** — `verify-deck` (did
  PowerChart write what it meant to: slot tags, groups, config parts, shape ids
  unique per slide), `ooxml-validate` (is it a legal `.pptx` at all, against the
  OOXML grammar, plus the one baselined finding it is allowed to have),
  `triage` (joining a saved deck to the run log that produced it). The first two
  catch nearly disjoint sets and both gate CI on `examples/showcase.pptx`.

## The fake PowerPoint host

`helpers/office-host.ts` is the Office.js double: recording proxies for shapes,
slides, tags and groups, plus a `faults` object of misbehaviours. Every fault in
it was added AFTER a real host taught us the behaviour — a stale shape proxy
refused by `getItem(id)`, a shape collection reading back shorter than it is
without throwing, a refused `addGroup`, a sync that answers minutes late, a
collection read the host never answers at all, a shape whose position stays
unreadable until the load that asked for it lands.

One family of them models the same thing at four levels — the host taking a
load and answering nothing: `unansweredShapeReads` (a whole shape collection),
`unansweredNullChecks` (one `getItemOrNullObject` proxy), `unansweredTagLoads`
(a tag), and `faults.strictShapeReads` (a shape's own `id`/`left`/`top`).
Reading any unanswered proxy is `PropertyNotLoaded` on a real host and was a
plain value here, which is how three self-test scenarios could fail in a real
PowerPoint against a fully green suite. `strictShapeReads` is off by default
only because the fake's shape objects double as the surface tests assert
geometry against — the others are always in force once armed.

`selectionWedgesHost` is the odd one out and the newest: it models a call that
is taken and _poisons later ones_. A programmatic `setSelectedShapes` succeeds,
and every selection sync after it then never settles at all — neither resolving
nor rejecting, which is what PowerPoint on the web actually does and the one
failure shape no `catch` can see. Only a bounded wait survives it, so it is the
fault that proves the bounds work.

Faults are opt-in per test. `applyWebProfile()` turns on the set a real web host
shows at once; call it AFTER `installHost`, which resets every fault. It is not
the default, because applied everywhere it would fail hundreds of tests for
reasons that have nothing to do with what they assert.

**It can take a generated `.pptx`.** `insertSlidesFromBase64` really decodes the
bytes — through `scripts/verify-deck.mjs`, the same decoder the audit tool uses,
so the fake cannot read a generated deck differently from the tool that checks
one — and materialises each slide with its slot tag and a `PowerChart` group
holding as many children as the file holds. The child count matters: a readback
measures a chart by what is inside its group, and a fake that put one shape
there made every generated chart read back as wreckage.

**`shapes.items` hands back fresh handles and leaves earlier ones stale**, the
way real Office.js does. The fake used to refresh the shape objects themselves,
so a re-fetch anywhere healed a stale proxy held anywhere — and that one
kindness is why a whole class of stale-proxy bug could only be found by a human
running the add-in in a real PowerPoint. Do not "simplify" it back.

## Lockstep-gated files — do not rename

These enforce the feature-set lockstep (see `CONTRIBUTING.md`) and are referenced
by name from `CONTRIBUTING.md`, the PR template, and build scripts:
`skill-docs.test.ts`, `showcase.test.ts`, `manual.test.ts`, `snapshots.test.ts`.
