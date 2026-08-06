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
- **Hostile input** — `chart-hostile-input`, now four sweeps rather than one:
  every value a **cell** can hold (huge, subnormal, NaN, infinite, empty,
  degenerate sizes), every **style** field with the wrong type, every
  **top-level key** with the wrong type through both offline renderers, every
  **decoration** key, and eighteen malformed **data shapes**. The cell sweep's
  bar is _termination_ — both bugs it found were loops whose bound came from the
  data. The type sweeps' bar is simply _not throwing_: their bugs were all one
  shape, a `string` API meeting a non-string out of user JSON, and they were
  spread across every layer.
  The live Office renderer gets the same treatment in `office-render`, which
  matters because it is the one no earlier sweep had ever pushed a malformed
  config through — and it held its own copy of the colour bug.
- **Repair planner** — `reconcile` (the rules), `reconcile-fuzz` (four thousand
  generated decks against the invariants that cost a user their work: never
  delete another run's slide, never delete every copy of an item, never act on
  a slide that is not there).
- **Renderers & app** — `office-render` (Office.js against the fake host),
  `web-host` (the same renderer against that host at its WORST — every
  misbehaviour a real PowerPoint on the web has shown us, on at once), `pptx-paint`
  (headless pptx node mapping), `svg-render` (SVG node emission — paths,
  polygons, options), `pane-state` / `pane-host-actions` / `pane-widgets` /
  `dom-pane` (task pane), `crashlog` (the record that outlives a run that
  never ends), `templates` (saving and re-picking a chart setup — a whole
  feature that had no tests until one of them turned up a bug),
  `host-probe` (the fake's own frozen answer sheet — what it CLAIMS about the
  host it stands for, so a real PowerPoint can be diffed against it),
  `host-contract` (the other half of that: `FAKE_BASELINE` diffed against a real
  PowerPoint's committed sheet, so a new divergence fails in CI in seconds
  instead of waiting for somebody with the app open), `selftest` (the in-host
  battery's own logic — that every scenario reports, that a blind deck scan is
  never read as an empty deck, and that a wedged host is attributed to the host),
  `skill*`, `parity`, `snapshots`, `a11y-svg`, `security-*`, `dark-theme`,
  `fuzz`, `hardening`, `degenerate-inputs`.
- **Things CI can check that are not code** — `manifest` (the four add-in
  manifests: Office's version floor, the `<Id>` GUIDs, no localhost surviving
  into a production manifest — pinned offline because the real validator calls a
  Microsoft service and cannot run everywhere), `office-js-watch` (the weekly
  tracker sweep's matching half, including that its `KNOWN_ISSUES` table covers
  every office-js issue the codebase cites and that every watched term is a call
  this repo actually makes), `visible-charts` (the verdict half of the visual
  gate — the rasterising half runs in a real browser and cannot run here).
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

`newSlideResolvesTimes` is the one that came from an answer sheet rather than a
crash. PowerPoint on the web resolved a freshly-added slide's id **once** and
refused it ever after — while still listing that same id in
`slides.load("items/id")`, so the slide was plainly there and only the lookup
was broken. `null` means ids stay good (every other host); a number is how many
`getItemOrNullObject` calls each added slide answers before it starts reporting
gone, and `0` means never, starting now. That `0` matters: a count-based version
of the delete test passed against the very code it was written to falsify,
because the fix spends one lookup more than the bug did and the lease outlasted
both.

`selectionWedgesHost` is the odd one out and the newest: it models a call that
is taken and _poisons later ones_. A programmatic `setSelectedShapes` succeeds,
and every selection sync after it then never settles at all — neither resolving
nor rejecting, which is what PowerPoint on the web actually does and the one
failure shape no `catch` can see. Only a bounded wait survives it, so it is the
fault that proves the bounds work.

The newest family came from answer sheets rather than from crashes, and each
models something a real PowerPoint on the web was measured doing: a freshly-added
slide's handle dying at the next sync (`expiringSlideHandle`, unconditional
because it is established); a shape that cannot be NAMED in the batch that
created it (`noIdInCreatingSync` — implied by five questions being the only five
never put); `getItem` refusing a slide this run added (`refuseGetItemOnNewSlide`);
and `addTextBox` deleting the selected shape (`textBoxDeletesSelection`,
office-js#2775). The last two are OFF by default and say so at their definitions:
nothing has established that the build the owner runs still does either, and a
fake that asserts unasked host behaviour turns every real answer into a
divergence.

`syncCostMs` is not a misbehaviour at all — it is a clock. It charges each sync
by a function of `syncsInContext` or `syncsTotal`, which are exactly the two
hypotheses the degradation experiment separates, and without it a measurement
could only be tested against a fake that is always instant.

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
