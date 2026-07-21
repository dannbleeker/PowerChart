# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).

**This is the only backlog document.** Items graduate from here into PRs and are
deleted when they ship — what has shipped is recorded by the README feature
table and by git, not here. Rejected ideas stay in §3 so they aren't
re-proposed.

Feasibility is judged against the live-add-in constraint: rects, lines, text,
ellipses, and any of PowerPoint's 177 preset geometries (all of them
PowerPointApi 1.4), plus polygon *outlines* — no freeform curves, and no
images. The SVG and skill-pptx renderers additionally have filled polygons and
patterns.

## 1. Open

- **SVG/PNG image output as an alternative render mode** — offer, beside the
  native-shapes renderer, a mode that inserts the whole chart as a *single image
  object*. The scene→SVG renderer already exists (`src/render/svg.ts`) and the
  canvas rasterisation primitive now ships too (the pane's PNG download does
  `drawImage`→`toDataURL` on the preview SVG), so the pixels are free on both
  paths; the remaining work is the *insert* side. The scene graph is pure
  geometry/text with no `foreignObject` or remote assets, so the canvas stays
  untainted. The motive is performance/reliability, not looks — one object
  sidesteps the PowerPoint-**web** wall where the live canvas stops answering past
  ~20 shapes per `context.sync()` (office-js #4272, #5022, #6498), which is
  exactly what caps the densest kinds today (`DEMO_SHAPE_BUDGET = 90`,
  `SHAPES_PER_SYNC = 10`). Insert paths, all shipped: `setSelectedDataAsync(png,
  {coercionType: Image})` (ImageCoercion 1.1, widest host reach) or
  `addGeometricShape` + `ShapeFill.setImage(png)` (PowerPointApi 1.8) — the latter
  returns a tracked `Shape`, so the `POWERCHART_CONFIG` tag round-trips onto the
  picture exactly as it does now, keeping the chart re-editable. Native-vector SVG
  via `setSelectedDataAsync(svg, {coercionType: XmlSvg})` (ImageCoercion 1.2) is
  the best quality/size answer — PowerPoint can even "Convert to Shape" it — but
  it's flaky (office-js #4967 open: vertical shift; #2881: complex SVG; #412: no
  iOS), so gate it behind a flag with a PNG fallback. The cost is real: an image
  is not editable in PowerPoint, ignores theme colours, and blurs on rescale
  unless rasterised high-DPI — so pair it with an **"explode to native shapes"**
  command that reads the tag, deletes the picture, and draws the real shapes on
  demand (paying the web cost only when the user actually edits). This is distinct
  from the rejected per-point **image / icon node** in §3: that was one bitmap
  *per data point* inside the scene graph, unreachable at the 1.4 pin; this is one
  image for the *whole chart* as an output format, with the `1.8`/ImageCoercion
  gates degrading to the native-shapes path on older hosts (like grouping
  already does). Feasibility: medium — needs Phase-2 real-host validation first,
  since the `setSelectedDataAsync` base64 size cap (office-js #225, fixed, no
  published threshold) and the shape-tag value-size limit are both undocumented.

Otherwise nothing new: the second sweep's six candidates all shipped (grand total
label, IBCS scenario notation + the stroke-only hollow-column primitive, IBCS
variance tier, polynomial scatter trendlines, PNG export, copy-config-as-URL) —
see the README feature table and git for what landed. Whatever else surfaces
starts from a fresh research pass.

**Considered and dropped in that sweep** (so they aren't re-proposed): a
screen-reader data-table alternative to `describeChart` — real WCAG best practice
beyond ~6 points, but the primary output is native PPT shapes where alt text is
linear only, so the gain is confined to the downloaded SVG; **CSV file import** —
already covered by the datasheet's TSV clipboard paste, which is what
Excel/Sheets put on the clipboard; **variance/integrated bars as a new kind**,
tornado, icicle, fan, Venn — recipes of shipped kinds or off-genre/curve-bound
(see §3's standing rejections of recipe-of-existing-kind proposals).

## 2. Residue of the July 2026 adversarial bug hunt

The hunt found 59 defects across 12 lenses, each reproduced by execution and
each re-verified by an adversarial verifier that defaulted to REFUTED. All 59
are fixed (PRs #186-#197) with a regression guard apiece, every guard proven
non-vacuous against the pre-fix file. What ships is recorded in the CHANGELOG
and in git; only what is still OPEN is listed here.

- **`formatPercent`'s locale is threaded through the axis/segment path but not
  the rest** — `src/core/decor.ts`, `funnel.ts`, `cascade.ts`, `mekko.ts`,
  `waffle.ts`. #193 made `formatPercent` locale-aware and passed
  `numberFormat.locale` from `segmentLabel`, which fixed the case the hunt
  measured (a de-DE funnel printing "35.8%" beside formatNumber's "12.000").
  The remaining call sites — the CAGR and difference arrows, and the four
  layouts' own percent labels — still call it without the locale, so a
  localized chart can mix number systems if those decorations are on. Small and
  mechanical; needs one fixture per call site so the guard isn't vacuous.

## 3. Rejected or already covered (do not re-propose)

- **An image / icon node** — not reachable in the live add-in, so nothing can
  be built on it. PowerPoint's `ShapeCollection` exposes exactly
  `addGeometricShape`, `addGroup`, `addLine`, `addTable`, `addTextBox`;
  `addImage` exists only in the Excel namespace. The one route to pixels is
  `ShapeFill.setImage`, which is **PowerPointApi 1.8** against a manifest
  pinned at **1.4** — it would vanish silently on any older host. It also
  bloats twice over: `ChartConfig` round-trips through a shape tag verbatim
  (`tagData: JSON.stringify(cfg)`), and PptxgenJS does not dedupe media, so an
  identical icon is re-embedded once per point. The only thing presets cannot
  do is a **logo** (a competitor-positioning map). If that is ever wanted, the
  shape is: user-supplied data URI (no asset library, so no licensing call),
  gated on `supports("1.8")` like grouping already is, and drawn *additively
  over* a marker so an old host degrades to a plain point rather than a
  missing one.
- **Scatter/bubble point icons** — the primitive above is unavailable, and the
  real itch was not bitmaps: shape as a categorical channel is what Excel and
  Highcharts mean by point markers, and PowerPoint's presets give it filled and
  native at 1.4. Shipped as `scatter.markers`.
- **Heatmap per-cell icon overlays** — a heatmap's rows ARE its data series, so
  a "Glyph" row yields one row of glyphs, not a per-cell matrix; carrying a
  genuine second dimension would mean inventing a datasheet convention for
  every heatmap, for a want nobody has stated. The demonstrable gap was
  narrower and is fixed: a diverging scale states direction in hue alone, and
  its strongest + and − sit at 1.12:1 in greyscale — the same tone — with no
  label to fall back on under `sizeEncode`. Shipped as `heatmap.symbols`.
- **An X / cross marker, and a five-point star** — a marker shape has to
  reproduce its OOXML preset *exactly*, because the SVG renderer draws the
  points while the PowerPoint renderers name the preset, and `markerScale`
  measures the area off those same points. A shape that only approximates its
  preset therefore breaks the bubble's "area ∝ size" claim in the deck while
  keeping it in the preview — the worst kind of divergence, since the preview
  looks right. `mathMultiply` fails on angled arms (and is redundant with
  `plus` at 3-4pt). `star5` fails on its `hf`/`vf` stretch: the preset widens
  itself 1.05146x and heightens itself 1.10557x so the star fills its box
  (a 5-point star spans 1.902R by 1.809R, and those factors are exactly
  2/1.902 and 2/1.809), which makes PowerPoint's star **16.2% larger in area**
  than an inscribed SVG one. Reproducing that faithfully is possible on paper
  but not verifiable here without a PowerPoint rasteriser, so the set stops at
  the shapes whose geometry is exact: circle, square, diamond, triangle, plus.

- **Waterfall connector re-routing** (drag to skip columns) — the drag is out of
  Office.js reach, and the rendering feature underneath is ill-posed: the
  connector is not an object with authored endpoints, it is a derived assertion
  that "this level carries into the next bar". Re-routing it is truthful only
  where the skipped columns leave the level unchanged — and there it already
  works, via `spacerIndices` and `totalIndices`. Anywhere else the line would
  end in mid-air pointing at a bar it does not touch. What was actually wanted
  shipped instead as `waterfall.detailGroups`.
- **Scatter-on-combo** — a continuous-x scatter over categorical columns needs a
  second x scale that means nothing beside category slots. The coherent reading
  — unconnected marks at category positions — shipped as `type: "marker"`
  (overlay) and already existed as `decorations.barStyle: "dot"` (clustered).
- **Free 2D bubble repulsion** — both axes carry data, so moving a marker in 2D
  corrupts two readings at once with nothing to bound them. The honest version
  shipped as `scatter.spread`: one named axis, hard-capped, cap printed in the
  footnote.
- **Gantt resource capacity vs load** — a per-resource per-week histogram against
  a capacity line. That is a stacked column chart with a `Target` row, both
  shipped; it needs a value axis, which a timeline does not have. Recipe in
  docs/MANUAL.md. (Lane grouping itself shipped as `gantt.lanes`.)
- **Dial / needle gauge** — bullet charts replace it deliberately (Few); low
  deck demand. Note the *semi-circle scorecard* gauge did ship (`doughnut` +
  `pie.semi`); what stays rejected is the dial-with-needle and its threshold
  bands.
- **Sankey / chord / arc** — need curved ribbons; infographic genre.
- **Ridgeline** — stacked density curves; academic register. (The single-column
  `violin` kind shipped at owner request; ridgeline stacking not pursued.)
- **Stream graph** — feasible but editorial aesthetic; no deck demand.
- **Pictogram with icon libraries** — needs the image node above, which is not
  reachable in the live add-in at all (see the first entry), plus an asset
  library. Waffle is the deliberate substitute: it covers the part-to-whole
  genre with square cells, it does not render icons.
- **Histogram as a kind** — the look is `clustered` + `gapWidth: 0`, both
  shipped. If this is ever revisited, auto-binning raw samples into categories
  is the only real gap; the bar geometry is not. (`histogramBins` in
  src/core/format.ts bins over a fixed domain, but nothing derives categories.)
- **Choropleth maps, 3D, drill-down interactivity** — out of scope by design
  (see CLAUDE.md). Tilemap proportional-area cartograms and tilemap drill-down
  fall here too: hard/infeasible.
- **Population pyramid, plain dot chart** — already covered by `butterfly`
  (+ `butterfly.split`) and `decorations.barStyle: "dot"`.
- **Radar vertex markers** — already there: radar emits `marker-*` ellipse
  nodes, which the Office.js renderer draws, so they appear in the live add-in
  too.
- ~~Candlestick / OHLC~~ — shipped as the `candlestick` kind at owner request,
  despite the thin consulting-demand signal.
- **Alt text in the headless pptx renderer** — the hunt flagged the missing
  `descr=` as a break of the SVG/Office.js/pptx parity contract, but pptxgenjs
  exposes alt text only on pictures and native charts, and every PowerChart
  shape is an `addShape`/`addText` autoshape. There is no seam to write it
  through short of hand-patching the generated OOXML. Recorded as a documented
  limit in `skill/reference.md` (#197) alongside the fixed deck font.
