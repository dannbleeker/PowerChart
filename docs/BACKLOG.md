# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).

**This is the only backlog document.** Items graduate from here into PRs and are
deleted when they ship — what has shipped is recorded by the README feature
table and by git, not here. Rejected ideas stay in §2 so they aren't
re-proposed.

Feasibility is judged against the live-add-in constraint: rects, lines, text,
ellipses, and any of PowerPoint's 177 preset geometries (all of them
PowerPointApi 1.4), plus polygon *outlines* — no freeform curves, and no
images. The SVG and skill-pptx renderers additionally have filled polygons and
patterns.

## 1. Open

### Non-finite geometry can leave `buildChart` and reach the .pptx

**Found:** 2026-08-01, by fuzzing every chart kind against hostile cell values.
**Reproduction:** any of nine kinds with all values at `1e308`:

| kind | node |
| --- | --- |
| stacked, waterfall, combo | `rect.y = -Infinity` |
| stacked100 | `line.y1 = NaN` |
| mekko | `rect.w = NaN` |
| area | `rect.y = NaN` |
| radar | `polygon.points[0].x = Infinity` |
| sunburst | `wedge.endAngle = NaN` |
| treemap (tiny values, `5e-324`) | `rect.w = Infinity` |

**Why it matters:** the SVG renderer already neutralises this — every numeric
goes through `num()` in `svg.ts` — but the two PowerPoint renderers do not. A
`NaN` coordinate goes straight into `addGeometricShape({left: NaN, …})` and into
OOXML as an EMU value, so the produced deck is one PowerPoint may refuse to
open. Nobody has hit it because nobody types `1e308`; it is reachable, not
likely.

**Shape of the fix:** one sanitising pass at the scene boundary, at the end of
`buildChart`, so all three renderers inherit the guarantee and no layout module
has to remember it. Deliberately NOT rushed in alongside the two hang fixes: it
touches the geometry of every chart, and this repo's convention is that visual
QA is part of done — render the samples, look at them, then land it.

**Already fixed, and not this:** the two *hangs* the same fuzz found are
shipped, with the fuzz kept as `test/chart-hostile-input.test.ts`. Image mode shipped end to end — the engine key (`render`), the
skill CLI rasteriser, the Office.js picture insert, the pane's **Insert as
picture** control and the **Explode to native shapes** command that converts one
back. Deliberately NOT built, so it is not re-proposed: native-vector SVG via
`setSelectedDataAsync(svg, {coercionType: XmlSvg})` (ImageCoercion 1.2) is the
best quality/size answer and PowerPoint can even "Convert to Shape" it, but it is
flaky (office-js #4967 vertical shift, #2881 complex SVG, #412 no iOS), it writes
to the *selection* so it cannot honour an explicit position or an off-screen
slide, and it needs an `ImageCoercion` manifest entry — i.e. an owner re-sideload
— for a path the PNG route already covers.

Still unverified on a real host (all degrade safely, none block): whether a
`PictureAndTexture` fill stretches or tiles, whether `lineFormat.visible = false`
survives it, and the true payload cap — `MAX_PICTURE_BASE64` is a 4 MB guard at
~30x the worst measured payload, not a measurement. Unfixed by design: an image
chart is opaque to a screen reader below PowerPointApi 1.10, the one axis where
it is strictly worse than native shapes.

Otherwise nothing new: the second sweep's six candidates all shipped (grand total
label, IBCS scenario notation + the stroke-only hollow-column primitive, IBCS
variance tier, polynomial scatter trendlines, PNG export, copy-config-as-URL) —
see the README feature table and git for what landed. Whatever else surfaces
starts from a fresh research pass.

**No open defects.** Five adversarial bug hunts have run (PRs #186–#197,
#202–#210, and #229–#231); every confirmed finding is fixed with a regression
guard proven non-vacuous against the pre-fix file. Ten additional harness-reliability PRs
(#212–#216) closed the Phase-2 sideload observations from Presentation_2.pptx
and Presentation_3.pptx — trust-the-readback, slot-tag naming, fresh-context
group+tag rescue, results/contents pagination, larger off-screen batch size,
unstamp+rescue for failed items, and `addSlides` self-heal for the silent
`slides.add()` drop pattern (bug prep in `docs/OFFICE_JS_LOST_ADDS.md` for
office-js submission). What shipped is recorded in the README feature table and
in git, not here.

The fifth hunt (five parallel hunters over the work merged in #224–#228) is
worth summarising, because what it found says where the risk in this codebase
actually lives. Fourteen confirmed bugs, none of them in chart geometry or the
engine: every one was in the **recovery machinery** — the code that runs when
PowerPoint misbehaves, which by construction is the code that never runs in a
healthy test. Two were destructive on ordinary input. Inserting the demo deck a
second time deleted one of the two runs, because slot+title names an *item* and
not an occurrence of one; PowerPoint's own Duplicate Slide triggered the same
delete. And Same Scale deleted every chart's old shapes in one committed sync
before redrawing them one at a time, so a single stalled redraw left every
chart after it blank — on the deck-wide operation that necessarily includes the
chart on the visible slide, the one condition documented here as reliably
stalling a redraw.

The rest divide into three shapes worth watching for in new code:

- **Signals with no owner.** A global late-sync counter meant any stalled call
  anywhere degraded whichever run happened to be in flight. Ownership has to be
  captured when a call is *issued*, not when it answers.
- **Two failures conflated into one answer.** "Nothing landed" and "the new
  slide landed but the old one would not go" both returned `false`; "measured
  zero" and "could not measure" both returned zero; "the host dropped this
  slide" and "we could not read that page" both reported `lost`. Each collapse
  sent a caller confidently down the wrong branch, and each was one extra
  return value away from being right.
- **Diagnostics that were never exercised.** The run log was not written on the
  path a real run takes, so the feature added specifically to diagnose host
  failures produced nothing on the failing runs. A diagnostic needs a test on
  its *default* path, not only where it was easiest to add one.

Fixed with 24 new guards, each stashed against its pre-fix source and confirmed
to fail without it.

**Considered and dropped in that sweep** (so they aren't re-proposed): a
screen-reader data-table alternative to `describeChart` — real WCAG best practice
beyond ~6 points, but the primary output is native PPT shapes where alt text is
linear only, so the gain is confined to the downloaded SVG; **CSV file import** —
already covered by the datasheet's TSV clipboard paste, which is what
Excel/Sheets put on the clipboard; **variance/integrated bars as a new kind**,
tornado, icicle, fan, Venn — recipes of shipped kinds or off-genre/curve-bound
(see §2's standing rejections of recipe-of-existing-kind proposals).

## 2. Rejected or already covered (do not re-propose)

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
