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

A second sweep (2026, against think-cell 13/14's 2025 additions, the IBCS
notation standard, and the FT/Zelazny/Few design canon) surfaced the candidates
below. The chart-*type* space stays saturated — these are notation, a summary
label, an analytic fit, and two conveniences. None is started; each needs the
owner's go-ahead before it graduates to a PR. Ranked by value ÷ effort.

- **IBCS scenario notation for the column/bar family** — *high value, medium
  effort.* A per-series `scenario` tag (`AC` actual / `PY` previous-year / `PL`
  plan / `BU` budget / `FC` forecast) that maps to IBCS's standard fill
  semantics: actual = solid dark, previous-year = lighter solid, plan/budget =
  **outlined/hollow**, forecast = **hatched**, plus the two-letter tags and an
  auto-legend. This is the project's stated "business communication standards"
  identity made systematic: the ad-hoc pieces already exist (line
  `forecastFrom`, waterfall Target-gap, Gantt baseline ghost, `series.pattern`
  hatch), but there is no scenario encoding for columns/bars. All primitives are
  present **except one**: an outlined/hollow column. `fill: "transparent"`
  currently drops the rect entirely (`column.ts`), so a framed-empty column is
  not expressible — the foundational sub-task is a **stroke-only fill mode**
  (rect with `fill:none` + `stroke`), which is native in all three renderers and
  unblocks Plan/Budget. FC-hatch degrades to solid in the live add-in — the
  already-accepted pattern-degrade path. Lockstep cost: a new top-level/decoration
  key across SKILL/reference/MANUAL/showcase/README. (IBCS v1.1 §notation.)

- **Grand total label** — *high value, low effort, low risk.* A single label at
  the top-right of a stacked/clustered column chart (and %-axis Mekko) showing
  the **sum of every category total**, distinct from the per-column totals that
  already ship. New in think-cell 14. A pure `text` scene node — no new
  primitive, identical in all three renderers, round-trips as a decoration key
  automatically. Only nuance is placement/de-collision against the value axis and
  the top segment label, which the existing `collide.ts` pass already handles for
  `total-` names. The cleanest quick win in this list.

- **IBCS variance columns (absolute Δ and relative Δ%)** — *medium value, medium
  effort.* Given a primary (AC) and a reference (PL/PY) series, append a small
  variance tier: signed Δ columns and/or Δ% with a zero reference line. Native
  rects + a line, a defined IBCS chart family. Partially pre-empted by the
  existing difference arrows and waterfall budget-gap; the new part is a
  *systematic per-category* tier. Best sequenced **after** scenario notation, so
  the `scenario` tags identify which series is actual vs reference.

- **Polynomial scatter trendlines (quadratic / cubic / quartic)** — *medium
  value, medium effort.* Extend the scatter `Trend` row beyond OLS-linear to a
  degree 2–4 least-squares fit, still reporting R². New in think-cell 14. Drawn
  as a **fine polyline** (sample ~40 x-positions → `line` segments), the same way
  the live add-in already renders any curve — no Bézier needed. Distinct from the
  `smooth` decoration, which is Catmull-Rom *data* smoothing, not a regression.
  Needs an overfitting guard (degree ≤ points − 1) and a sane default degree; the
  live add-in pays in shape count for the polyline.

- **PNG export of the preview** — *low–medium value, low effort.* A "Download
  PNG" beside the existing "Download SVG": rasterize the preview SVG through a
  canvas (`drawImage` → `toBlob`). Pure browser, no Office API; the SVG uses no
  `foreignObject`/external refs so the canvas won't taint. Useful for pasting
  into email/chat where SVG doesn't render. A convenience on a fallback path
  (the product's real output is native shapes), not a gap of principle.

- **Copy-config-as-URL (deep-link share)** — *low–medium value, low effort.*
  "Copy shareable link" that base64-encodes `currentConfig()` into the URL hash;
  the pane hydrates from it on load (the `?kind=` deep-link plumbing already
  exists). Overlaps JSON export / localStorage templates, and long configs make
  long URLs, but it is cheap and good for support/collaboration.

**Considered and dropped this sweep** (so they aren't re-proposed): a
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
