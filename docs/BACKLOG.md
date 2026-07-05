# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).
Every claim was verified against the source before listing — items the
engine already has were filtered out.

**This is the only backlog document.** Items graduate from here into PRs
(and then out of here); rejected ideas stay in §4 so they aren't re-proposed. (Shipped since the
sweep: the cascade/decomposition chart.)
Feasibility is judged against the live-add-in constraint (rects, lines,
text, ellipses, triangles, polygon *outlines* — no freeform curves; the SVG
and skill-pptx renderers additionally have filled polygons and patterns).

## 1. Top candidates (value ÷ effort, cross-cutting rank)

| # | Item | Type | Effort | Why first |
|---|---|---|---|---|
| 1 | **Bullet chart** — `Target` row (per-category tick) + existing bands | option | easy | KPI staple; ~80% already built (bands, error-row pattern) |
| 2 | **Combo: clustered / 100% columns + line** — `layoutCombo` hardcodes `"stacked"` | fix | easy | Excel's most common combo default is clustered+line |
| 3 | **Bubble size legend** — reference circles + labels | fix | easy | Near-correctness gap: bubble area is unreadable without a key |
| 4 | **Forecast styling on lines** — dashed continuation of actuals | option | easy | Extremely common actuals/plan split in trend decks |
| 5 | **Scatter quadrant preset** — one X+Y crossing → 4 shaded, labeled zones | option | easy | BCG-matrix framing in one step (bands+labels exist) |
| 6 | **Funnel / pyramid** — native `funnel`/`trapezoid` autoshapes | new kind | easy | Pipeline/conversion demand; zero rendering risk |
| 7 | **Lollipop + dumbbell/range styles** on clustered | option | easy | Two FT-taxonomy messages (comparison, gap) from stem+dot |
| 8 | **Gantt: percent-complete fill + baseline ghost bars** | option | easy | Plan-vs-actual + progress: the two biggest PM asks |
| 9 | **Grouped boxplots** (category × series sub-boxes) | option | medium | Most natural distribution-comparison extension |

## 2. Gaps within existing kinds (per kind)

### Column family (stacked / clustered / stacked100)
- **Per-column target marker** (tick at a value on one column) — easy, high;
  folds into the bullet-chart item above.
- **Automatic "Other" bucket** for long-tail series (think-cell's "Move to
  Other Series") — medium, medium.
- **Configurable gap width / column overlap** (Excel exposes 0–500%) —
  easy, medium; today `colThick` is hardcoded at 2/3 slot.
- **Transparent "no-fill" segment** (floating bars) — easy, low (waterfall
  covers most cases).
- stacked100: **negative values are silently clamped** — medium, low/niche.

### Waterfall
- **Variance / budget-vs-actual dual bridge** with a hatched gap-to-target
  segment — medium, **high** (finance staple).
- **Blank spacer categories** to group bridge sections — easy, medium.
- Connector re-routing (drag to skip columns) — authoring UX, low for a
  rendering engine.

### Butterfly
- **Value ticks/gridlines on both flanks** (today: only bar-end labels) —
  easy, medium.
- **Stacked flanks** (>1 series per side; today hardcoded to series 0/1) —
  medium, medium.

### Combo
- **Clustered / 100% bars + line** (see top-10 #2) — easy, high.
- **Multiple independent secondary axes** (per-line-series scales) —
  medium, high for KPI dashboards.
- **Waterfall/Mekko + line overlay** (e.g. bridge with %-of-revenue line) —
  medium, medium.
- Stacked-area + line, scatter-on-combo — hard/low; skip for now.

### Line / Area
- **Forecast dash styling** (top-10 #4) — easy, high.
- **Confidence/uncertainty band** around a line (slab technique exists) —
  easy, high for forecasting slides.
- **Area between two series** (plan-vs-actual ribbon) — easy, high.
- **Stepped line/area** — easy, medium. **Smoothed lines** — medium
  (polyline approximation), medium.
- **Area with negative values** — today clamped to 0 — medium, high for
  P&L-over-time.
- **Horizontal "profile chart"** (think-cell parity: rotate line/area) —
  trivial-medium, medium.
- Missing-data bridge option, sparklines — low/niche.

### Scatter / Bubble
- **Size legend** (top-10 #3) — easy, high.
- **Quadrant preset** (top-10 #5) — easy, high.
- **Trajectory mode** (connect points in row order, Gapminder-trail) —
  easy, medium.
- **Continuous color scale** as a third variable (color helpers exist) —
  easy/medium, medium.
- Bubble overlap-repulsion pass, marginal histograms, point icons —
  medium/hard, low/niche.

### Gantt
- **Percent-complete fill** (`% Complete` row → inner bar) — easy, high.
- **Baseline vs. actual ghost bars** (`Baseline start`/`Baseline end` rows) —
  easy, high.
- **Auto summary bars** (section rows span min(start)→max(end) of children) —
  medium, medium.
- **Critical-path highlight** (CPM over existing `After` edges) —
  medium-hard, medium (PM-tool territory; think-cell lacks it too).
- Resource lanes with capacity, working-day calendars, cost columns —
  hard/low; out of deck-tool scope.

### Boxplot
- **Grouped boxplots** (top-10 #9) — easy/medium, high.
- **Jittered raw-data dots** over the box — easy, medium.
- **Notched boxplots** (median CI) — medium, medium.
- Violin (density curves) — infeasible live / medium SVG-only; low.
  Mean±SD box variant — easy, niche.

### Radar
- **Vertex markers in the live add-in** — easy, high (closes the
  add-in-vs-SVG quality gap; already present in scene, verify add-in path).
- **Min–max band across many series** ("peer range + us") — medium
  (SVG/pptx fill; dashed outlines live), high for competitive profiling.
- Per-spoke scales (mixed KPI units) — medium, medium. Radial bar variant —
  really a new kind; medium, medium. Stacked radar — low.

### Pie / Doughnut
- **Bar-of-pie / pie-of-pie breakout** (top-5 + "Other" detailed) —
  medium, high for market-share decks.
- **Semi-circle / gauge-style half doughnut** — easy, medium-high for
  scorecards.
- **Sunburst** (nested rings; wedge fans repeat per ring) — medium-hard,
  medium.
- Variable-radius pie — medium, niche.

### Heatmap
- **Marginal totals** (sum row/column strips) — easy, medium-high.
- **Calendar heatmap layout** (day-of-week × week) — easy-medium, medium.
- Per-cell icon overlays — medium, low (think-cell excludes these too).
  Clustering/dendrograms, cell-size encoding — hard/low; skip.

### Tilemap
- **Mini glyphs per tile** (sparkline/bar inside each region) — medium,
  medium.
- Hex tiles — SVG/pptx-only (no filled polygons live) — medium, low/medium.
- Proportional-area cartograms, drill-down — hard/infeasible; rejected.

## 3. New chart types worth adding (ranked)

1. **Funnel + pyramid** — one declining series; native autoshapes; optional
   auto drop-off % between stages. New kind; easy.
2. **Bullet chart** — as a `Target` row + bands on column charts, not a new
   kind. Easy.
3. **Lollipop / dot / dumbbell-range** — `series.style` marker options on
   clustered. Easy.
4. **Slope chart** — line chart with 2 categories: suppress axis, label both
   ends. Option on line; easy.
5. **Waffle chart** — 10×10 unit grid, part-to-whole for one dominant %.
   New kind; easy.
6. **KPI / number tile** — big number + delta arrow + label; an *Element*
   (with Harvey ball etc.), reusing table-token semantics. Trivial.
7. **Small multiples** — grid-of-charts orchestration over any kind, reusing
   Same Scale + batch insert. Feature; easy-medium.
8. **Bump chart** — rank-over-time via inverted integer axis on line.
   Option; easy-medium.
9. **Treemap** — squarified rect packing (Mekko-adjacent math); blank-row
    grouping for 2 levels. New kind; medium.
10. **Pareto helper** — computed `Cumulative %` row on combo. Trivial.

## 4. Rejected (do not re-propose)

- **Gauge** — bullet chart replaces it deliberately (Few); low deck demand.
- **Sankey / chord / arc** — need curved ribbons; infographic genre.
- **Violin / ridgeline** — smoothed density curves; academic register.
- **Candlestick / OHLC** — trading charts, no consulting demand signal.
- **Stream graph** — feasible but editorial aesthetic; no deck demand.
- **Pictogram with icon libraries** — asset/licensing scope; waffle covers it.
- **Histogram as a kind** — it's clustered with zero gap (see gap-width item).
- **Choropleth maps, 3D, drill-down interactivity** — out of scope by
  design (see CLAUDE.md).
- **Population pyramid, plain dot chart** — already covered (butterfly;
  dot style option).

