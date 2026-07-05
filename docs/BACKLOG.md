# Backlog

Curated candidate work, from a research sweep (July 2026) comparing every
existing chart kind against think-cell, Excel, Highcharts, Datawrapper, and
Mekko Graphics, plus a chart-type survey across the Zelazny / FT Visual
Vocabulary taxonomies and competitor add-ins (Zebra BI, Vizzlo, UpSlide).
Every claim was verified against the source before listing — items the
engine already has were filtered out.

**This is the only backlog document.** Items graduate from here into PRs
(and then out of here); rejected ideas stay in §4 so they aren't re-proposed. (Shipped since the
sweep: the full top-10 — cascade, bullet targets, combo column modes,
bubble size legend, forecast styling, scatter quadrants, funnel,
lollipop/dot/dumbbell styles, gantt progress + baselines, grouped
boxplots — plus, from the next tier: waterfall budget-vs-actual bridge,
line confidence bands, area-between-series (`fillBetween`), heatmap
marginal totals, slope chart, waffle, KPI tile, bar-of-pie breakout,
small multiples; plus batch G of the §2 within-kind gaps: stepped
line/area, Excel-style column gap width + clustered overlap, and butterfly
value ticks/gridlines on both flanks; plus batch H: area with negative
values, scatter/bubble trajectory trail, and boxplot jittered raw-data
dots; plus batch I: scatter/bubble continuous color scale, smoothed lines,
and waterfall grouping spacers; plus batch J: gantt auto-summary bars,
notched boxplots, and the radar min–max "peer range + us" band; plus
batch K: automatic "Other" bucket, calendar heatmap layout, and butterfly
stacked flanks; plus batch L: radar per-spoke scales, transparent floating
column segments, and the line missing-data bridge; plus batch M: combo
waterfall/Mekko base modes and independent per-line-series axes; plus
batch N: tilemap hex tiles, tilemap mini-glyphs, and stacked100 negatives.)
Feasibility is judged against the live-add-in constraint (rects, lines,
text, ellipses, triangles, polygon *outlines* — no freeform curves; the SVG
and skill-pptx renderers additionally have filled polygons and patterns).

## 1. Top candidates (value ÷ effort, cross-cutting rank)

The researched top-10 and the full next tier are shipped (waterfall
bridge, confidence bands, fill-between, heatmap totals, slope, waffle,
KPI tile, bar-of-pie, small multiples). Next candidates come from §2/§3
below (strongest remaining: semi-circle doughnut, bump chart, treemap,
Pareto helper).

## 2. Gaps within existing kinds (per kind)

### Column family (stacked / clustered / stacked100)
- (all researched column-family gaps shipped)

### Waterfall
- Connector re-routing (drag to skip columns) — authoring UX, low for a
  rendering engine.

### Combo
- Stacked-area + line, scatter-on-combo — hard/low; skip for now.

### Line / Area
- **Horizontal "profile chart"** (think-cell parity: rotate line/area) —
  trivial-medium, medium.
- Sparklines — low/niche.

### Scatter / Bubble
- Bubble overlap-repulsion pass, marginal histograms, point icons —
  medium/hard, low/niche.

### Gantt
- **Critical-path highlight** (CPM over existing `After` edges) —
  medium-hard, medium (PM-tool territory; think-cell lacks it too).
- Resource lanes with capacity, working-day calendars, cost columns —
  hard/low; out of deck-tool scope.

### Boxplot
- Violin (density curves) — infeasible live / medium SVG-only; low.
  Mean±SD box variant — easy, niche.

### Radar
- **Vertex markers in the live add-in** — easy, high (closes the
  add-in-vs-SVG quality gap; already present in scene, verify add-in path).
- Radial bar variant — really a new kind; medium, medium. Stacked radar — low.

### Pie / Doughnut
- **Semi-circle / gauge-style half doughnut** — easy, medium-high for
  scorecards.
- **Sunburst** (nested rings; wedge fans repeat per ring) — medium-hard,
  medium.
- Variable-radius pie — medium, niche.

### Heatmap
- Per-cell icon overlays — medium, low (think-cell excludes these too).
  Clustering/dendrograms, cell-size encoding — hard/low; skip.

### Tilemap
- Proportional-area cartograms, drill-down — hard/infeasible; rejected.

## 3. New chart types worth adding (ranked)

2. **Bullet chart** — as a `Target` row + bands on column charts, not a new
   kind. Easy.
3. **Lollipop / dot / dumbbell-range** — `series.style` marker options on
   clustered. Easy.
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

