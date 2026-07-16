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
ellipses, triangles and polygon *outlines* — no freeform curves, and no images
(the scene graph has no image node). The SVG and skill-pptx renderers
additionally have filled polygons and patterns.

The researched top-10 and the whole next tier have shipped, as has every new
chart type the sweep ranked and every within-kind gap it found worth doing.
Nothing below is queued behind them: each item is parked for the reason it
states.

## 1. Open

### Blocked on an image node the engine doesn't have

Both need the same two new things — an image/icon node kind that all three
renderers can draw, and an icon asset library (licensing scope). Pictogram (§2)
was rejected on exactly that; these are the same call under another name.

- **Scatter/bubble point icons** — medium/hard, low/niche.
- **Heatmap per-cell icon overlays** — medium, low.

### Authoring UX rather than rendering

PowerChart renders from a config; this wants in-canvas manipulation, which is
one of the three things Office.js can't reach anyway (see docs/RESEARCH.md).

- **Waterfall connector re-routing** (drag to skip columns) — low for a
  rendering engine.

### Judged low value

- **Scatter-on-combo** — hard/low. (`Series.type` is `"column" | "line"`, and
  the combo bases are column/waterfall/Mekko/area.)
- **Bubble overlap-repulsion pass** — medium/hard, low/niche. Label collision
  avoidance already ships; this is about separating the bubbles themselves.
- **Marginal histograms** on scatter/bubble — low/niche.
- **Gantt resource lanes with capacity, working-day calendars, cost columns** —
  hard/low; project-management scope, not deck-tool scope.

## 2. Rejected or already covered (do not re-propose)

- **Dial / needle gauge** — bullet charts replace it deliberately (Few); low
  deck demand. Note the *semi-circle scorecard* gauge did ship (`doughnut` +
  `pie.semi`); what stays rejected is the dial-with-needle and its threshold
  bands.
- **Sankey / chord / arc** — need curved ribbons; infographic genre.
- **Ridgeline** — stacked density curves; academic register. (The single-column
  `violin` kind shipped at owner request; ridgeline stacking not pursued.)
- **Stream graph** — feasible but editorial aesthetic; no deck demand.
- **Pictogram with icon libraries** — needs the image node above plus an asset
  library (licensing scope). Waffle is the deliberate substitute: it covers the
  part-to-whole genre with square cells, it does not render icons.
- **Histogram as a kind** — the look is `clustered` + `gapWidth: 0`, both
  shipped. If this is ever revisited, auto-binning raw samples into categories
  is the only real gap; the bar geometry is not.
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
