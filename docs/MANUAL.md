# PowerChart User Manual

PowerChart makes think-cell-style charts in PowerPoint as **native, editable
shapes** — every bar, label, and line is a real PowerPoint object you can
tweak after insertion. This manual covers the task pane; for driving
PowerChart from JSON or from Claude, see the [automation](#automation-json)
section and `skill/reference.md`.

## Getting started

1. Sideload the add-in: in PowerPoint, **Insert → Add-ins → Upload My
   Add-in** → pick `manifest.xml` (the dev server must be running:
   `npm run dev`). The **PowerChart** ribbon group has an **Insert chart**
   button and menu (direct entries for the most common chart types) plus an
   **Insert element** menu (Harvey ball, checkbox, process flow, KPI tile,
   table) that opens the pane on the Elements tab with that element highlighted.
2. Requirements: PowerPoint on Windows 2207+, Mac 16.62+, or the web
   (PowerPointApi 1.4+). Grouping needs 1.8+, chart re-editing (tags) 1.3+,
   pies 1.10+ (shape rotation), deck-theme colors 1.10+. Marker symbols need
   only 1.4 — they are preset geometry, not rotated shapes. Missing
   capabilities degrade gracefully — charts still insert.
3. No host? The pane also runs as a plain web page (`npm run dev` →
   `localhost:3000`) with everything except insertion, plus a demo gallery.

## Inserting and editing a chart

The pane is organised into four **tabs** — **Chart** (the main flow), **Elements**,
**Agenda**, and **Automation** — and the **Insert / Update** button is pinned to a
bar at the bottom so it's always reachable. Within the Chart tab, *Chart type*,
*Decorations* and *Preview & size* are collapsible sections (click the header);
picking a type auto-collapses the gallery to a compact "current type" summary.

1. **Pick a chart type** from the gallery — the types are grouped by family
   (columns & bars, line & area, parts of a whole, distribution, correlation,
   matrix & spatial), with a search box to filter by name; thumbnails are live
   previews.
2. **Type or paste your data** into the datasheet — row 1 is categories,
   column A is series names. Pasting a range from Excel (Ctrl+V) works.
3. Adjust options; the preview updates live.
4. **Insert into slide.** If a placeholder or shape is selected, the chart
   fills its bounds; repeated inserts cascade.
5. **Re-edit later**: select the chart on the slide — the pane shows an
   "Edit it" banner. Or click **Edit selected chart** to load its data and
   options back into the pane, change anything, and **Update chart** to
   replace it in place. Charts remember their configuration in the file.
6. **Auto-update chart** pushes every pane edit to the slide live (debounced).
7. Ctrl+Z / Ctrl+Y undo and redo pane edits.

## Chart types

| Type | Use for | Notes |
|---|---|---|
| Stacked | composition over categories | one unified engine — a simple column chart is a stacked chart with one series |
| Clustered | side-by-side comparison | bar styles: plain bars, lollipop (stem + dot), dot plot, or a two-series dumbbell range |
| 100% | share of total | `100%=` row lets columns fall short of 100% |
| Waterfall | bridges (EBITDA, headcount) | values are deltas; type `e` in a cell for a computed total; multiple series stack |
| Mekko | two-dimensional composition | width ∝ column total; `X extent` row for explicit widths |
| Line | trends | date categories ("2025-01", "Jan 2025") space proportionally to time; forecast periods draw dashed with hollow markers; `decorations.slope: true` turns two-period data into a slope chart (end rails, "Name value" labels both ends, no axis) |
| Area | stacked trends | |
| Butterfly | two-sided comparison (pyramids) | first two series, back to back, one scale |
| Scatter | correlation | `X`/`Y` rows; `Group` colors points; `X line`/`Y line` partitions; `Trend` fits an OLS line labelled with R² and p, or a **polynomial** curve (`scatter.trendDegree` 2–4, quadratic/cubic/quartic); quadrants shade a 2×2 matrix |
| Bubble | scatter + size | add a `Size` row; a size legend with reference circles appears automatically |
| Gantt | project plans | `Start`/`End`/`Milestone` rows; see [Gantt](#gantt) |
| Combo | columns + line | mark a series with type "line"; secondary axis optional; columns can be stacked, clustered, or 100% |
| Pie / Doughnut | simple shares | slices from 12 o'clock, clockwise; explode slices to highlight; `pie.breakout: [indices]` collapses categories into an "Other" slice detailed in a bar beside the pie (bar-of-pie) |
| Boxplot | distributions | `Min`/`Q1`/`Median`/`Q3`/`Max` rows, or raw samples → automatic Tukey whiskers and outliers; suffix rows with `\| group` for side-by-side grouped boxes |
| Radar | multi-dimension profiles | categories = spokes, series = translucent polygons; keep ≤ 3 series |
| Heatmap | value matrices | series = rows, categories = columns, one global color scale (diverging when data spans zero); marginal row/column totals optional |
| Tile map | values by geography | categories = region codes (US postal / ISO-2 / world regions); layout auto-detected |
| Cascade | stage-by-stage breakdown of a total | categories `"Stage \| Drop label \| Group"`; each bar is a subset of the previous, remainders hang as labeled boxes |
| Funnel | pipeline / conversion stages | centered bands, width ∝ value, conversion % between stages; order values ascending for a pyramid |
| Waffle | one dominant share ("68% of…") | 10×10 unit grid, 1 cell = 1%, filled from the bottom-left; a single category ≤ 100 reads as a literal % with a big-number legend; `100%=` overrides the denominator |
| Treemap | part-to-whole with many items | squarified rectangles, area ∝ value; categories named `"Group \| Item"` nest into two levels (group headers + child tiles) |
| Sunburst | a hierarchy as nested rings | `"Group \| Item"` categories → groups on the inner ring, items on the outer ring |
| Violin | full distribution per category | raw samples per column form a mirrored kernel-density shape with a median tick (fills degrade to outline in the live add-in) |
| Candlestick | financial OHLC / price action | datasheet rows `Open`/`High`/`Low`/`Close`; green when the period rose, red when it fell |

All column-family charts (stacked/clustered/100%), waterfall, Mekko, and
boxplot rotate into **horizontal** bars/rows with the "Horizontal (bar)"
toggle — the axis and every decoration rotate with them.

## The datasheet

- **Row 1** = category headers; **column A** = series names; the corner cell
  is fixed.
- **Navigation**: Enter/↓, ↑, and ←/→ at the text edges move between cells,
  Excel-style. **+ Row / + Column** insert at the cursor; **− Row / −
  Column** delete at the cursor; **⇄ Transpose** swaps rows and columns.
- **Formulas**: cells starting with `=` evaluate A1-style references,
  `+ - * /`, and `SUM`/`AVG`/`MIN`/`MAX` over ranges (`=SUM(B2:E2)`).
- **Blank rows split stacks**: in a stacked chart, series separated by a
  blank row form side-by-side sub-stacks (clustered-stacked).
- **Dates**: ISO (`2026-01-15`), `15.01.2026`, and `Jan 2026` parse as
  calendar dates for Gantt timelines.

### Special rows (magic series names)

| Row name | Effect |
|---|---|
| `100%=` | per-category denominators for 100% charts |
| `X extent` | Mekko-with-units column widths |
| `X`, `Y`, `Size`, `Group` | scatter/bubble coordinates, bubble area, point groups |
| `X line`, `Y line` | scatter partition lines |
| `Trend` | OLS trend line with R² and p-value |
| `Error`, `Error+`, `Error-` | error bars at the column total / line point (symmetric or asymmetric) |
| `Target` | bullet tick across each column; on waterfalls also a hatched gap-to-target segment |
| `Band low`, `Band high` | line charts: shaded confidence/uncertainty ribbon |
| `Min`, `Q1`, `Median`, `Q3`, `Max`, `Mean` | boxplot five-number summary (+ mean marker) |
| `Outlier <n>` | extra boxplot outlier dots |
| `Start`, `End`, `Milestone`, `After`, `Today`, `Holiday`, `Bracket <label>` | Gantt — see below |
| `% Complete`, `Baseline start`, `Baseline end` | Gantt progress fill and plan-vs-actual ghost bars |
| `Column <label>` | Gantt: a numeric table column beside the task labels (cost, FTE, days) |

### Gantt

`Start`/`End` values (numbers or dates) become bars; `Milestone` diamonds;
`After` (1-based row number) draws dependency arrows; `Today` a dashed red
line; `Holiday` shades dates; `Bracket <label>` spans an annotation.
Category cells support `Activity | Owner | Remark` columns, a leading `>`
for indenting, and section-header rows (no Start/End). Calendar timelines
pick weeks (with weekend shading), months, or quarters automatically.

**Waterfall "of which" columns** (`waterfall.detailGroups: [{ of, indices }]`):
break one step of a bridge into its parts, in line with the rest. The detail
columns decompose that step's delta starting from its own base, and sit OFF the
chain — they carry no running total, so every later total and the final one are
unaffected no matter what the parts add up to — while the dashed connector
steps over the whole group to the next real column. Name the parent (`of`); the
engine never guesses it. If the parts don't sum to the parent, you get exactly
the bars you typed — nothing is reconciled behind your back.

**Combo marker series** (`type: "marker"` on a combo series): draws the points
without a connecting line — a benchmark, target or consensus that belongs to
each category rather than to a trend. Keep it on the columns' own scale: on a
`secondaryAxis`, or under `combo.lineAxes: "independent"`, the marks would be
measured against a different axis than the columns they benchmark, which is
the one thing the mark is for. `pareto` keeps only the bars and its cumulative
line, so a marker series does not survive it.

**Overlap relief** (`scatter.spread: "x" | "y"` on scatter/bubble): when markers
cover each other, PowerChart can nudge them along ONE axis you name — the other
stays exact. The shift is hard-capped (`scatter.spreadLimit`, in that axis's
own units; 2% of the range by default) and the cap is printed in the footnote,
so the chart discloses its own approximation. It is ignored under `quadrants`:
a nudge must never move a point across a quadrant line. Both axes carry data,
so PowerChart will not move a marker freely in 2D — that would corrupt two
readings at once.

**Marker symbols** (`scatter.markers` on scatter/bubble): give each `Group` a
point shape — `["circle", "diamond", "triangle"]` makes group 1 circles, group
2 diamonds, group 3 triangles, cycling if there are more groups than shapes.
Shape carries what color cannot: a printed deck is often greyscale, and ~8% of
men cannot separate red from green, either of which flattens a palette back to
one group. The legend draws the shapes, so the channel is always explained —
including alongside a `Color` row, where color means something else. Shapes are
area-matched to the circle they replace, so a bubble's "area tracks Size"
survives: without that a group drawn as stars would read as a quarter the
magnitude of the same value drawn as squares. Available: `circle`, `square`,
`diamond`, `triangle`, `plus` — all PowerPoint preset geometry, so they stay
filled and editable in the deck on any 1.4+ host, and each reproduces its
preset exactly, so the preview and the deck agree on the ink. (A five-point
star and an X are deliberately absent: their presets reshape themselves in
ways the preview cannot reproduce, which would make the preview disagree with
the deck and quietly break the area match.)

**Heatmap sign marks** (`heatmap.symbols: "sign"`): draw a `+` or `−` in each
cell. A diverging heatmap states its direction in hue alone, and hue is what a
greyscale printer takes away: on the default palette the strongest positive
and the strongest negative land about 1.1:1 apart in grey, which is to say
the same tone. A palette chosen for contrasting lightness fares better, but
the two ends of a diverging ramp are the same distance from white by
construction, so the collapse is the rule rather than the exception. Marks are drawn only where the cell has no value label,
since the label already prints the minus sign; that makes them matter most
under `sizeEncode`, which suppresses labels entirely and leaves colour as the
only sign carrier — and on calendar heatmaps, which never draw labels at all.
A cell of exactly zero gets no mark: the scale already paints it the neutral
midpoint, and zero has no sign to state. The option is inert on one-signed
data, where every mark would say the same thing. Not up/down arrows: an arrow on a matrix of KPIs
reads as movement against last period, a claim PowerChart does not check.

**Marginal histograms** (`decorations.marginals: "x" | "y" | "both"` on
scatter/bubble): distribution bars in a gutter along the top and/or right
margin — where the points bunch up, beside the chart that shows how they
relate. The bins subdivide the axis's own ticks, so every tick is a bin edge
and a bar can be read straight against the scale next to it. The gutter takes
real space, so the plot shrinks; on a chart too small to give any away, the
marginals are dropped instead.

Set `gantt.lanes: "owner"` to regroup the tasks under one header per owner
(the middle part of an `Activity | Owner | Remark` category). It is opt-in
because a plan's row order IS the plan — nothing regroups behind your back —
and it groups rather than sorts: inside a lane your rows keep the order you
wrote them, unowned tasks stay together at the end, and `After` dependencies
are renumbered so the arrows still join the same tasks. To control the
grouping yourself, type a section-header row and indent its tasks with `>` —
that has always worked and this changes nothing about it.

**Resource capacity vs load** is a different chart, not a Gantt option: make a
stacked column chart with the weeks as categories and one series per task, and
add a `Target` row for the capacity line. A load histogram needs a value axis,
which a timeline does not have.

Set `gantt.workdays` to make a bar's LENGTH mean working days rather than
elapsed days: a Mon→Mon task reads as 5 units, not 7, and weekends stop
inflating every bar. `true` is Mon–Fri; pass ISO weekday numbers for a custom
week (`[7,1,2,3,4]` = Sun–Thu). `Holiday` rows come out of the scale as well,
and weekend/holiday shading switches off — those days have no width left to
shade. Calendar timelines only.

`Column <label>` adds a numeric table column between the task labels and the
plan — the MS-Project look — right-aligned under `<label>` as its heading. Use
one row per column (`Column Cost €k`, `Column FTE`): each resolves its own
precision from its own values, so money and headcount don't share decimals. Put
the unit in the label. Section-header rows show whatever value you type in
them; nothing is summed for you. Unlike `Remark`, these align and format.

## Options and decorations

The *Format* section organises these into five collapsible groups — **Labels**,
**Axes & scale**, **Analysis**, **Layout**, and **Colours & style** — each with a
count of how many of its options are currently on.

- **Toggles**: Segment labels (auto-hidden when they don't fit),
  Series labels, Column totals, **Grand total** (one label at the top-right of
  a stacked/clustered column chart summing every category total — think-cell 14;
  independent of the per-column totals), Category labels, Value axis, Gridlines,
  **Connector lines** (joins stacked-segment boundaries between columns),
  **100% = note**, **Datamark axis (ticks only)** — Tufte-style tick dashes
  with no axis line.
- **CAGR arrow** and **Difference arrow** between two categories (the
  difference arrow can compare totals, one series' level, or start from a
  value line).
- **Value line**: at the mean (drawn as "Ø") and/or fixed values.
- **Segment order**: sheet / reversed / ascending / descending.
- **Axis scale** min/max pinning and **Axis break** (compresses an
  out-of-scale range, clearly marked).
- **Labels**: decimals / suffix ("€m") / locale (1.234,5 for de-DE), and
  **label content** — any combination of value, %, series, category.
- **Footnote / source** — every good chart cites its source; renders
  bottom-left.
- **Explode slices** (pie/doughnut).
- **Stepped line/area** (`stepped: "before"|"after"|"center"`): draws a
  staircase instead of sloped segments — good for rates that hold then jump.
- **Area with negatives**: area charts now stack positives above the zero
  line and negatives below it, so P&L-style series that go negative dip
  under the baseline instead of being clamped to zero.
- **Scatter trajectory** (`trajectory`): connects scatter/bubble points in
  row order with a direction trail (Gapminder-style path over time).
- **Continuous color** (`Color` row): colors scatter/bubble points on a
  sequential ramp by a numeric value, with a gradient legend — a third
  variable beyond X/Y (and Size).
- **Smoothed lines** (`smooth`): draws line charts as smooth curves instead
  of straight segments.
- **Waterfall spacers** (`spacerIndices`): blank grouping gaps that split a
  long bridge into sections; the running total and connector bridge across.
- **Gantt summary bars** (`summaryBars`): a capped summary bar on each
  section-header row spanning its child activities' min start → max end.
- **Notched boxplots** (`notch`): pinch each box at the median confidence
  interval — non-overlapping notches flag a significant median difference.
- **Radar band** (`radarBand`): shade the peer min–max envelope and draw the
  last series on top — a "peer range + us" competitive profile.
- **Other bucket** (`otherBucket`): collapse a long tail of stack segments
  into one "Other" series, keeping the largest few.
- **Calendar heatmap** (`heatmap.calendar`): lay a daily date series out as a
  weekday × week grid with month labels (contributions view).
- **Butterfly stacked flanks** (`butterfly.split`): stack more than one
  series on each side of a butterfly chart.
- **Radar per-spoke scales** (`radar.perSpoke`): normalize each spoke to its
  own maximum so mixed-unit KPIs compare by shape.
- **Missing-data bridge** (`bridgeGaps`): connect a line straight across
  missing (null) points instead of breaking it.
- **Floating segments**: a stacked-column series with `color: "transparent"`
  reserves its space without drawing, floating the bars above it.
- **Combo base modes** (`combo.columns`): put the lines over a waterfall
  bridge or a Mekko as well as stacked / clustered / 100% columns.
- **Combo independent line axes** (`combo.lineAxes: "independent"`): give each
  line series its own scale so unlike-unit KPIs share one chart.
- **Hex tile maps** (`tilemap.shape: "hex"`): draw the cartogram with
  hexagonal tiles instead of squares.
- **Tilemap mini-glyphs** (`tilemap.glyph: "bars"`): a mini bar chart per
  region from a multi-series datasheet.
- **100% charts with negatives**: `stacked100` now shows negative segments
  below the zero line instead of clamping them away.
- **Semi-circle gauge** (`doughnut` + `pie.semi`): a half-doughnut scorecard.
- **Pareto** (`pareto`): sorted bars + a cumulative-% line (80/20 view).
- **Bump chart** (`decorations.bump`): rank-over-time on an inverted axis.
- **Horizontal profile chart** (`horizontal` on line/area): rotates the chart
  so categories run down the left and values extend to the right.
- **Critical path** (`decorations.criticalPath` on Gantt): red-outlines the
  longest chain of `After`-dependent tasks — the tasks that drive the finish
  date — and draws its dependency arrows thicker in red.
- **Mean±SD box** (`boxplot.meanSd`): draws each box as mean ± 1 SD with the
  centre line at the mean and whiskers to mean ± 2 SD (raw-sample mode) — the
  scientific summary rather than the quartile one.
- **Sparklines** (`decorations.sparkline` on line/area): compact, axis-less
  word-sized trend lines with min/max/last dots; pair with small multiples for
  a KPI dashboard table, one row per series.
- **Radial bar chart** (`radar.bars`): a circular bar chart / Nightingale rose —
  each category is an equal-angle wedge whose radius encodes its value; multiple
  series stack outward.
- **Stacked radar** (`radar.stacked`): series stack cumulatively along each
  spoke as nested filled bands (part-to-whole across dimensions).
- **Variable-radius pie** (`pie.variableRadius`, or a `Radius` datasheet row):
  slice angle encodes the first series while radius encodes a second metric.
- **Cell-size heatmap** (`heatmap.sizeEncode`): each cell's area encodes the
  value's magnitude (colour still encodes the signed value) — the corrplot view
  for correlation / signed matrices.
- **Heatmap clustering** (`heatmap.cluster`): reorder rows by average-linkage
  similarity so similar rows sit together, with a dendrogram in a left gutter.
- **Combo area base** (`combo.columns: "area"`): put the line series over a
  stacked-area base instead of columns — trend-of-mix plus a KPI line.
- **Boxplot jitter** (`jitter`): overlays the raw observations as jittered
  dots over each box (raw-sample mode).
- **Gap width & overlap** (`gapWidth`, `overlap`): Excel's two column-spacing
  controls. Gap width (0–500%) sets how wide the columns are — 0 makes them
  touch for a histogram look; overlap (−100…100) sets how much clustered bars
  within a category overlap or gap.
- **Same scale** (deck-wide or selection) re-renders charts on a common axis.
- **Download SVG / Download PNG** (overflow menu) save the preview as a vector
  or a 2× raster image — handy for email or chat, where the native-shape output
  can't go. The real deliverable is still **Insert into slide** (editable shapes).
- **Copy chart link** (overflow menu) copies a URL that reopens the exact chart
  on the hosted gallery — the config rides in the link's `#c=` hash (never a
  server log). Good for support or sharing a config without a file.
- **Small multiples** (`multiples: {columns?}`, JSON): splits a multi-series
  chart into a grid of single-series panels titled by series name, on one
  shared value scale (column family, line/area, waterfall, radar).
- **IBCS scenario notation** (`series.scenario`, JSON, column family): encodes
  the data's nature in the fill so actual vs plan vs forecast reads without a
  legend — **AC** solid, **PY** lighter, **PL** / **BU** outlined/hollow, **FC**
  hatched (solid in the live add-in). The two-letter code is appended to the
  legend/label; give each scenario series the same `color` for the single-hue
  IBCS look.
- **IBCS variance tier** (`decorations.variance`, JSON, column family): a strip
  below the columns showing an `actual` series' deviation from a `reference`
  (plan / previous year) per category, as signed bars from a zero line — absolute
  Δ (default) or relative `mode:"percent"`. `goodIsUp` (default true) colours the
  sign green (favourable) / red (not); set it false for cost-like metrics.
- Advanced (JSON-only, survive re-editing): callouts (speech bubbles),
  background bands, per-cell highlight colors, pattern fills, log scale
  presets — see `skill/reference.md` for the full schema.

## Colors and style

- **Palette** presets, per-series color pickers, and **Use deck theme** —
  reads the presentation's Accent 1–6 theme colors (PowerPointApi 1.10).
- **Export style / Import style**: a corporate style file (JSON: palette,
  font, negative/total colors) applied to every chart you make.
- **Templates** — the dropdown groups built-in **Starters** (polished, ready-to-
  edit presets: revenue bridge, growth columns + CAGR, channel mix, market
  share, KPI trend, bullet vs target, programme Gantt) with **My templates**.
  **Save as template** stores the whole current chart (data + options) locally
  under *My templates*; **Delete** removes a saved one (starters stay).

## Elements

- **Harvey ball** — fraction-filled circle, size slider.
- **Checkbox** — ✓ / ✗ / − status marks.
- **Process flow** — chevron step bar with a highlighted step.
- **KPI tile** — big number + delta arrow + caption; the arrow takes
  semantic colors from the delta's sign, and **↓ good** flips them for
  metrics where falling is better (costs, churn).
- **Table from datasheet** — rule-based table (lines top/header/bottom only,
  never side borders, a gap every 5 rows, optional bold **Total row**).
  Cells accept effect tokens: `[hb:0.75]` mini harvey ball, `[up]` `[down]`
  `[flat]` trend arrows, `[good]` `[bad]` semantic colors.
- **Agenda** — inserts one chapter slide per line, each highlighting its own
  chapter.

## Automation (JSON)

**Export current** copies the chart's `ChartConfig` JSON; **Import** loads
one; **Insert batch** inserts an array of configs at once. The same JSON
drives the CLI (`npm run render`) and the **Claude Agent Skill**
(`skill-dist/powerchart-charts.zip` — upload at claude.ai → Settings →
Capabilities → Skills, then ask Claude for "a waterfall of …" on any
surface, including Claude for PowerPoint).

### Testing

**Insert demo deck** appends a fresh slide for every chart kind plus a set of
feature and element highlights — a fast way to eyeball how everything renders
as native shapes in your PowerPoint. Each chart stays editable (click it and
**Edit selected chart**); delete the slides when you're done. The deck opens
with a title slide stamping the build and host, and closes with a results slide
summarising what rendered, what the host skipped or failed, and how long the run
took — so an exported PDF is a self-contained regression record. Tick **Smoke
test (10 slides)** first for a fast pass over one representative chart per family
instead of the full deck. A slide that stalls is retried once automatically.

## Excel companion

Sideload `manifest-excel.xml` in Excel, select a range, and **Generate**
turns it into PowerChart JSON to paste into the pane's Automation box — the
practical substitute for live data links.

## Troubleshooting

- **Charts insert ungrouped** — host below PowerPointApi 1.8; everything
  still works, shapes just aren't grouped.
- **Pie slices missing** — host below 1.10 (no shape rotation).
- **"Theme unavailable"** — host below 1.10.
- **Chart not recognized for editing** — it must carry the PowerChart tag;
  charts inserted before tagging existed can't be re-opened.
- **German UI** — the pane localizes automatically when Office reports German.
