# PowerChart User Manual

PowerChart makes think-cell-style charts in PowerPoint as **native, editable
shapes** — every bar, label, and line is a real PowerPoint object you can
tweak after insertion. This manual covers the task pane; for driving
PowerChart from JSON or from Claude, see the [automation](#automation-json)
section and `skill/reference.md`.

## Getting started

1. Sideload the add-in: in PowerPoint, **Insert → Add-ins → Upload My
   Add-in** → pick `manifest.xml` (the dev server must be running:
   `npm run dev`). The **PowerChart** ribbon menu opens the pane, with
   direct entries for the most common chart types.
2. Requirements: PowerPoint on Windows 2207+, Mac 16.62+, or the web
   (PowerPointApi 1.4+). Grouping needs 1.8+, chart re-editing (tags) 1.3+,
   pies 1.9+, deck-theme colors 1.10+. Missing capabilities degrade
   gracefully — charts still insert.
3. No host? The pane also runs as a plain web page (`npm run dev` →
   `localhost:3000`) with everything except insertion, plus a demo gallery.

## Inserting and editing a chart

1. **Pick a chart type** from the gallery (thumbnails are live previews).
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
| Scatter | correlation | `X`/`Y` rows; `Group` colors points; `X line`/`Y line` partitions; `Trend` fits an OLS line labelled with R² and p; quadrants shade a 2×2 matrix |
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

### Gantt

`Start`/`End` values (numbers or dates) become bars; `Milestone` diamonds;
`After` (1-based row number) draws dependency arrows; `Today` a dashed red
line; `Holiday` shades dates; `Bracket <label>` spans an annotation.
Category cells support `Activity | Owner | Remark` columns, a leading `>`
for indenting, and section-header rows (no Start/End). Calendar timelines
pick weeks (with weekend shading), months, or quarters automatically.

## Options and decorations

- **Toggles**: Segment labels (auto-hidden when they don't fit),
  Series labels, Column totals, Category labels, Value axis, Gridlines,
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
- **Boxplot jitter** (`jitter`): overlays the raw observations as jittered
  dots over each box (raw-sample mode).
- **Gap width & overlap** (`gapWidth`, `overlap`): Excel's two column-spacing
  controls. Gap width (0–500%) sets how wide the columns are — 0 makes them
  touch for a histogram look; overlap (−100…100) sets how much clustered bars
  within a category overlap or gap.
- **Same scale** (deck-wide or selection) re-renders charts on a common axis.
- **Small multiples** (`multiples: {columns?}`, JSON): splits a multi-series
  chart into a grid of single-series panels titled by series name, on one
  shared value scale (column family, line/area, waterfall, radar).
- Advanced (JSON-only, survive re-editing): callouts (speech bubbles),
  background bands, per-cell highlight colors, pattern fills, log scale
  presets — see `skill/reference.md` for the full schema.

## Colors and style

- **Palette** presets, per-series color pickers, and **Use deck theme** —
  reads the presentation's Accent 1–6 theme colors (PowerPointApi 1.10).
- **Export style / Import style**: a corporate style file (JSON: palette,
  font, negative/total colors) applied to every chart you make.
- **Save as template** stores the whole chart (data + options) locally for
  reuse.

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

## Excel companion

Sideload `manifest-excel.xml` in Excel, select a range, and **Generate**
turns it into PowerChart JSON to paste into the pane's Automation box — the
practical substitute for live data links.

## Troubleshooting

- **Charts insert ungrouped** — host below PowerPointApi 1.8; everything
  still works, shapes just aren't grouped.
- **Pie slices missing** — host below 1.9 (no shape rotation).
- **"Theme unavailable"** — host below 1.10.
- **Chart not recognized for editing** — it must carry the PowerChart tag;
  charts inserted before tagging existed can't be re-opened.
- **German UI** — the pane localizes automatically when Office reports German.
