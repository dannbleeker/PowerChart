# ChartConfig reference

Everything the PowerChart engine accepts. All lengths in points (1pt = 1/72").

```ts
{
  kind: "stacked" | "clustered" | "stacked100" | "waterfall" | "mekko"
      | "line" | "area" | "butterfly" | "scatter" | "bubble" | "gantt"
      | "combo" | "pie" | "doughnut" | "boxplot" | "radar" | "heatmap"
      | "tilemap" | "cascade" | "funnel" | "waffle",
  width?: 480, height?: 300,          // frame size in pt
  title?: string,
  horizontal?: boolean,               // rotate columns/waterfall/mekko/boxplot into bars/rows
  data: {
    categories: string[],             // column headers (x categories / activities / points)
    series: [{
      name: string,
      values: (number | null)[],      // one per category; null = blank
      color?: "#rrggbb",              // per-series override
      colors?: ("#rrggbb"|null)[],    // per-CELL override: highlight one segment/point/slice
      pattern?: "diagonal"|"crosshatch"|"dots"|"horizontal",
                                      // hatch over the fill (SVG/preview; solid in PPT output)
      type?: "line",                  // combo: draw this series as a line
      stack?: number                  // clustered-stacked: stack group index
    }],
    hundredPercent?: (number|null)[], // "100%=" row: denominators for 100% charts
    xExtent?: (number|null)[],        // "X extent" row: Mekko-with-units widths
    dates?: true                      // set when values are days-since-epoch (Gantt)
  },
  decorations?: {
    segmentLabels?: boolean,          // default true, auto-hidden when too small
    seriesLabels?: boolean,           // default true (right of last column / legend)
    totals?: boolean,                 // column totals
    categoryAxis?: boolean,           // default true
    valueAxis?: boolean | "datamarks", gridlines?: boolean,  // default false;
                                      // "datamarks" = Tufte ticks-only axis, no line
    tickMode?: "nice" | "data",       // datamarks at round values or at data min/max
    gridShape?: "polygon" | "circle", // radar grid web (default polygon)
    fillOpacity?: number,             // radar series fill (default 0.18)
    labelContent?: ("value"|"percent"|"series"|"category")[],
    cagr?: { from: number, to: number, series?: number },        // category indices
    difference?: { from, to, percent?, series?, fromValueLine? },
    valueLines?: ({mode:"mean"} | {mode:"value", value:number})[],
    connectors?: boolean,             // lines joining stacked-segment boundaries between columns
    callouts?: [{ text, category, series?, dx?, dy? }],  // speech-bubble comments on a value
    bands?: [{ axis:"x"|"y", from, to, color?, label? }], // shaded background region
                                      // (y = value range; x = category indices;
                                      //  scatter/bubble: both axes in value units)
    hundredPercentNote?: boolean,     // "100% = N" note (pie/doughnut/stacked100)
    forecastFrom?: number,            // line: dashed segments + hollow markers from this index
    barStyle?: "bar"|"lollipop"|"dot"|"range",  // clustered: stems+dots / dots / dumbbell
    fillBetween?: [number, number],   // line: shade the gap between two series (plan vs actual)
    stepped?: "before"|"after"|"center",  // line/area: staircase segments (jump at start / end / midpoint)
    smooth?: boolean,                 // line: smooth Catmull-Rom curves (sampled polyline)
    slope?: boolean,                  // line: slope-chart mode — end rails + "Name value" labels, no axis
    trajectory?: boolean,             // scatter/bubble: connect points in row order with a direction trail
    summaryBars?: boolean,            // gantt: summary bar on section rows (spans children min→max)
    radarBand?: boolean,              // radar: shade the peer min–max envelope, draw last series on top
    quadrants?: { x, y, labels? }     // scatter: 4 tinted zones + corner labels at one crossing
  },
  footnote?: string,                  // source line, bottom-left ("Source: …, 2024")
  pie?: { explode?: number[],         // slice indices offset radially to highlight
          breakout?: number[] },      // pie only: collapse these categories into one "Other"
                                      // slice detailed in a stacked bar beside the pie (bar-of-pie)
  multiples?: { columns?: number },   // small multiples: one single-series panel per series,
                                      // shared value scale (stacked/clustered/line/area/waterfall/radar)
  boxplot?: { whiskers?: "tukey"|"minmax", quartileMethod?: "exclusive"|"inclusive",
              showMean?: boolean, iqrMultiplier?: number,
              jitter?: boolean,       // overlay raw observations as jittered dots (raw-sample mode)
              notch?: boolean },      // notch each box at the median CI (raw-sample mode)
  map?: "us" | "eu" | "europe" | "world",   // tilemap layout (auto-detected if omitted)
  heatmap?: { color?, negativeColor?, mode?: "sequential"|"diverging"|"auto",
              totals?: "row"|"column"|"both" },  // marginal sum strips
  combo?: { columns?: "stacked"|"clustered"|"stacked100" },  // column mode under the lines
  waterfall?: { totalIndices?: number[],    // categories drawn as running totals ("e")
                spacerIndices?: number[] }, // blank grouping gaps (empty category name)
  scale?: { min?: number, max?: number },   // pin the value axis
  axisBreak?: { from: number, to: number }, // compress an out-of-scale range
  logScale?: boolean,                       // clustered/line, positive data
  gapWidth?: number,                        // column family: Excel gap width 0–500 (% of column width; default 50)
  overlap?: number,                         // clustered: Excel bar overlap −100…100 (default 0 = edge to edge)
  valueAxisTitle?: string,                  // units label, e.g. "€m"
  segmentOrder?: "sheet"|"reverse"|"ascending"|"descending",
  categorySort?: "ascending"|"descending",  // by column total (not waterfall/gantt)
  secondaryAxis?: boolean,                  // combo: right-hand axis for line series
  labelOffsets?: { [nodeName]: {dx, dy} },  // manual label nudges
  numberFormat?: { decimals?: number|"auto", suffix?: string,
                   forceSign?: boolean, locale?: "de-DE"|... },
  style?: { palette?: string[], fontFamily?, fontSize?, negative?, neutral? }
}
```

## Special series-name conventions (the "datasheet rows")

| Row name | Effect |
|---|---|
| `100%=` | per-category denominator for `stacked100` (columns can stay short of 100%) |
| `X extent` | Mekko-with-units: explicit column widths |
| `X`, `Y`, `Size`, `Group` | scatter/bubble coordinates, bubble area, point coloring (1..k) |
| `Color` | scatter/bubble: numeric third variable on a sequential color ramp (with a gradient legend); supersedes `Group` coloring |
| `X line`, `Y line` | scatter partition lines at those values |
| `Trend` | any value present → OLS trend line, labelled with its R² and p-value |
| `Start`, `End`, `Milestone` | Gantt bars & milestone markers (numbers or day values) |
| `% Complete` | Gantt: darker progress fill inside the bar (0-100 or 0-1) |
| `Baseline start`, `Baseline end` | Gantt: thin ghost bars showing the original plan vs actual |
| `After` | Gantt dependency: 1-based predecessor index → elbow arrow |
| `Today` | Gantt today line at the (single) value |
| `Holiday` | Gantt: shade these dates |
| `Bracket <label>` | Gantt: interval annotation spanning min→max of the row's values |
| `Error`, `Error+`, `Error-` | error bars at the column total / line point (± or asymmetric); stacked/clustered/line/area, vertical |
| `Target` | bullet tick across each column; on waterfalls also a hatched gap-to-target segment + label |
| `Band low`, `Band high` | line charts: shaded confidence/uncertainty ribbon (rows never draw as lines) |
| `Min`, `Q1`, `Median`, `Q3`, `Max` | boxplot five-number summary (whiskers to Min/Max, think-cell style); suffix "\| group" ("Min \| 2024") for side-by-side grouped boxes |
| `Mean` | boxplot mean marker (×) |
| `Outlier <n>` | boxplot: extra outlier dots in precomputed mode |

Gantt category conventions: `"Activity | Owner | Remark"` adds responsible and
remark columns; a leading `">"` indents; a category with no Start/End/Milestone
renders as a bold section-header band.

**Dates**: for Gantt rows, pass days-since-epoch numbers and set `data.dates:
true` (compute as `Math.round(Date.parse("2026-01-15")/86400000)`), or let the
interactive datasheet parse ISO strings. Calendar timelines pick weeks (with
weekend shading), months, or quarters by span automatically.

**Waterfall semantics**: values are deltas; `waterfall.totalIndices` lists
categories drawn from the baseline to the running total (think-cell's `e`).
Multiple series → stacked waterfall (contributions stack per column).

**Clustered-stacked**: give series `stack: 0`, `stack: 1`, … — same index
stacks together; different indices sit side by side per category.

**Boxplot without summary rows**: when no `Min`/`Q1`/… rows are present,
every series row is a raw observation — quartiles are computed (exclusive
method) and whiskers use Tukey 1.5×IQR fences with outliers drawn as dots.

**Funnel**: one series of stage values (order ascending for a pyramid);
centered bands with width ∝ value, conversion % vs the previous stage in
the gaps, stage names on the left.

**Waffle**: a 10×10 unit grid (1 cell = 1%), filling from the bottom-left.
Categories are the parts (pie semantics, first series); shares round to
whole cells by largest remainder and unassigned cells stay gray. The
denominator is the `100%=` row when present, else the value sum — except a
single category ≤ 100 with no denominator, which reads as a literal
percentage (`68` → 68 cells + a big "68%" beside the grid).

**Stepped line/area** (`decorations.stepped`): draw values as a staircase
instead of sloped segments — `"after"` holds each value then jumps at the
next category (HV), `"before"` jumps immediately (VH), `"center"` steps at
the midpoint (HVH). Works for `line` and `area`; area fills follow the steps.

**Area with negative values**: `area` charts stack positives above the zero
baseline and negatives below it, so a series that goes negative (e.g. net
income over time) dips under the axis instead of being clamped to 0.

**Scatter/bubble trajectory** (`decorations.trajectory`): connects the points
in datasheet (row) order with a trail and a direction arrowhead at each
segment midpoint — a Gapminder-style path of one entity through the X/Y
space over time. Order the categories chronologically.

**Scatter/bubble continuous color** (a `Color` row): maps each point's numeric
value onto a sequential color ramp and swaps the group chip legend for a
gradient legend with min/max labels. Encodes a third variable (or a fourth,
on bubbles alongside `Size`). Supersedes `Group` coloring when both exist.

**Smoothed lines** (`decorations.smooth`): draws smooth Catmull-Rom curves
through the points instead of straight segments (sampled to a dense polyline
since the renderers have no freeform paths). Ignored when `stepped` is set.

**Waterfall spacers** (`waterfall.spacerIndices`): blank categories that draw
no bar and leave a gap, grouping a long bridge into sections. The running
total carries across and the dashed connector bridges the gap — give the
spacer category an empty name so no axis label shows.

**Boxplot jitter** (`boxplot.jitter`): in raw-sample mode, overlays every
observation as a deterministically jittered dot over its box, so the reader
sees the distribution and sample size behind the summary (outlier dots are
subsumed by the jittered points).

**Boxplot notch** (`boxplot.notch`): in raw-sample mode, pinches each box at
the median ± 1.57·IQR/√n confidence interval — boxes whose notches don't
overlap have significantly different medians. Precomputed boxes (no sample
size) stay rectangular.

**Gantt summary bars** (`decorations.summaryBars`): draws a capped summary bar
on each section-header row (a category with no Start/End/Milestone), spanning
min(start)→max(end) of the activities beneath it up to the next header.

**Radar min–max band** (`decorations.radarBand`): shades the per-spoke min–max
envelope of the peer series (all series except the last) as a band and draws
the last series prominently on top — the "peer range + us" competitive
profile. The legend collapses the peers into one "Peer range" swatch.

**Gap width & overlap** (`gapWidth`, `overlap`): mirror Excel's two spacing
controls for the column family. `gapWidth` (0–500, default 50) is the gap
between columns as a percentage of column width — 0 makes them touch
(histogram look), higher makes them thinner. `overlap` (−100…100, default 0)
sets how much clustered bars within a category overlap: positive overlaps
them, negative opens a gap.

**Slope chart**: `kind: "line"` + `decorations.slope: true` with (ideally)
two categories — the before/after comparison. No value axis; vertical
rails carry the two periods and every series gets a colored "Name value"
label at both ends (labels de-overlap automatically).

**Bar-of-pie** (`pie.breakout`): the listed category indices collapse into
one muted "Other" slice (rotated to face 3 o'clock) and are detailed in a
stacked bar on the right with dashed connectors; detail labels show the
share of the grand total.

**Small multiples** (`multiples`): a multi-series chart becomes a grid of
single-series panels titled by series name, pinned to one shared value
scale so panels compare honestly. Special rows (Error, Target, Band
low/high) are carried into every panel. `columns` overrides the grid
(default: one row up to 3 panels, else a near-square).

**Cascade** (decomposition): each stage's bar is a subset of the previous
one, top-aligned on one volume scale; the complement hangs as a muted
labeled box at each split. Category syntax: `"Stage | Drop label | Group
header"` — part 2 captions the remainder of the split into this stage,
consecutive stages sharing part 3 get one spanning header band.
series[0] = kept counts per stage.

**Radar**: categories = spokes (first at 12 o'clock, clockwise), series =
translucent polygons; keep 5–10 spokes and ≤3 series. **Heatmap**: series =
rows, categories = columns; one global color scale (diverging through white
when the data spans zero). **Tilemap**: categories = region codes (US postal,
ISO-2 for eu/europe, or NA/LATAM/EU/MEA/SSA/CIS/SA/EA/SEA/OCE for world) —
the layout is auto-detected from the codes or pinned with `map:`.

## Chart-design formalia (baked-in defaults, per "the good chart" practice)

Every chart should carry: a message title, a `footnote` citing source and
period, sorted categories where order is free, labels on the data instead of
a value axis, and — on percentage charts — `hundredPercentNote` so readers
know what 100% is. Use `decorations.connectors` on stacked charts to make
segment development followable, `series.colors` to draw the eye to one data
point, `callouts` to comment on it, and `bands` to mark a reference region.

## Palette

Default 8-slot categorical palette (CVD-validated): `#2a78d6 #1baf7a #eda100
#008300 #4a3aa7 #e34948 #e87ba4 #eb6834`. Waterfall decreases use `#e34948`,
totals `#898781`. Override via `style.palette` or per-series `color`.

## Output

`scripts/render-pptx.mjs` emits one 13.33×7.5in slide per config with the chart
centered, every bar/label/line as a native PowerPoint shape, and exact
adjustable pie geometry. `scripts/render-svg.mjs` emits SVGs for fast visual QA.
