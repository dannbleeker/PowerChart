# ChartConfig reference

Everything the PowerChart engine accepts. All lengths in points (1pt = 1/72").

```ts
{
  kind: "stacked" | "clustered" | "stacked100" | "waterfall" | "mekko"
      | "line" | "area" | "butterfly" | "scatter" | "bubble" | "gantt"
      | "combo" | "pie" | "doughnut" | "boxplot" | "radar" | "heatmap"
      | "tilemap" | "cascade" | "funnel",
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
    quadrants?: { x, y, labels? }     // scatter: 4 tinted zones + corner labels at one crossing
  },
  footnote?: string,                  // source line, bottom-left ("Source: …, 2024")
  pie?: { explode?: number[] },       // slice indices offset radially to highlight
  boxplot?: { whiskers?: "tukey"|"minmax", quartileMethod?: "exclusive"|"inclusive",
              showMean?: boolean, iqrMultiplier?: number },
  map?: "us" | "eu" | "europe" | "world",   // tilemap layout (auto-detected if omitted)
  heatmap?: { color?, negativeColor?, mode?: "sequential"|"diverging"|"auto" },
  combo?: { columns?: "stacked"|"clustered"|"stacked100" },  // column mode under the lines
  waterfall?: { totalIndices?: number[] },  // categories drawn as running totals ("e")
  scale?: { min?: number, max?: number },   // pin the value axis
  axisBreak?: { from: number, to: number }, // compress an out-of-scale range
  logScale?: boolean,                       // clustered/line, positive data
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
| `X line`, `Y line` | scatter partition lines at those values |
| `Trend` | any value present → OLS trend line, labelled with its R² and p-value |
| `Start`, `End`, `Milestone` | Gantt bars & milestone markers (numbers or day values) |
| `After` | Gantt dependency: 1-based predecessor index → elbow arrow |
| `Today` | Gantt today line at the (single) value |
| `Holiday` | Gantt: shade these dates |
| `Bracket <label>` | Gantt: interval annotation spanning min→max of the row's values |
| `Error`, `Error+`, `Error-` | error bars at the column total / line point (± or asymmetric); stacked/clustered/line/area, vertical |
| `Target` | bullet-chart tick across each column at the value (combine with `bands` for zones) |
| `Min`, `Q1`, `Median`, `Q3`, `Max` | boxplot five-number summary (whiskers to Min/Max, think-cell style) |
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
