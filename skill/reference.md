# ChartConfig reference

Everything the PowerChart engine accepts. All lengths in points (1pt = 1/72").

```ts
{
  kind: "stacked" | "clustered" | "stacked100" | "waterfall" | "mekko"
      | "line" | "area" | "butterfly" | "scatter" | "bubble" | "gantt"
      | "combo" | "pie" | "doughnut",
  width?: 480, height?: 300,          // frame size in pt
  title?: string,
  horizontal?: boolean,               // rotate columns/waterfall/mekko into bars/rows
  data: {
    categories: string[],             // column headers (x categories / activities / points)
    series: [{
      name: string,
      values: (number | null)[],      // one per category; null = blank
      color?: "#rrggbb",              // per-series override
      colors?: ("#rrggbb"|null)[],    // per-CELL override: highlight one segment/point/slice
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
    valueAxis?: boolean, gridlines?: boolean,   // default false (think-cell style)
    labelContent?: ("value"|"percent"|"series"|"category")[],
    cagr?: { from: number, to: number, series?: number },        // category indices
    difference?: { from, to, percent?, series?, fromValueLine? },
    valueLines?: ({mode:"mean"} | {mode:"value", value:number})[],
    connectors?: boolean,             // lines joining stacked-segment boundaries between columns
    callouts?: [{ text, category, series?, dx?, dy? }],  // speech-bubble comments on a value
    bands?: [{ axis:"x"|"y", from, to, color?, label? }], // shaded background region
                                      // (y = value range; x = category indices;
                                      //  scatter/bubble: both axes in value units)
    hundredPercentNote?: boolean      // "100% = N" note (pie/doughnut/stacked100)
  },
  footnote?: string,                  // source line, bottom-left ("Source: …, 2024")
  pie?: { explode?: number[] },       // slice indices offset radially to highlight
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
| `Trend` | any value present → OLS trend line over all points |
| `Start`, `End`, `Milestone` | Gantt bars & milestone markers (numbers or day values) |
| `After` | Gantt dependency: 1-based predecessor index → elbow arrow |
| `Today` | Gantt today line at the (single) value |
| `Holiday` | Gantt: shade these dates |
| `Bracket <label>` | Gantt: interval annotation spanning min→max of the row's values |

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
