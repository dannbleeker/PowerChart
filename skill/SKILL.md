---
name: powerchart-charts
description: Create think-cell-style consulting charts as native, editable PowerPoint shapes — waterfalls (EBITDA bridges), Mekko/Marimekko, stacked/clustered columns, Gantt plans, combo, pie, scatter/bubble, butterfly — plus agenda slides, from a JSON chart config. Use when the user asks for professional/consulting/think-cell-style charts, a chart in PowerPoint (.pptx), a waterfall or bridge chart, a Mekko, a Gantt/project plan slide, or CAGR/difference-arrow annotations.
---

# PowerChart: think-cell-style charts for PowerPoint

Turn a JSON **ChartConfig** into a `.pptx` slide made of **native, individually
editable PowerPoint shapes** (never pictures) — with think-cell's signature
decorations: CAGR arrows, difference arrows, value lines, automatic totals, and
collision-avoiding labels.

## Workflow

1. **Author a ChartConfig** (JSON) from the user's data. Full schema and all
   datasheet-row conventions: read `reference.md`. Minimal example:

```json
{
  "kind": "waterfall",
  "title": "EBITDA bridge FY24",
  "data": {
    "categories": ["FY23", "Volume", "Price", "Cost", "FX", "FY24"],
    "series": [{ "name": "Delta", "values": [86, 14, 9, -12, -4, 0] }]
  },
  "waterfall": { "totalIndices": [5] },
  "decorations": { "difference": { "from": 0, "to": 5 } }
}
```

2. **Render.** One-time setup: `npm install pptxgenjs` (in this skill's folder).

```bash
node scripts/render-pptx.mjs charts.json out.pptx   # native shapes, 1 chart/slide
node scripts/render-svg.mjs charts.json out/        # quick SVG previews
```

`charts.json` may hold one config or an array (one slide each).

3. **QA visually.** Render the SVGs and inspect them (or convert the pptx to
   images) before delivering: check label overlaps, totals, axis sanity. Fix
   the config, not the output.

## Choosing the chart kind

| User asks for | kind |
|---|---|
| Bridge / walk / build-up / EBITDA waterfall | `waterfall` (`"e"` semantics via `waterfall.totalIndices`; several series → stacked waterfall) |
| Market map with variable column widths | `mekko` (add an `X extent` row for explicit widths) |
| Share-of-total over time | `stacked100` (optionally a `100%=` row) |
| Revenue by segment / stacked bars | `stacked` (`horizontal: true` for bars; `stack` per series for clustered-stacked) |
| Project / timeline slide | `gantt` (`Start`/`End`/`Milestone`/`After`/`Today`/`Holiday`/`Bracket …` rows; ISO dates supported) |
| KPI line over columns | `combo` (`"type": "line"` on a series; `secondaryAxis: true` for %-on-right) |
| Positioning / portfolio map | `scatter` or `bubble` (`X`/`Y`/`Size`/`Group` rows, auto label placement) |
| Population pyramid / two-sided compare | `butterfly` |
| Simple share | `pie` / `doughnut` (exact wedge geometry in the pptx output) |
| Side-by-side comparison per category | `clustered` |
| Distribution per category | `boxplot` (Min/Q1/Median/Q3/Max rows, or raw samples → Tukey whiskers) |
| Multi-dimension profile / score | `radar` (translucent polygons, ≤3 series) |
| Matrix of values (region × period) | `heatmap` (one global color scale) |
| Values by geography | `tilemap` (tile-grid cartogram: `map: "us"/"eu"/"europe"/"world"`) |
| Stage-by-stage breakdown of a total (contacts → answered → solved) | `cascade` ("Stage | Drop label | Group" categories) |
| Pipeline / conversion stages | `funnel` (centered bands + conversion %) |
| One dominant share ("68% of…") | `waffle` (10×10 unit grid; single category reads as a literal %) |
| Share with a long tail ("top 3 + the rest in detail") | `pie` + `pie.breakout: [indices]` (bar-of-pie) |
| Same chart repeated per series ("one panel per region") | any column/line/area/waterfall/radar kind + `multiples: {}` (shared scale) |
| Before/after comparison of a few series | `line` + `decorations.slope: true` (slope chart: end rails, labels both ends) |
| Trend over time | `line` (date categories space proportionally) or `stacked100`-style share via `area` |
| Rates that hold then jump (policy rate, tiers) | `line`/`area` + `decorations.stepped: "after"` (staircase) |

Decorations (`decorations` object): `totals`, `cagr {from,to}`,
`difference {from,to,series?,fromValueLine?}`, `valueLines`, `labelContent`,
`segmentOrder`, `connectors` (segment-boundary lines between stacked columns),
`callouts` (speech-bubble comments), `bands` (shaded background regions),
`hundredPercentNote`, `stepped` (line/area staircase), plus `scale {min,max}`,
`axisBreak {from,to}`, `logScale`, `gapWidth`/`overlap` (Excel-style column
spacing), `categorySort`, `footnote` (source line), `pie {explode}` and
per-cell `series.colors` highlights at the top level. Defaults are
think-cell-like: labels on, axis off; always set `footnote` with the data
source when you know it.

## Rules

- Emit **valid JSON only** for configs; validate values are numbers (dates as
  ISO strings only in Gantt rows).
- Keep one message of data → one chart. For decks, pass an array.
- The output shapes are grouped per chart is NOT done here (pptx groups are
  flat) — that's fine: think-cell users expect to tweak individual shapes.
- Prefer `waterfall` over generic columns whenever the story is a change
  decomposition. Prefer `mekko` when both share and size matter.
