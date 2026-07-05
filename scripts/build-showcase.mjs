#!/usr/bin/env node
/**
 * Build the feature showcase deck: one slide per chart kind and per signature
 * feature, committed at examples/showcase.json + examples/showcase.pptx.
 *
 *   npm run showcase          (regenerate after changing features)
 *
 * CI regenerates the deck and fails when the committed copy is stale, and
 * test/showcase.test.ts fails when a chart kind or feature has no slide —
 * so the deck always demonstrates the full feature set.
 */
import { writeFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { sampleConfig, CHART_KINDS } from "../dist-lib/powerchart.js";

const day = (iso) => Math.round(Date.parse(iso) / 86400000);

/** Every chart kind, via its curated sample. */
const kinds = CHART_KINDS.map(({ kind, label }) => ({
  ...sampleConfig(kind),
  title: `${label} — ${sampleConfig(kind).title ?? "sample"}`,
}));

/** Feature slides: everything a sample doesn't already show. */
const features = [
  {
    ...sampleConfig("stacked"),
    title: "Decorations — totals, CAGR, difference, value line",
    decorations: {
      totals: true,
      cagr: { from: 0, to: 3 },
      difference: { from: 0, to: 3 },
      valueLines: [{ mode: "mean" }],
    },
  },
  {
    ...sampleConfig("stacked"),
    title: "Level difference (series) + label content Value+%",
    decorations: {
      difference: { from: 0, to: 3, series: 0 },
      labelContent: ["value", "percent"],
    },
  },
  {
    ...sampleConfig("stacked"),
    title: "Bar orientation + segment order descending",
    horizontal: true,
    segmentOrder: "descending",
    decorations: { totals: true, seriesLabels: true },
  },
  {
    ...sampleConfig("stacked"),
    title: "Category sort descending + pinned axis + de-DE locale",
    categorySort: "descending",
    scale: { max: 150 },
    numberFormat: { decimals: 1, locale: "de-DE" },
    decorations: { totals: true, valueAxis: true, gridlines: true },
  },
  {
    kind: "stacked",
    width: 480,
    height: 300,
    title: "Clustered-stacked (stack groups) ",
    data: {
      categories: ["2024", "2025"],
      series: [
        { name: "EU Product", values: [30, 36], stack: 0 },
        { name: "EU Services", values: [12, 16], stack: 0 },
        { name: "US Product", values: [26, 33], stack: 1 },
        { name: "US Services", values: [10, 14], stack: 1 },
      ],
    },
    decorations: { totals: true },
  },
  {
    kind: "clustered",
    width: 480,
    height: 300,
    title: "Axis break (620 outlier) + axis title",
    data: { categories: ["Q1", "Q2", "Q3", "Q4"], series: [{ name: "Sales", values: [42, 55, 620, 61] }] },
    axisBreak: { from: 80, to: 580 },
    valueAxisTitle: "€m",
    decorations: { valueAxis: true, gridlines: true, seriesLabels: false },
  },
  {
    kind: "clustered",
    width: 480,
    height: 300,
    title: "Log scale",
    data: { categories: ["Seed", "A", "B", "C"], series: [{ name: "Valuation", values: [4, 40, 220, 1900] }] },
    logScale: true,
    valueAxisTitle: "$m (log)",
    decorations: { valueAxis: true, gridlines: true, seriesLabels: false },
  },
  {
    kind: "stacked100",
    width: 480,
    height: 300,
    title: "100%= row (columns short of 100%)",
    data: {
      categories: ["2023", "2024", "2025"],
      series: [
        { name: "Won", values: [40, 55, 70] },
        { name: "Lost", values: [30, 25, 20] },
      ],
      hundredPercent: [100, 100, 100],
    },
  },
  {
    ...sampleConfig("waterfall"),
    title: "Stacked waterfall + value-line-anchored difference",
    data: {
      categories: ["FY23", "Organic", "M&A", "Cost", "FY24"],
      series: [
        { name: "Europe", values: [50, 8, 5, -6, 0] },
        { name: "Americas", values: [36, 6, 9, -6, 0] },
      ],
    },
    waterfall: { totalIndices: [4] },
    decorations: {
      categoryAxis: true,
      valueLines: [{ mode: "value", value: 100 }],
      difference: { from: 0, to: 4, fromValueLine: 0, percent: false },
    },
  },
  { ...sampleConfig("waterfall"), title: "Rotated waterfall", horizontal: true },
  {
    ...sampleConfig("mekko"),
    title: "Mekko with units (X extent) — rotated",
    horizontal: true,
    data: { ...sampleConfig("mekko").data, xExtent: [50, 30, 20] },
  },
  { ...sampleConfig("combo"), title: "Combo + secondary axis", secondaryAxis: true },
  {
    kind: "scatter",
    width: 480,
    height: 300,
    title: "Scatter — partition lines, trend, groups",
    data: {
      categories: ["Alpha", "Bravo", "Core", "Delta", "Echo", "Foxtrot"],
      series: [
        { name: "X", values: [12, 25, 40, 55, 62, 74] },
        { name: "Y", values: [18, 30, 38, 52, 55, 68] },
        { name: "Group", values: [1, 1, 2, 2, 3, 3] },
        { name: "X line", values: [45, null, null, null, null, null] },
        { name: "Y line", values: [40, null, null, null, null, null] },
        { name: "Trend", values: [1, null, null, null, null, null] },
      ],
    },
  },
  {
    kind: "gantt",
    width: 480,
    height: 300,
    title: "Gantt — sections, owners, remarks, deps, today, holiday, bracket",
    data: {
      categories: [
        "Phase 1 — Discovery",
        "> Interviews | Anna | 12 done",
        "> Synthesis | Ben",
        "Phase 2 — Delivery",
        "> Build | Cato | at risk",
      ],
      series: [
        { name: "Start", values: [null, day("2026-01-05"), day("2026-01-19"), null, day("2026-01-26")] },
        { name: "End", values: [null, day("2026-01-23"), day("2026-01-30"), null, day("2026-02-20")] },
        { name: "Milestone", values: [null, null, day("2026-01-30"), null, day("2026-02-20")] },
        { name: "After", values: [null, null, 2, null, 3] },
        { name: "Today", values: [day("2026-02-02"), null, null, null, null] },
        { name: "Holiday", values: [day("2026-01-15"), null, null, null, null] },
        { name: "Bracket Sprint 1", values: [day("2026-01-05"), day("2026-01-19"), null, null, null] },
        { name: "% Complete", values: [null, 100, 60, null, 20] },
        { name: "Baseline start", values: [null, day("2026-01-05"), day("2026-01-15"), null, day("2026-01-22")] },
        { name: "Baseline end", values: [null, day("2026-01-20"), day("2026-01-27"), null, day("2026-02-12")] },
      ],
      dates: true,
    },
  },
  {
    ...sampleConfig("stacked"),
    title: "Manual label nudge (labelOffsets)",
    labelOffsets: { "total-3": { dx: 0, dy: -6 } },
    decorations: { totals: true },
  },
  // --- "The good chart" batch: design formalia & data-highlighting ---
  {
    ...sampleConfig("stacked"),
    title: "Connector lines + footnote (source line)",
    footnote: "Source: Statistics Denmark, 2024",
    decorations: { connectors: true, totals: true },
  },
  {
    kind: "stacked",
    width: 480,
    height: 300,
    title: "Point highlight + speech-bubble callout",
    data: {
      categories: ["2021", "2022", "2023", "2024"],
      series: [
        { name: "Base", values: [40, 44, 47, 52] },
        { name: "New", values: [8, 10, 21, 14], colors: [null, null, "#e34948", null] },
      ],
    },
    decorations: {
      connectors: true,
      callouts: [{ text: "One-off boost", category: 2, dx: 30 }],
    },
  },
  {
    kind: "clustered",
    width: 480,
    height: 300,
    title: "Background band (target range) + line highlight",
    data: { categories: ["Q1", "Q2", "Q3", "Q4"], series: [{ name: "Margin", values: [38, 45, 52, 49] }] },
    decorations: {
      bands: [{ axis: "y", from: 40, to: 50, label: "Target range" }],
      seriesLabels: false,
      valueAxis: true,
    },
  },
  {
    kind: "scatter",
    width: 480,
    height: 300,
    title: "Scatter background bands (quadrant regions)",
    data: {
      categories: ["A", "B", "C", "D", "E"],
      series: [
        { name: "X", values: [15, 30, 45, 60, 75] },
        { name: "Y", values: [20, 45, 35, 60, 70] },
      ],
    },
    decorations: { bands: [{ axis: "x", from: 50, to: 80, color: "#e8f0e4", label: "Focus" }] },
  },
  {
    kind: "pie",
    width: 480,
    height: 300,
    title: "Exploding slice + slice color + 100% = note",
    footnote: "Source: annual report FY24",
    pie: { explode: [1] },
    data: {
      categories: ["Retail", "Online", "Wholesale", "Other"],
      series: [{ name: "Revenue", values: [46, 28, 18, 8], colors: [null, "#e34948", null, null] }],
    },
    decorations: { hundredPercentNote: true },
  },
  {
    kind: "stacked",
    width: 480,
    height: 300,
    title: "Pattern fills (SVG/preview; solid in PPT)",
    data: {
      categories: ["2022", "2023", "2024"],
      series: [
        { name: "Base", values: [40, 44, 50] },
        { name: "Plan", values: [10, 14, 16], pattern: "diagonal" },
        { name: "Risk", values: [6, 8, 9], pattern: "dots" },
      ],
    },
    decorations: { connectors: true },
  },
  {
    kind: "boxplot",
    width: 480,
    height: 300,
    title: "Boxplot from raw samples (Tukey whiskers, outliers)",
    boxplot: { showMean: true },
    data: {
      categories: ["North", "South", "East"],
      series: [
        { name: "", values: [4, 7, 3] }, { name: "", values: [5, 8, 4] },
        { name: "", values: [5, 9, 4] }, { name: "", values: [6, 9, 5] },
        { name: "", values: [6, 10, 5] }, { name: "", values: [7, 11, 6] },
        { name: "", values: [7, 11, 6] }, { name: "", values: [8, 12, 7] },
        { name: "", values: [15, 12, 14] },
      ],
    },
    decorations: { categoryAxis: true, valueAxis: true },
  },
  {
    ...sampleConfig("boxplot"),
    title: "Horizontal boxplot (same shared axis, rotated)",
    horizontal: true,
  },
  {
    kind: "boxplot",
    width: 480,
    height: 300,
    title: "Grouped boxplots (2024 vs 2025)",
    data: {
      categories: ["North", "South", "East"],
      series: [
        { name: "Min | 2024", values: [2, 3, 1] },
        { name: "Q1 | 2024", values: [3, 5, 2] },
        { name: "Median | 2024", values: [4, 7, 3] },
        { name: "Q3 | 2024", values: [6, 9, 5] },
        { name: "Max | 2024", values: [8, 12, 7] },
        { name: "Min | 2025", values: [3, 4, 2] },
        { name: "Q1 | 2025", values: [4, 6, 3] },
        { name: "Median | 2025", values: [5, 8, 4] },
        { name: "Q3 | 2025", values: [7, 10, 6] },
        { name: "Max | 2025", values: [9, 13, 8] },
      ],
    },
    decorations: { valueAxis: true },
  },
  {
    kind: "clustered",
    width: 480,
    height: 300,
    title: "Error bars (Error / Error+ / Error- rows)",
    data: {
      categories: ["Trial 1", "Trial 2", "Trial 3", "Trial 4"],
      series: [
        { name: "Result", values: [42, 55, 48, 61] },
        { name: "Error", values: [5, 7, null, null] },
        { name: "Error+", values: [null, null, 9, 4] },
        { name: "Error-", values: [null, null, 3, 8] },
      ],
    },
    decorations: { valueAxis: true, seriesLabels: false },
  },
  {
    kind: "clustered",
    width: 480,
    height: 300,
    title: "Bullet chart (Target row + range bands)",
    data: {
      categories: ["North", "South", "East", "West"],
      series: [
        { name: "Actual", values: [42, 55, 48, 61] },
        { name: "Target", values: [50, 50, 55, 58] },
      ],
    },
    decorations: {
      bands: [
        { axis: "y", from: 0, to: 35, color: "#e4e2da" },
        { axis: "y", from: 35, to: 55, color: "#efeee8" },
      ],
      seriesLabels: false,
    },
  },
  {
    kind: "combo",
    width: 480,
    height: 300,
    title: "Combo: clustered columns + line",
    combo: { columns: "clustered" },
    data: {
      categories: ["2023", "2024", "2025"],
      series: [
        { name: "Actual", values: [40, 46, 52] },
        { name: "Plan", values: [38, 48, 55] },
        { name: "Margin %", values: [31, 34, 38], type: "line" },
      ],
    },
    secondaryAxis: true,
  },
  {
    kind: "clustered",
    width: 480,
    height: 300,
    title: "Dumbbell range (before vs after)",
    data: {
      categories: ["North", "South", "East", "West"],
      series: [
        { name: "2024", values: [40, 55, 48, 61] },
        { name: "2025", values: [52, 60, 45, 70] },
      ],
    },
    decorations: { barStyle: "range", seriesLabels: true },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Forecast with confidence band",
    data: {
      categories: ["2023", "2024", "2025", "2026", "2027"],
      series: [
        { name: "Revenue", values: [40, 46, 52, 60, 69] },
        { name: "Band low", values: [null, null, 52, 55, 59] },
        { name: "Band high", values: [null, null, 52, 66, 82] },
      ],
    },
    decorations: { forecastFrom: 3, seriesLabels: false },
  },
  {
    kind: "waterfall",
    width: 480,
    height: 300,
    title: "Budget-vs-actual bridge (gap to target)",
    data: {
      categories: ["FY23", "Volume", "Price", "Cost", "FY24"],
      series: [
        { name: "Delta", values: [86, 14, 9, -12, 0] },
        { name: "Target", values: [null, null, null, null, 110] },
      ],
    },
    waterfall: { totalIndices: [4] },
    decorations: { categoryAxis: true },
  },
  {
    kind: "heatmap",
    width: 480,
    height: 300,
    title: "Heatmap with marginal totals",
    heatmap: { totals: "both" },
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [
        { name: "North", values: [42, 48, 51, 58] },
        { name: "South", values: [30, 34, 31, 36] },
        { name: "East", values: [22, 25, 29, 33] },
      ],
    },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Slope chart — market share, before vs after",
    data: {
      categories: ["2020", "2025"],
      series: [
        { name: "Brand A", values: [24, 31] },
        { name: "Brand B", values: [28, 26] },
        { name: "Brand C", values: [19, 22] },
        { name: "Private label", values: [12, 18] },
      ],
    },
    decorations: { slope: true },
    footnote: "Source: internal market model, 2026",
  },
  {
    kind: "waffle",
    width: 480,
    height: 300,
    title: "Waffle — order mix by channel",
    data: {
      categories: ["Online", "Retail", "Wholesale"],
      series: [{ name: "Orders", values: [52, 33, 15] }],
    },
  },
  {
    kind: "pie",
    width: 480,
    height: 300,
    title: "Bar-of-pie — revenue with the long tail detailed",
    data: {
      categories: ["EMEA", "Americas", "APAC", "Nordics", "Benelux", "DACH"],
      series: [{ name: "Revenue", values: [80, 100, 60, 20, 15, 25] }],
    },
    pie: { breakout: [3, 4, 5] },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Small multiples — one panel per brand, shared scale",
    multiples: {},
    data: {
      categories: ["2022", "2023", "2024", "2025"],
      series: [
        { name: "Brand A", values: [40, 48, 55, 66] },
        { name: "Brand B", values: [30, 32, 31, 35] },
        { name: "Brand C", values: [12, 18, 26, 41] },
      ],
    },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true },
  },
  {
    kind: "bubble",
    width: 480,
    height: 300,
    title: "Bubble size legend + quadrant matrix",
    data: {
      categories: ["Alpha", "Bravo", "Core", "Delta", "Echo"],
      series: [
        { name: "X", values: [20, 35, 55, 70, 85] },
        { name: "Y", values: [25, 60, 40, 75, 55] },
        { name: "Size", values: [30, 80, 45, 100, 60] },
      ],
    },
    decorations: { quadrants: { x: 50, y: 50, labels: ["Question marks", "Stars", "Dogs", "Cash cows"] } },
  },
  {
    ...sampleConfig("line"),
    title: "Datamark axis (Tufte range frame)",
    decorations: { segmentLabels: false, valueAxis: "datamarks", tickMode: "data", gridlines: false },
  },
  {
    kind: "heatmap",
    width: 480,
    height: 300,
    title: "Diverging heatmap (YoY change, %)",
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [
        { name: "North", values: [4, 6, -2, 8] },
        { name: "South", values: [-5, -3, 2, 1] },
        { name: "East", values: [1, 3, 5, 9] },
      ],
    },
  },
  {
    kind: "tilemap",
    width: 480,
    height: 300,
    title: "US tile map (auto-detected layout)",
    data: {
      categories: ["CA", "TX", "NY", "FL", "IL", "WA", "MA", "CO", "GA", "NC"],
      series: [{ name: "Stores", values: [120, 95, 88, 74, 51, 44, 38, 29, 27, 22] }],
    },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Highlighted max point (per-cell color on line)",
    data: {
      categories: ["Jan", "Feb", "Mar", "Apr", "May"],
      series: [{ name: "Orders", values: [120, 135, 178, 150, 162], colors: [null, null, "#e34948", null, null] }],
    },
    decorations: { seriesLabels: false },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Stepped line (staircase)",
    data: {
      categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
      series: [{ name: "Base rate", values: [2.0, 2.0, 2.5, 2.5, 3.0, 2.75] }],
    },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true, stepped: "after" },
  },
  {
    kind: "area",
    width: 480,
    height: 300,
    title: "Stepped area",
    data: {
      categories: ["2021", "2022", "2023", "2024", "2025"],
      series: [
        { name: "Mobile", values: [20, 32, 45, 60, 74] },
        { name: "Desktop", values: [35, 37, 36, 34, 30] },
      ],
    },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true, stepped: "before" },
  },
  {
    ...sampleConfig("butterfly"),
    title: "Butterfly with value ticks on both flanks",
    decorations: { segmentLabels: true, valueAxis: true, gridlines: true },
  },
  {
    ...sampleConfig("clustered"),
    title: "Clustered overlap (Excel-style, 45%)",
    overlap: 45,
    decorations: { seriesLabels: true },
  },
  {
    ...sampleConfig("stacked"),
    title: "Narrow gap width (histogram look, gapWidth 0)",
    gapWidth: 0,
    decorations: { totals: true },
  },
  {
    kind: "area",
    width: 480,
    height: 300,
    title: "Area with negative values (net income over time)",
    data: {
      categories: ["2020", "2021", "2022", "2023", "2024", "2025"],
      series: [{ name: "Net income", values: [12, -8, -15, 4, 18, 26] }],
    },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true },
  },
  {
    kind: "scatter",
    width: 480,
    height: 300,
    title: "Trajectory trail (one entity over time)",
    data: {
      categories: ["2019", "2020", "2021", "2022", "2023", "2024"],
      series: [
        { name: "X", values: [20, 28, 33, 45, 52, 60] },
        { name: "Y", values: [30, 26, 40, 38, 55, 68] },
      ],
    },
    decorations: { trajectory: true },
  },
  {
    kind: "boxplot",
    width: 480,
    height: 300,
    title: "Boxplot with jittered raw-data dots",
    boxplot: { jitter: true },
    data: {
      categories: ["North", "South", "East"],
      series: [
        { name: "s1", values: [4, 6, 3] },
        { name: "s2", values: [5, 7, 4] },
        { name: "s3", values: [6, 7, 5] },
        { name: "s4", values: [5, 9, 4] },
        { name: "s5", values: [7, 8, 6] },
        { name: "s6", values: [6, 11, 5] },
        { name: "s7", values: [8, 7, 7] },
        { name: "s8", values: [5, 10, 4] },
        { name: "s9", values: [9, 8, 6] },
        { name: "s10", values: [6, 12, 5] },
      ],
    },
    decorations: { categoryAxis: true, valueAxis: true },
  },
];

const configs = [...kinds, ...features];
writeFileSync("examples/showcase.json", JSON.stringify(configs, null, 2) + "\n");
execFileSync("node", ["skill/scripts/render-pptx.mjs", "examples/showcase.json", "examples/showcase.pptx"], {
  stdio: "inherit",
});
console.log(`examples/showcase.json + examples/showcase.pptx (${configs.length} slides)`);
