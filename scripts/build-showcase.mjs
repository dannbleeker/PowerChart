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
    ...sampleConfig("clustered"),
    title: "Grand total — sum of all category totals (top-right)",
    decorations: { totals: true, grandTotal: true, categoryAxis: true },
  },
  {
    kind: "clustered",
    width: 560,
    height: 320,
    title: "IBCS scenarios — PY / AC / FC / PL by fill",
    decorations: { categoryAxis: true, seriesLabels: false, valueAxis: true },
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [
        { name: "Sales", scenario: "PY", color: "#3b6ea5", values: [80, 85, 90, 95] },
        { name: "Sales", scenario: "AC", color: "#3b6ea5", values: [82, 88, 96, null] },
        { name: "Sales", scenario: "FC", color: "#3b6ea5", values: [null, null, null, 102] },
        { name: "Sales", scenario: "PL", color: "#3b6ea5", values: [85, 90, 95, 100] },
      ],
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
        { name: "", values: [4, 7, 3] },
        { name: "", values: [5, 8, 4] },
        { name: "", values: [5, 9, 4] },
        { name: "", values: [6, 9, 5] },
        { name: "", values: [6, 10, 5] },
        { name: "", values: [7, 11, 6] },
        { name: "", values: [7, 11, 6] },
        { name: "", values: [8, 12, 7] },
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
  {
    kind: "scatter",
    width: 480,
    height: 300,
    title: "Continuous color scale (a Color row)",
    data: {
      categories: ["Alpha", "Bravo", "Core", "Delta", "Echo", "Foxtrot", "Golf", "Hotel"],
      series: [
        { name: "X", values: [12, 25, 40, 55, 62, 74, 30, 48] },
        { name: "Y", values: [30, 55, 42, 70, 35, 60, 22, 50] },
        { name: "Color", values: [5, 12, 18, 22, 28, 34, 9, 16] },
      ],
    },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Smoothed line (Catmull-Rom)",
    data: {
      categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul"],
      series: [
        { name: "A", values: [20, 45, 30, 55, 40, 62, 48] },
        { name: "B", values: [15, 22, 28, 26, 35, 40, 52] },
      ],
    },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true, smooth: true },
  },
  {
    kind: "waterfall",
    width: 480,
    height: 300,
    title: "Waterfall with grouping spacers",
    data: {
      categories: ["FY23", "Volume", "Price", "", "Cost", "FX", "", "FY24"],
      series: [{ name: "Delta", values: [86, 14, 9, null, -12, -4, null, 0] }],
    },
    waterfall: { totalIndices: [7], spacerIndices: [3, 6] },
    decorations: { categoryAxis: true },
  },
  {
    kind: "gantt",
    width: 480,
    height: 300,
    title: "Gantt with auto-summary bars",
    data: {
      categories: [
        "Phase 1: Discovery",
        "> Research",
        "> Interviews",
        "Phase 2: Build",
        "> Backend",
        "> Frontend",
        "> QA",
      ],
      series: [
        { name: "Start", values: [null, 1, 2, null, 5, 7, 11] },
        { name: "End", values: [null, 4, 5, null, 10, 12, 13] },
      ],
    },
    decorations: { summaryBars: true, segmentLabels: false },
  },
  {
    kind: "boxplot",
    width: 480,
    height: 300,
    title: "Notched boxplots (median CI)",
    boxplot: { notch: true },
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "o1", values: [10, 14, 20] },
        { name: "o2", values: [12, 15, 22] },
        { name: "o3", values: [13, 16, 23] },
        { name: "o4", values: [14, 17, 25] },
        { name: "o5", values: [15, 18, 26] },
        { name: "o6", values: [16, 19, 28] },
        { name: "o7", values: [17, 20, 30] },
        { name: "o8", values: [18, 22, 33] },
        { name: "o9", values: [20, 24, 36] },
        { name: "o10", values: [22, 27, 40] },
      ],
    },
    decorations: { categoryAxis: true, valueAxis: true },
  },
  {
    kind: "radar",
    width: 480,
    height: 300,
    title: "Radar peer range + us",
    scale: { min: 0, max: 5 },
    data: {
      categories: ["Strategy", "Data", "Talent", "Process", "Tech", "Culture"],
      series: [
        { name: "Peer 1", values: [3, 2, 4, 3, 3, 2] },
        { name: "Peer 2", values: [2, 3, 3, 4, 2, 3] },
        { name: "Peer 3", values: [4, 4, 2, 3, 4, 3] },
        { name: "Us", values: [3, 4, 4, 2, 5, 4] },
      ],
    },
    decorations: { seriesLabels: true, radarBand: true },
  },
  {
    kind: "stacked",
    width: 480,
    height: 300,
    title: "Long tail collapsed into Other (otherBucket)",
    otherBucket: { max: 5 },
    data: {
      categories: ["2023", "2024", "2025"],
      series: [
        { name: "Enterprise", values: [40, 46, 52] },
        { name: "SMB", values: [22, 26, 31] },
        { name: "Consumer", values: [18, 17, 19] },
        { name: "Gov", values: [6, 7, 8] },
        { name: "Edu", values: [4, 5, 5] },
        { name: "NGO", values: [3, 3, 4] },
        { name: "Startup", values: [2, 3, 3] },
      ],
    },
    decorations: { totals: true },
  },
  {
    kind: "heatmap",
    width: 480,
    height: 300,
    title: "Calendar heatmap (daily commits)",
    heatmap: { calendar: true },
    data: {
      categories: Array.from({ length: 30 }, (_, i) => {
        const d = new Date(Date.UTC(2025, 0, 6 + i));
        return d.toISOString().slice(0, 10);
      }),
      series: [
        {
          name: "Commits",
          values: [3, 5, 2, 8, 4, 0, 0, 6, 1, 9, 7, 3, 0, 0, 5, 8, 2, 4, 6, 0, 0, 10, 3, 5, 7, 2, 0, 0, 4, 6],
        },
      ],
    },
  },
  {
    kind: "butterfly",
    width: 480,
    height: 300,
    title: "Butterfly with stacked flanks",
    butterfly: { split: 2 },
    data: {
      categories: ["<30", "30–39", "40–49", "50–59", "60+"],
      series: [
        { name: "Women FT", values: [220, 340, 300, 180, 60] },
        { name: "Women PT", values: [200, 270, 240, 140, 60] },
        { name: "Men FT", values: [240, 420, 400, 280, 110] },
        { name: "Men PT", values: [140, 230, 200, 130, 50] },
      ],
    },
    decorations: { segmentLabels: true },
  },
  {
    kind: "radar",
    width: 480,
    height: 300,
    title: "Radar with per-spoke scales (mixed KPIs)",
    radar: { perSpoke: true },
    data: {
      categories: ["Revenue $m", "NPS", "Churn %", "Headcount", "Uptime %"],
      series: [
        { name: "2024", values: [120, 42, 8, 340, 99.2] },
        { name: "2025", values: [155, 55, 5, 410, 99.7] },
      ],
    },
    decorations: { seriesLabels: true },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Missing-data bridge (bridgeGaps)",
    data: {
      categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
      series: [{ name: "Signups", values: [20, null, 35, null, 30, 48] }],
    },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true, bridgeGaps: true },
  },
  {
    kind: "stacked",
    width: 480,
    height: 300,
    title: "Floating range columns (transparent base)",
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [
        { name: "Floor", color: "transparent", values: [10, 14, 12, 18] },
        { name: "Range", values: [20, 18, 24, 16] },
      ],
    },
    decorations: { segmentLabels: true },
  },
  {
    kind: "combo",
    width: 480,
    height: 300,
    title: "Combo: waterfall bridge + %-of-total line",
    secondaryAxis: true,
    combo: { columns: "waterfall" },
    waterfall: { totalIndices: [5] },
    data: {
      categories: ["Start", "Q1", "Q2", "Q3", "Q4", "End"],
      series: [
        { name: "Delta", values: [100, 20, -15, 25, -10, 0] },
        { name: "Margin %", type: "line", values: [40, 42, 38, 44, 41, 43] },
      ],
    },
    decorations: { segmentLabels: true, seriesLabels: true },
  },
  {
    kind: "combo",
    width: 480,
    height: 300,
    title: "Combo: Mekko base + overlaid line",
    secondaryAxis: true,
    combo: { columns: "mekko" },
    data: {
      categories: ["EMEA", "Americas", "APAC"],
      series: [
        { name: "Prod A", values: [30, 45, 20] },
        { name: "Prod B", values: [20, 25, 30] },
        { name: "Share %", type: "line", values: [62, 70, 45] },
      ],
    },
    decorations: { seriesLabels: true },
  },
  {
    kind: "combo",
    width: 480,
    height: 300,
    title: "Combo: independent line axes (mixed KPIs)",
    combo: { lineAxes: "independent" },
    data: {
      categories: ["Jan", "Feb", "Mar", "Apr"],
      series: [
        { name: "Revenue", values: [40, 48, 45, 60] },
        { name: "NPS", type: "line", values: [30, 45, 42, 55] },
        { name: "Orders k", type: "line", values: [1200, 1500, 1400, 1900] },
      ],
    },
    decorations: { seriesLabels: true },
  },
  {
    kind: "tilemap",
    width: 480,
    height: 300,
    title: "Hexagonal tile map (US)",
    map: "us",
    tilemap: { shape: "hex" },
    data: {
      categories: ["CA", "TX", "NY", "FL", "IL", "WA", "MA", "CO", "GA", "NC", "OH", "PA", "MI", "AZ", "VA"],
      series: [{ name: "Stores", values: [120, 95, 88, 74, 51, 44, 38, 29, 27, 22, 40, 55, 33, 25, 30] }],
    },
    decorations: { segmentLabels: true },
  },
  {
    kind: "tilemap",
    width: 480,
    height: 300,
    title: "Tile map with per-region mini-bars",
    map: "us",
    tilemap: { glyph: "bars" },
    data: {
      categories: ["CA", "TX", "NY", "FL", "IL", "WA"],
      series: [
        { name: "Q1", values: [30, 20, 18, 15, 12, 10] },
        { name: "Q2", values: [35, 25, 20, 17, 14, 12] },
        { name: "Q3", values: [40, 28, 22, 19, 15, 13] },
        { name: "Q4", values: [48, 32, 26, 22, 18, 16] },
      ],
    },
  },
  {
    kind: "stacked100",
    width: 480,
    height: 300,
    title: "100% chart with negative segments",
    data: {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "New", values: [60, 70, 80] },
        { name: "Renewal", values: [40, 45, 50] },
        { name: "Returns", values: [-15, -10, -20] },
      ],
    },
    decorations: { segmentLabels: true, valueAxis: true, gridlines: true },
  },
  {
    kind: "doughnut",
    width: 480,
    height: 300,
    title: "Semi-circle gauge (market share)",
    pie: { semi: true },
    data: {
      categories: ["Us", "Rival A", "Rival B", "Others"],
      series: [{ name: "Share", values: [42, 28, 18, 12] }],
    },
    decorations: { segmentLabels: true },
  },
  {
    kind: "clustered",
    width: 480,
    height: 300,
    title: "Pareto chart (defect causes)",
    pareto: true,
    data: {
      categories: ["Scratches", "Dents", "Misalign", "Cracks", "Paint", "Other"],
      series: [{ name: "Defects", values: [120, 80, 45, 30, 15, 10] }],
    },
    decorations: { seriesLabels: true },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Bump chart (brand rank by year)",
    data: {
      categories: ["2021", "2022", "2023", "2024", "2025"],
      series: [
        { name: "Acme", values: [1, 2, 2, 1, 1] },
        { name: "Globex", values: [2, 1, 1, 3, 2] },
        { name: "Initech", values: [3, 3, 4, 2, 3] },
        { name: "Umbrella", values: [4, 4, 3, 4, 5] },
        { name: "Soylent", values: [5, 5, 5, 5, 4] },
      ],
    },
    decorations: { bump: true },
  },
  {
    kind: "line",
    width: 480,
    height: 300,
    title: "Horizontal profile chart (line)",
    horizontal: true,
    data: {
      categories: ["Strategy", "Execution", "Culture", "Innovation", "Finance", "Talent"],
      series: [{ name: "Score", values: [72, 58, 80, 65, 90, 48] }],
    },
    decorations: { segmentLabels: true, categoryAxis: true, valueAxis: true, gridlines: true },
  },
  {
    kind: "area",
    width: 480,
    height: 300,
    title: "Horizontal profile chart (stacked area)",
    horizontal: true,
    data: {
      categories: ["North", "South", "East", "West", "Central"],
      series: [
        { name: "Retail", values: [40, 55, 30, 48, 35] },
        { name: "Online", values: [25, 20, 35, 22, 30] },
      ],
    },
    decorations: { seriesLabels: true, categoryAxis: true, valueAxis: true, gridlines: true },
  },
  {
    kind: "treemap",
    width: 480,
    height: 300,
    title: "Two-level treemap (region | product)",
    data: {
      categories: [
        "EMEA | Cloud",
        "EMEA | Licenses",
        "EMEA | Services",
        "Americas | Cloud",
        "Americas | Licenses",
        "APAC | Cloud",
        "APAC | Hardware",
        "APAC | Services",
      ],
      series: [{ name: "Revenue", values: [60, 40, 25, 80, 50, 45, 30, 20] }],
    },
    decorations: { segmentLabels: true },
  },
  {
    kind: "sunburst",
    width: 480,
    height: 300,
    title: "Sunburst (region → product)",
    data: {
      categories: [
        "EMEA | Cloud",
        "EMEA | Licenses",
        "EMEA | Services",
        "Americas | Cloud",
        "Americas | Licenses",
        "APAC | Cloud",
        "APAC | Hardware",
      ],
      series: [{ name: "Revenue", values: [60, 40, 25, 80, 50, 45, 30] }],
    },
    decorations: { segmentLabels: true },
  },
  {
    kind: "violin",
    width: 480,
    height: 300,
    title: "Response-time distribution (ms)",
    data: {
      categories: ["API", "Web", "Batch"],
      series: Array.from({ length: 10 }, (_, i) => ({
        name: `s${i + 1}`,
        values: [90 + i * 4, 140 + i * 9, 300 + i * 22],
      })),
    },
    decorations: { categoryAxis: true, valueAxis: true },
  },
  {
    kind: "candlestick",
    width: 480,
    height: 300,
    title: "Share price — last two weeks",
    data: {
      categories: ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10"],
      series: [
        { name: "Open", values: [40, 42, 45, 44, 47, 46, 43, 45, 48, 47] },
        { name: "High", values: [43, 46, 47, 48, 49, 47, 46, 49, 50, 49] },
        { name: "Low", values: [39, 41, 43, 43, 45, 42, 42, 44, 46, 45] },
        { name: "Close", values: [42, 45, 44, 47, 46, 43, 45, 48, 47, 46] },
      ],
    },
    decorations: { valueAxis: true },
  },
  {
    kind: "gantt",
    width: 560,
    height: 300,
    title: "Launch plan — critical path highlighted",
    data: {
      categories: ["Research", "Design", "Build", "Copywriting", "QA", "Launch"],
      series: [
        { name: "Start", values: [1, 3, 8, 3, 14, 18] },
        { name: "End", values: [3, 8, 14, 5, 18, 20] },
        { name: "After", values: [null, 1, 2, 1, 3, 5] },
      ],
    },
    decorations: { criticalPath: true },
  },
  {
    kind: "boxplot",
    width: 480,
    height: 300,
    title: "Assay reproducibility (mean ± SD)",
    data: {
      categories: ["Lot A", "Lot B", "Lot C"],
      series: Array.from({ length: 8 }, (_, i) => ({
        name: `r${i + 1}`,
        values: [98 + ((i * 5) % 7) - 3, 90 + ((i * 3) % 9), 74 + i * 3],
      })),
    },
    boxplot: { meanSd: true },
    decorations: { valueAxis: true, segmentLabels: true },
  },
  {
    kind: "line",
    width: 480,
    height: 260,
    title: "KPI sparklines — last 12 weeks",
    data: {
      categories: Array.from({ length: 12 }, (_, i) => `w${i + 1}`),
      series: [
        { name: "Signups", values: [40, 44, 42, 50, 48, 60, 58, 72, 66, 80, 78, 92] },
        { name: "Revenue", values: [12, 13, 15, 14, 18, 17, 22, 25, 24, 30, 34, 33] },
        { name: "Churn %", values: [5.1, 4.8, 4.9, 4.4, 4.6, 4.0, 3.8, 3.9, 3.3, 3.1, 2.9, 2.7] },
        { name: "NPS", values: [30, 32, 31, 35, 38, 40, 39, 44, 47, 46, 52, 55] },
      ],
    },
    multiples: { columns: 1 },
    decorations: { sparkline: true },
  },
  {
    kind: "radar",
    width: 420,
    height: 320,
    title: "Monthly sales (radial bars)",
    data: {
      categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"],
      series: [{ name: "Units", values: [40, 55, 30, 70, 62, 48, 80, 35] }],
    },
    radar: { bars: true },
  },
  {
    kind: "radar",
    width: 460,
    height: 320,
    title: "Team capability (stacked radar)",
    data: {
      categories: ["Strategy", "Delivery", "Design", "Data", "Comms", "Ops"],
      series: [
        { name: "Alice", values: [4, 3, 5, 2, 4, 3] },
        { name: "Bob", values: [3, 4, 2, 5, 3, 4] },
        { name: "Cara", values: [2, 3, 3, 3, 5, 2] },
      ],
    },
    radar: { stacked: true },
    decorations: { seriesLabels: true },
  },
  {
    kind: "pie",
    width: 460,
    height: 320,
    title: "Region share (angle = revenue, radius = margin)",
    data: {
      categories: ["EMEA", "Americas", "APAC", "LatAm"],
      series: [
        { name: "Revenue", values: [45, 30, 18, 7] },
        { name: "Radius", values: [90, 55, 70, 40] },
      ],
    },
    pie: { variableRadius: true },
    decorations: { segmentLabels: true },
  },
  {
    kind: "heatmap",
    width: 460,
    height: 320,
    title: "Correlation matrix — size = magnitude, +/− marks the direction",
    data: {
      categories: ["Rev", "Cost", "Head", "NPS", "Churn"],
      series: [
        { name: "Rev", values: [1.0, 0.8, 0.6, 0.5, -0.7] },
        { name: "Cost", values: [0.8, 1.0, 0.7, 0.1, -0.3] },
        { name: "Head", values: [0.6, 0.7, 1.0, 0.2, -0.2] },
        { name: "NPS", values: [0.5, 0.1, 0.2, 1.0, -0.9] },
        { name: "Churn", values: [-0.7, -0.3, -0.2, -0.9, 1.0] },
      ],
    },
    // The canonical case for sign marks: sizeEncode suppresses the value
    // labels, so without them a printed copy has only hue to tell +0.8 from
    // -0.7 — and hue is the first thing a greyscale printer discards.
    heatmap: { sizeEncode: true, mode: "diverging", symbols: "sign" },
  },
  {
    kind: "heatmap",
    width: 500,
    height: 320,
    title: "Region KPIs (rows clustered)",
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [
        { name: "North", values: [80, 82, 85, 88] },
        { name: "South", values: [20, 22, 19, 25] },
        { name: "East", values: [78, 80, 83, 90] },
        { name: "West", values: [22, 18, 24, 21] },
        { name: "Central", values: [50, 52, 48, 55] },
        { name: "Metro", values: [81, 79, 86, 87] },
      ],
    },
    heatmap: { cluster: true },
    decorations: { segmentLabels: true },
  },
  {
    kind: "combo",
    width: 520,
    height: 320,
    title: "Revenue mix (stacked area) + margin %",
    data: {
      categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
      series: [
        { name: "Cloud", values: [20, 24, 28, 30, 35, 40] },
        { name: "Licenses", values: [15, 14, 16, 15, 17, 18] },
        { name: "Services", values: [10, 11, 12, 14, 13, 15] },
        { name: "Margin %", type: "line", values: [22, 24, 26, 25, 28, 30] },
      ],
    },
    combo: { columns: "area" },
    secondaryAxis: true,
    decorations: { seriesLabels: true },
  },
  {
    kind: "gantt",
    width: 620,
    height: 260,
    title: "Programme plan — cost and effort beside the bars",
    data: {
      categories: [
        "Discovery | Ana",
        "> Research | Ana",
        "> Concept | Ana",
        "Delivery | Ben",
        "> Build | Ben",
        "> QA | Cara",
      ],
      series: [
        { name: "Start", values: [null, 0, 3, null, 5, 11] },
        { name: "End", values: [null, 3, 5, null, 11, 14] },
        { name: "Column Cost k€", values: [120, 45, 75, 260, 190.5, 69.5] },
        { name: "Column FTE", values: [2, 1, 1, 3, 2, 1] },
      ],
    },
    decorations: { summaryBars: true },
  },
  {
    kind: "gantt",
    width: 560,
    height: 240,
    title: "Working-day timeline — weekends carry no width",
    data: {
      categories: ["Spec", "Build", "Review", "Ship"],
      series: [
        { name: "Start", values: [day("2026-01-05"), day("2026-01-08"), day("2026-01-15"), day("2026-01-20")] },
        { name: "End", values: [day("2026-01-08"), day("2026-01-15"), day("2026-01-20"), day("2026-01-22")] },
        { name: "Holiday", values: [day("2026-01-13"), null, null, null] },
        { name: "After", values: [null, 1, 2, 3] },
      ],
      dates: true,
    },
    gantt: { workdays: true },
  },
  {
    kind: "scatter",
    width: 520,
    height: 320,
    title: "Marginal histograms — distribution beside the relationship",
    data: {
      categories: Array.from({ length: 40 }, (_, i) => `A${i + 1}`),
      series: [
        // Two loose clusters, so the marginals have a shape worth reading.
        {
          name: "X",
          values: Array.from({ length: 40 }, (_, i) => (i < 22 ? 18 + ((i * 7) % 22) : 55 + ((i * 5) % 30))),
        },
        {
          name: "Y",
          values: Array.from({ length: 40 }, (_, i) => (i < 22 ? 22 + ((i * 11) % 26) : 52 + ((i * 3) % 34))),
        },
      ],
    },
    decorations: { marginals: "both", segmentLabels: false },
  },
  {
    kind: "bubble",
    width: 520,
    height: 300,
    title: "Overlap relief — markers nudged along Y only, by a disclosed cap",
    data: {
      categories: ["Alfa", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel"],
      series: [
        // Two tight clusters that would otherwise sit on top of each other.
        { name: "X", values: [30, 31, 32, 30.5, 68, 69, 70, 68.5] },
        { name: "Y", values: [40, 41, 39, 40.5, 62, 61, 63, 62.5] },
        { name: "Size", values: [70, 55, 60, 45, 80, 50, 65, 40] },
      ],
    },
    scatter: { spread: "y", spreadLimit: 6 },
  },
  {
    kind: "scatter",
    width: 520,
    height: 300,
    title: "Shape per group — the split survives greyscale and color blindness",
    data: {
      categories: [
        "Northwind",
        "Acme",
        "Globex",
        "Initech",
        "Umbra",
        "Soylent",
        "Vandelay",
        "Hooli",
        "Stark",
        "Wayne",
        "Tyrell",
        "Cyberdyne",
      ],
      series: [
        { name: "X", values: [12, 18, 22, 15, 44, 51, 47, 55, 74, 81, 78, 85] },
        { name: "Y", values: [30, 41, 25, 36, 55, 48, 62, 51, 72, 66, 80, 74] },
        { name: "Group", values: [1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3] },
      ],
    },
    footnote: "Groups: challengers, contenders, leaders.",
    scatter: { markers: ["circle", "diamond", "triangle"] },
  },
  {
    kind: "combo",
    width: 500,
    height: 300,
    // No secondaryAxis and no lineAxes: "independent" — a benchmark marker is
    // only meaningful measured against the columns it benchmarks.
    title: "Benchmark markers — consensus per quarter, on the columns' own scale",
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [
        { name: "Licenses", values: [28, 31, 30, 36] },
        { name: "Services", values: [12, 14, 17, 15] },
        { name: "Consensus", type: "marker", values: [42, 44, 50, 49] },
      ],
    },
    decorations: { totals: true },
  },
  {
    kind: "waterfall",
    width: 560,
    height: 300,
    title: "Bridge with “of which” detail columns — the walk steps over them",
    data: {
      categories: ["FY23", "Volume", "Cost", "> Labour", "> Freight", "> Energy", "FX", "FY24"],
      series: [{ name: "Delta", values: [86, 14, -12, -7, -3, -2, -4, 0] }],
    },
    waterfall: { totalIndices: [7], detailGroups: [{ of: 2, indices: [3, 4, 5] }] },
    decorations: { categoryAxis: true },
  },
  {
    kind: "gantt",
    width: 560,
    height: 280,
    title: "Plan grouped into owner lanes — dependencies follow their rows",
    data: {
      categories: ["Spec | Ana", "Build | Ben", "Review | Ana", "Ship | Ben", "Handover"],
      series: [
        {
          name: "Start",
          values: [day("2026-01-05"), day("2026-01-12"), day("2026-01-19"), day("2026-01-26"), day("2026-02-02")],
        },
        {
          name: "End",
          values: [day("2026-01-12"), day("2026-01-19"), day("2026-01-26"), day("2026-02-02"), day("2026-02-05")],
        },
        { name: "After", values: [null, 1, 2, 3, 4] },
      ],
      dates: true,
    },
    gantt: { lanes: "owner" },
    decorations: { summaryBars: true },
  },
];

const configs = [...kinds, ...features];
writeFileSync("examples/showcase.json", JSON.stringify(configs, null, 2) + "\n");
execFileSync("node", ["skill/scripts/render-pptx.mjs", "examples/showcase.json", "examples/showcase.pptx"], {
  stdio: "inherit",
});
console.log(`examples/showcase.json + examples/showcase.pptx (${configs.length} slides)`);
