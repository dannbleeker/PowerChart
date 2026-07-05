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
];

const configs = [...kinds, ...features];
writeFileSync("examples/showcase.json", JSON.stringify(configs, null, 2) + "\n");
execFileSync("node", ["skill/scripts/render-pptx.mjs", "examples/showcase.json", "examples/showcase.pptx"], {
  stdio: "inherit",
});
console.log(`examples/showcase.json + examples/showcase.pptx (${configs.length} slides)`);
