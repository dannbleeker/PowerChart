import { buildChart } from "../core/chart";
import { sceneToSvg } from "../render/svg";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import type { ChartConfig } from "../core/types";

const gallery = document.getElementById("gallery")!;

/** ISO date → days since epoch (the datasheet's calendar parsing, inlined). */
function dstr(iso: string): number {
  return Math.round(Date.parse(iso) / 86400000);
}

const extras: ChartConfig[] = [
  {
    ...sampleConfig("stacked"),
    title: "Stacked bar (rotated column)",
    horizontal: true,
    decorations: { totals: true, seriesLabels: true, segmentLabels: true, categoryAxis: true },
  },
  {
    ...sampleConfig("mekko"),
    title: "Mekko with units (X extent row)",
    data: {
      categories: ["EMEA", "Americas", "APAC"],
      series: [
        { name: "Enterprise", values: [42, 55, 18] },
        { name: "SMB", values: [28, 30, 22] },
        { name: "Consumer", values: [14, 25, 30] },
      ],
      xExtent: [50, 30, 20],
    },
  },
  {
    ...sampleConfig("stacked100"),
    title: "100% with 100%= row (short columns)",
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
    ...sampleConfig("stacked"),
    title: "Level difference arrow (series 1)",
    decorations: { difference: { from: 0, to: 3, series: 0 }, segmentLabels: true },
  },
  {
    ...sampleConfig("stacked"),
    title: "Segment order: descending + pinned axis max 150",
    segmentOrder: "descending",
    scale: { max: 150 },
    decorations: { totals: true, valueAxis: true, gridlines: true },
  },
  {
    ...sampleConfig("clustered"),
    title: "Axis break 80–580 (one outlier column)",
    data: {
      categories: ["Q1", "Q2", "Q3", "Q4"],
      series: [{ name: "Sales", values: [42, 55, 620, 61] }],
    },
    axisBreak: { from: 80, to: 580 },
    decorations: { segmentLabels: true, totals: false, valueAxis: true, gridlines: true, seriesLabels: false },
  },
  {
    ...sampleConfig("waterfall"),
    title: "Stacked waterfall (two series per delta)",
    data: {
      categories: ["FY23", "Organic", "M&A", "Cost", "FY24"],
      series: [
        { name: "Europe", values: [50, 8, 5, -6, 0] },
        { name: "Americas", values: [36, 6, 9, -6, 0] },
      ],
    },
    waterfall: { totalIndices: [4] },
  },
  {
    ...sampleConfig("combo"),
    title: "Combo: stacked columns + margin line",
  },
  {
    ...sampleConfig("pie"),
    title: "Pie (category + % labels)",
  },
  {
    ...sampleConfig("doughnut"),
    title: "Doughnut with center total",
  },
  {
    ...sampleConfig("clustered"),
    title: "Log scale + axis title",
    data: {
      categories: ["Seed", "A", "B", "C"],
      series: [{ name: "Valuation", values: [4, 40, 220, 1900] }],
    },
    logScale: true,
    valueAxisTitle: "$m (log)",
    decorations: { valueAxis: true, gridlines: true, seriesLabels: false, segmentLabels: true },
  },
  {
    ...sampleConfig("stacked"),
    title: "Value-line-anchored difference + de-DE locale",
    decorations: {
      totals: true,
      valueLines: [{ mode: "value", value: 100 }],
      difference: { from: 0, to: 3, fromValueLine: 0, percent: false },
    },
    numberFormat: { decimals: 1, locale: "de-DE" },
  },
  {
    ...sampleConfig("gantt"),
    title: "Calendar Gantt (dates in the datasheet)",
    data: {
      categories: ["Scoping", "Design", "Build", "Test", "Rollout"],
      series: [
        { name: "Start", values: ["2026-01-05", "2026-02-01", "2026-03-10", "2026-06-01", "2026-07-01"].map(dstr) },
        { name: "End", values: ["2026-02-01", "2026-03-15", "2026-06-01", "2026-07-01", "2026-07-20"].map(dstr) },
        { name: "Milestone", values: [null, null, dstr("2026-06-01"), null, dstr("2026-07-20")] },
      ],
      dates: true,
    },
  },
  {
    ...sampleConfig("stacked"),
    title: "Stacked with difference arrow + value line",
    decorations: {
      totals: true,
      difference: { from: 0, to: 3 },
      valueLine: { mode: "mean" },
    },
  },
  {
    ...sampleConfig("waterfall"),
    title: "Waterfall with CAGR-style difference",
    decorations: { categoryAxis: true, difference: { from: 0, to: 5 } },
  },
];

for (const { kind, label } of CHART_KINDS) {
  addFigure(`${label} (sample)`, sampleConfig(kind));
}
for (const cfg of extras) {
  addFigure(cfg.title ?? "extra", cfg);
}

function addFigure(caption: string, cfg: ChartConfig) {
  const fig = document.createElement("figure");
  const cap = document.createElement("figcaption");
  cap.textContent = caption;
  fig.appendChild(cap);
  const holder = document.createElement("div");
  holder.innerHTML = sceneToSvg(buildChart(cfg), { background: "#ffffff" });
  fig.appendChild(holder);
  gallery.appendChild(fig);
}
