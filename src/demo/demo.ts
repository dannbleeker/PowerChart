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
    ...sampleConfig("stacked"),
    title: "Clustered-stacked (two stack groups per category)",
    data: {
      categories: ["2024", "2025"],
      series: [
        { name: "EU Product", values: [30, 36], stack: 0 },
        { name: "EU Services", values: [12, 16], stack: 0 },
        { name: "US Product", values: [26, 33], stack: 1 },
        { name: "US Services", values: [10, 14], stack: 1 },
      ],
    },
    decorations: { totals: true, segmentLabels: true, seriesLabels: true },
  },
  {
    ...sampleConfig("waterfall"),
    title: "Rotated waterfall (horizontal)",
    horizontal: true,
  },
  {
    ...sampleConfig("mekko"),
    title: "Rotated Mekko (horizontal)",
    horizontal: true,
  },
  {
    ...sampleConfig("gantt"),
    title: "Gantt: weeks, weekend shading, sections, remarks",
    data: {
      categories: [
        "Phase 1 — Discovery",
        "> Interviews | Anna | 12 done",
        "> Synthesis | Ben",
        "Phase 2 — Delivery",
        "> Build | Cato | at risk",
      ],
      series: [
        { name: "Start", values: [null, "2026-01-05", "2026-01-19", null, "2026-01-26"].map((d) => (d ? dstr(d) : null)) },
        { name: "End", values: [null, "2026-01-23", "2026-01-30", null, "2026-02-20"].map((d) => (d ? dstr(d) : null)) },
        { name: "After", values: [null, null, 2, null, 3] },
        { name: "Today", values: [dstr("2026-02-02"), null, null, null, null] },
      ],
      dates: true,
    },
  },
  {
    ...sampleConfig("scatter"),
    title: "Scatter: partition lines, trend, group legend",
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
    title: "Gantt: owners, dependencies, today line",
    data: {
      categories: ["Scoping | Anna", "Design | Ben", "Build | Cato", "Test | Dee", "Rollout | Anna"],
      series: [
        { name: "Start", values: ["2026-01-05", "2026-02-01", "2026-03-10", "2026-06-01", "2026-07-01"].map(dstr) },
        { name: "End", values: ["2026-02-01", "2026-03-15", "2026-06-01", "2026-07-01", "2026-07-20"].map(dstr) },
        { name: "Milestone", values: [null, null, dstr("2026-06-01"), null, dstr("2026-07-20")] },
        { name: "After", values: [null, 1, 2, 3, 4] },
        { name: "Today", values: [dstr("2026-04-15"), null, null, null, null] },
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

// --- Elements showcase -------------------------------------------------------
import { buildCheckbox, buildHarveyBall, buildProcessFlow, buildTableScene } from "../core/elements";

const elements = document.getElementById("elements")!;
const elementScenes: [string, ReturnType<typeof buildHarveyBall>][] = [
  ["Harvey balls 0 / 25 / 50 / 75 / 100%", combineRow([0, 0.25, 0.5, 0.75, 1].map((f) => buildHarveyBall(f, 28)), 8)],
  ["Checkboxes", combineRow((["yes", "no", "partial"] as const).map((s) => buildCheckbox(s, 24)), 10)],
  ["Process flow (step 3 active)", buildProcessFlow(["Scope", "Design", "Build", "Test", "Launch"], 2, 520, 44)],
  ["Table", buildTableScene([["", "2024", "2025", "Δ"], ["Revenue", "78", "91", "+13"], ["EBITDA", "21", "27", "+6"], ["Margin", "27%", "30%", "+3pp"]], 520)],
];
for (const [caption, scene] of elementScenes) {
  const fig = document.createElement("figure");
  const cap = document.createElement("figcaption");
  cap.textContent = caption;
  fig.appendChild(cap);
  const holder = document.createElement("div");
  holder.innerHTML = sceneToSvg(scene, { background: "#ffffff" });
  fig.appendChild(holder);
  elements.appendChild(fig);
}

/** Lay several small scenes out in one row. */
function combineRow(scenes: ReturnType<typeof buildHarveyBall>[], gap: number) {
  let x = 0;
  const nodes = scenes.flatMap((s) => {
    const shifted = s.nodes.map((n) => shiftNode(structuredClone(n), x));
    x += s.width + gap;
    return shifted;
  });
  return { width: x - gap, height: Math.max(...scenes.map((s) => s.height)), nodes };
}
function shiftNode<T extends { kind: string }>(n: T, dx: number): T {
  const node = n as unknown as Record<string, number>;
  for (const k of ["x", "x1", "x2", "cx"]) if (k in node) node[k] += dx;
  return n;
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
