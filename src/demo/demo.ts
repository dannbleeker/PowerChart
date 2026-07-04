import { buildChart } from "../core/chart";
import { sceneToSvg } from "../render/svg";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import type { ChartConfig } from "../core/types";

const gallery = document.getElementById("gallery")!;

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
