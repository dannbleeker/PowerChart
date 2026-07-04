import { buildChart } from "../core/chart";
import { sceneToSvg } from "../render/svg";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import type { ChartConfig } from "../core/types";

const gallery = document.getElementById("gallery")!;

const extras: ChartConfig[] = [
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
