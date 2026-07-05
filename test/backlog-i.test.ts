import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, LineNode, RectNode, TextNode } from "../src/core/scene";

/**
 * Backlog batch I — more §2 within-kind gaps: scatter/bubble continuous
 * color scale, smoothed lines, waterfall grouping spacers.
 */

describe("scatter/bubble continuous color scale", () => {
  const base: ChartConfig = {
    kind: "scatter",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C", "D"],
      series: [
        { name: "X", values: [1, 2, 3, 4] },
        { name: "Y", values: [4, 3, 2, 1] },
        { name: "Color", values: [0, 10, 20, 30] },
      ],
    },
  };

  it("maps points onto a ramp and draws a gradient legend", () => {
    const s = buildChart(base);
    const points = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("point-"));
    const fills = new Set(points.map((p) => p.fill));
    expect(fills.size).toBe(4); // four distinct colors along the ramp
    expect(s.nodes.some((n) => n.name === "color-legend-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "color-legend-min")).toBe(true);
    expect(s.nodes.some((n) => n.name === "color-legend-max")).toBe(true);
  });

  it("supersedes group coloring (no group chips under a color scale)", () => {
    const s = buildChart({
      ...base,
      data: { ...base.data, series: [...base.data.series, { name: "Group", values: [1, 2, 1, 2] }] },
    });
    expect(s.nodes.some((n) => n.name?.startsWith("legend-chip-"))).toBe(false);
    expect(s.nodes.some((n) => n.name === "color-legend-0")).toBe(true);
  });

  it("no color row → no gradient legend (plain scatter)", () => {
    const s = buildChart({
      ...base,
      data: { categories: ["A", "B"], series: [{ name: "X", values: [1, 2] }, { name: "Y", values: [3, 4] }] },
    });
    expect(s.nodes.some((n) => n.name?.startsWith("color-legend"))).toBe(false);
  });
});

describe("smoothed lines", () => {
  const base: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "v", values: [1, 4, 2, 5] }] },
    decorations: { segmentLabels: false },
  };

  it("default draws straight connectors", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    expect(s.nodes.some((n) => n.name?.startsWith("line-0-1-s"))).toBe(false);
  });

  it("smooth replaces straight segments with a sampled spline polyline", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, smooth: true } });
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(false);
    const sampled = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.match(/^line-0-\d+-s\d+$/));
    // 3 segments × 16 samples.
    expect(sampled.length).toBe(48);
  });

  it("stepped wins over smooth (mutually exclusive)", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, smooth: true, stepped: "after" } });
    expect(s.nodes.some((n) => n.name?.startsWith("line-0-1-s"))).toBe(false);
    expect(s.nodes.some((n) => n.name === "line-0-1a")).toBe(true);
  });
});

describe("waterfall grouping spacers", () => {
  const cfg: ChartConfig = {
    kind: "waterfall",
    ...DEFAULT_SIZE,
    data: {
      categories: ["FY23", "Volume", "Price", "", "Cost", "FX", "", "FY24"],
      series: [{ name: "Delta", values: [86, 14, 9, null, -12, -4, null, 0] }],
    },
    waterfall: { totalIndices: [7], spacerIndices: [3, 6] },
    decorations: { categoryAxis: true },
  };
  const s = buildChart(cfg);

  it("draws no bar in a spacer slot", () => {
    expect(s.nodes.some((n) => n.name === "bar-3")).toBe(false);
    expect(s.nodes.some((n) => n.name === "bar-6")).toBe(false);
    expect(s.nodes.some((n) => n.name === "bar-1")).toBe(true);
    expect(s.nodes.some((n) => n.name === "bar-4")).toBe(true);
  });

  it("carries the running total across the gap (final total unaffected)", () => {
    // 86 + 14 + 9 − 12 − 4 = 93.
    const total = s.nodes.find((n): n is TextNode => n.kind === "text" && n.name === "label-7");
    expect(total?.text).toBe("93");
  });

  it("bridges the connector across the spacer slot", () => {
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && n.name === "spacer-bridge-3")).toBe(true);
    expect(s.nodes.some((n) => n.name === "spacer-bridge-6")).toBe(true);
  });

  it("without spacers, those categories render normally", () => {
    const plain = buildChart({ ...cfg, waterfall: { totalIndices: [7] } });
    expect(plain.nodes.some((n) => n.name?.startsWith("spacer-bridge"))).toBe(false);
  });
});
