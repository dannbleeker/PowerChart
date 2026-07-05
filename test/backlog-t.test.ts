import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, RectNode, TextNode } from "../src/core/scene";

/**
 * Batch T — heatmap cell-size encoding, heatmap row clustering (dendrogram),
 * and the stacked-area combo base.
 */

describe("heatmap cell-size encoding", () => {
  const cfg: ChartConfig = {
    kind: "heatmap",
    ...DEFAULT_SIZE,
    data: {
      categories: ["X"],
      series: [
        { name: "big", values: [1] },
        { name: "small", values: [0.25] },
      ],
    },
    heatmap: { sizeEncode: true },
  };
  const s = buildChart(cfg);
  const cell = (ri: number) => s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === `cell-${ri}-0`)!;

  it("sizes each cell by its magnitude (area ∝ |value|)", () => {
    const big = cell(0);
    const small = cell(1);
    // √(0.25) = 0.5 → the small cell's side is about half the big one's.
    expect(small.w).toBeLessThan(big.w);
    expect(small.w).toBeGreaterThan(0);
    expect(Math.abs(small.w / big.w - 0.5)).toBeLessThan(0.05);
    // Cells stay square and centred.
    expect(big.w).toBeCloseTo(big.h, 5);
  });

  it("plain heatmaps still fill the whole cell", () => {
    const plain = buildChart({ ...cfg, heatmap: {} });
    const full = plain.nodes.find((n): n is RectNode => n.name === "cell-0-0")!;
    const sized = cell(0);
    expect(full.w).toBeGreaterThan(sized.w);
  });
});

describe("heatmap row clustering", () => {
  const cfg: ChartConfig = {
    kind: "heatmap",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "Hi1", values: [80, 82, 85] },
        { name: "Lo1", values: [20, 22, 19] },
        { name: "Hi2", values: [78, 80, 83] },
        { name: "Lo2", values: [22, 18, 24] },
      ],
    },
    heatmap: { cluster: true },
  };
  const s = buildChart(cfg);
  const rowName = (ri: number) => (s.nodes.find((n): n is TextNode => n.kind === "text" && n.name === `row-${ri}`)!).text;

  it("reorders rows so similar rows are adjacent and draws a dendrogram", () => {
    const order = [0, 1, 2, 3].map(rowName);
    const hiPositions = [order.indexOf("Hi1"), order.indexOf("Hi2")].sort((a, b) => a - b);
    const loPositions = [order.indexOf("Lo1"), order.indexOf("Lo2")].sort((a, b) => a - b);
    // Each similar pair ends up adjacent.
    expect(hiPositions[1] - hiPositions[0]).toBe(1);
    expect(loPositions[1] - loPositions[0]).toBe(1);
    // The dendrogram tree is drawn.
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && n.name === "dendro-v")).toBe(true);
    expect(s.nodes.some((n) => n.name === "dendro-h")).toBe(true);
  });

  it("leaves rows in sheet order without the flag", () => {
    const plain = buildChart({ ...cfg, heatmap: {} });
    const name0 = (plain.nodes.find((n): n is TextNode => n.name === "row-0")!).text;
    expect(name0).toBe("Hi1");
    expect(plain.nodes.some((n) => n.name === "dendro-v")).toBe(false);
  });
});

describe("combo stacked-area base", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Jan", "Feb", "Mar"],
      series: [
        { name: "Cloud", values: [20, 24, 28] },
        { name: "Licenses", values: [15, 14, 16] },
        { name: "Margin %", type: "line", values: [22, 24, 26] },
      ],
    },
    combo: { columns: "area" },
    secondaryAxis: true,
  };
  const s = buildChart(cfg);

  it("draws a stacked area base with the line overlaid", () => {
    // The area base emits filled area slabs...
    expect(s.nodes.some((n) => n.name?.startsWith("area-"))).toBe(true);
    // ...and a secondary axis for the line series.
    expect(s.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
    // ...and line segments over the top.
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("combo-line-"))).toBe(true);
  });
});
