import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, RectNode, TextNode } from "../src/core/scene";

/** Backlog batch D: waterfall gap-to-target, bands, fillBetween, heatmap totals. */

describe("waterfall budget-vs-actual (Target row)", () => {
  const cfg: ChartConfig = {
    kind: "waterfall",
    ...DEFAULT_SIZE,
    data: {
      categories: ["FY23", "Volume", "Price", "FY24"],
      series: [
        { name: "Delta", values: [80, 10, 6, 0] },
        { name: "Target", values: [null, null, null, 110] },
      ],
    },
    waterfall: { totalIndices: [3] },
  };
  const s = buildChart(cfg);

  it("draws the target tick and a hatched gap segment on the final total", () => {
    const tick = s.nodes.find((n) => n.name === "target-3") as LineNode;
    expect(tick).toBeDefined();
    const gap = s.nodes.find((n) => n.name === "target-gap-3") as RectNode;
    expect(gap.pattern).toBe("diagonal");
    // Gap spans actual (96) → target (110): its bottom meets the bar top.
    const bar = s.nodes.find((n) => n.name === "bar-3") as RectNode;
    expect(gap.y + gap.h).toBeCloseTo(bar.y, 1);
    const label = s.nodes.find((n) => n.name === "target-gap-label-3") as TextNode;
    expect(label.text).toBe("Gap +14");
  });

  it("Target row never renders as waterfall bars and the scale covers it", () => {
    expect(s.nodes.filter((n) => n.name?.startsWith("bar-"))).toHaveLength(4);
    const gap = s.nodes.find((n) => n.name === "target-gap-3") as RectNode;
    expect(gap.y).toBeGreaterThan(0); // inside the plot — scale widened
  });

  it("a missed target (actual above target) hatches downward with a negative gap", () => {
    const over = buildChart({
      ...cfg,
      data: { ...cfg.data, series: [cfg.data.series[0], { name: "Target", values: [null, null, null, 90] }] },
    });
    expect((over.nodes.find((n) => n.name === "target-gap-label-3") as TextNode).text).toBe("Gap -6");
  });
});

describe("line confidence bands and fill-between", () => {
  const line: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: {
      categories: ["2024", "2025", "2026"],
      series: [
        { name: "Forecast", values: [50, 56, 63] },
        { name: "Band low", values: [50, 52, 55] },
        { name: "Band high", values: [50, 60, 72] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("shades a ribbon between Band low/high without drawing them as lines", () => {
    const s = buildChart(line);
    const slabs = s.nodes.filter((n) => n.name?.startsWith("band-ribbon"));
    expect(slabs.length).toBeGreaterThan(20);
    // Only the forecast line renders; band rows are not series.
    expect(s.nodes.some((n) => n.name === "line-1-1")).toBe(false);
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    // Ribbon renders behind the line (earlier in the node list).
    expect(s.nodes.findIndex((n) => n.name?.startsWith("band-ribbon"))).toBeLessThan(
      s.nodes.findIndex((n) => n.name === "line-0-1"),
    );
    // The band widens the scale: max tick covers 72.
    const slab = s.nodes.find((n) => n.name?.startsWith("band-ribbon-1")) as RectNode;
    expect(slab.y).toBeGreaterThan(0);
  });

  it("fillBetween shades the gap between two series", () => {
    const s = buildChart({
      kind: "line",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Plan", values: [40, 50, 60] },
          { name: "Actual", values: [38, 46, 52] },
        ],
      },
      decorations: { fillBetween: [0, 1], segmentLabels: false },
    });
    expect(s.nodes.filter((n) => n.name?.startsWith("fill-between")).length).toBeGreaterThan(20);
    // Both series still draw as lines.
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    expect(s.nodes.some((n) => n.name === "line-1-1")).toBe(true);
  });
});

describe("heatmap marginal totals", () => {
  const heat: ChartConfig = {
    kind: "heatmap",
    ...DEFAULT_SIZE,
    heatmap: { totals: "both" },
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [10, 40] },
        { name: "South", values: [20, 30] },
      ],
    },
  };

  it("adds row and column sum strips outside the color scale", () => {
    const s = buildChart(heat);
    expect((s.nodes.find((n) => n.name === "row-total-0") as TextNode).text).toBe("50");
    expect((s.nodes.find((n) => n.name === "col-total-1") as TextNode).text).toBe("70");
    // Totals sit outside the matrix: right of the last cell / below the last row.
    const cell = s.nodes.find((n) => n.name === "cell-0-1") as RectNode;
    const rowTotal = s.nodes.find((n) => n.name === "row-total-bg-0") as RectNode;
    expect(rowTotal.x).toBeGreaterThanOrEqual(cell.x + cell.w);
    // Neutral fill, not the value color scale.
    expect(rowTotal.fill).toBe("#f0efec");
  });

  it("row-only mode omits the column strip; default has neither", () => {
    const rowOnly = buildChart({ ...heat, heatmap: { totals: "row" } });
    expect(rowOnly.nodes.some((n) => n.name === "row-total-0")).toBe(true);
    expect(rowOnly.nodes.some((n) => n.name === "col-total-0")).toBe(false);
    const plain = buildChart({ ...heat, heatmap: {} });
    expect(plain.nodes.some((n) => n.name?.includes("total"))).toBe(false);
  });
});
