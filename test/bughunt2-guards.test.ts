import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";

/**
 * Degenerate-frame / degenerate-scale guards found by a layout bug-hunt. Each is
 * byte-identical for normal charts (snapshots unchanged) and only repairs an edge
 * that previously emitted negative geometry or NaN coordinates.
 */

const negDims = (nodes: ReturnType<typeof buildChart>["nodes"]) =>
  nodes.filter(
    (n) =>
      ((n.kind === "rect" || n.kind === "text") && ((n as RectNode).w < 0 || (n as RectNode).h < 0)) ||
      Object.entries(n).some(([k, v]) => ["x", "y", "w", "h"].includes(k) && typeof v === "number" && Number.isNaN(v)),
  );

describe("funnel bands never go negative on a short, crowded frame", () => {
  it("floors band height", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: 640,
      height: 60, // 20 stages + 1.5em gaps can't fit → bands went negative
      data: {
        categories: Array.from({ length: 20 }, (_, i) => `S${i}`),
        series: [{ name: "v", values: Array.from({ length: 20 }, (_, i) => 20 - i) }],
      },
    };
    expect(negDims(buildChart(cfg).nodes)).toHaveLength(0);
  });
});

describe("butterfly keeps non-negative geometry on a very narrow frame", () => {
  it("floors plot width and header width", () => {
    const cfg: ChartConfig = {
      kind: "butterfly",
      width: 40, // narrower than the two value-label strips
      height: 300,
      data: {
        categories: ["A"],
        series: [
          { name: "left", values: [5] },
          { name: "right", values: [3] },
        ],
      },
    };
    expect(negDims(buildChart(cfg).nodes)).toHaveLength(0);
  });
});

describe("heatmap forced-diverging on single-signed data draws no phantom zero", () => {
  it("omits the 0 tick when zero is outside the data range", () => {
    const cfg: ChartConfig = {
      kind: "heatmap",
      width: 640,
      height: 400,
      heatmap: { mode: "diverging" },
      data: { categories: ["A", "B", "C"], series: [{ name: "r", values: [10, 55, 100] }] },
    };
    expect(buildChart(cfg).nodes.some((n) => n.name === "legend-zero")).toBe(false);
  });

  it("still draws the 0 tick, on-strip, when the data spans zero", () => {
    const cfg: ChartConfig = {
      kind: "heatmap",
      width: 640,
      height: 400,
      heatmap: { mode: "diverging" },
      data: { categories: ["A", "B", "C"], series: [{ name: "r", values: [-40, 0, 60] }] },
    };
    const zero = buildChart(cfg).nodes.find((n) => n.name === "legend-zero") as TextNode | undefined;
    expect(zero).toBeDefined();
    expect(zero!.x).toBeGreaterThan(0);
  });
});

describe("log scale with a manual min above the data does not blank the axis", () => {
  it("clamps to at least one decade instead of producing NaN geometry", () => {
    const cfg: ChartConfig = {
      kind: "clustered",
      width: 640,
      height: 400,
      logScale: true,
      scale: { min: 1_000_000 }, // far above the data → empty ticks → NaN toY
      data: { categories: ["A", "B"], series: [{ name: "s", values: [3, 9] }] },
    };
    expect(negDims(buildChart(cfg).nodes)).toHaveLength(0);
  });
});
