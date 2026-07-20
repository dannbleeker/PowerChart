import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode } from "../src/core/scene";

/**
 * Backlog batch M — combo family: waterfall/mekko base under a line overlay,
 * and independent per-line-series axes.
 */

describe("combo waterfall base", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    secondaryAxis: true,
    combo: { columns: "waterfall" },
    waterfall: { totalIndices: [3] },
    data: {
      categories: ["Start", "Up", "Down", "End"],
      series: [
        { name: "Delta", values: [100, 20, -15, 0] },
        { name: "Margin %", type: "line", values: [40, 42, 38, 41] },
      ],
    },
    decorations: { segmentLabels: false, seriesLabels: false },
  };

  it("draws waterfall bars with a line overlaid on the secondary axis", () => {
    const s = buildChart(cfg);
    // Waterfall base bars.
    expect(s.nodes.some((n) => n.name === "bar-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "bar-3")).toBe(true);
    // Line overlay + secondary axis ticks.
    expect(s.nodes.filter((n) => n.name?.startsWith("combo-marker-0-"))).toHaveLength(4);
    expect(s.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
  });
});

describe("combo mekko base", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    secondaryAxis: true,
    combo: { columns: "mekko" },
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "P1", values: [30, 45, 20] },
        { name: "P2", values: [20, 25, 30] },
        { name: "Share %", type: "line", values: [62, 70, 45] },
      ],
    },
    decorations: { segmentLabels: false, seriesLabels: false },
  };

  it("draws mekko columns (variable width) with a line over them", () => {
    const s = buildChart(cfg);
    // Mekko segments exist and column widths differ (A total 50 vs B total 70).
    const segs = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("seg-"));
    expect(segs.length).toBeGreaterThan(0);
    const widths = new Set(segs.map((r) => Math.round(r.w)));
    expect(widths.size).toBeGreaterThan(1); // variable-width columns
    expect(s.nodes.filter((n) => n.name?.startsWith("combo-marker-0-"))).toHaveLength(3);
  });
});

describe("combo independent line axes", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    combo: { lineAxes: "independent" },
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "Revenue", values: [40, 60] },
        { name: "Small KPI", type: "line", values: [10, 20] },
        { name: "Huge KPI", type: "line", values: [1000, 2000] },
      ],
    },
    decorations: { segmentLabels: false, seriesLabels: false },
  };

  it("scales each line to its own range so both are visible despite 100× units", () => {
    const s = buildChart(cfg);
    const y = (nm: string) => (s.nodes.find((n): n is RectNode => n.name === nm) as RectNode).y;
    // Line 0 (10→20) and line 1 (1000→2000): each spans its own range, so the
    // min points share a y and the max points share a y.
    expect(y("combo-marker-0-0")).toBeCloseTo(y("combo-marker-1-0"), 0); // both at their min
    expect(y("combo-marker-0-1")).toBeCloseTo(y("combo-marker-1-1"), 0); // both at their max
    expect(y("combo-marker-0-1")).toBeLessThan(y("combo-marker-0-0")); // max is higher up
    // Independent mode forces value labels (no shared axis to read).
    expect(s.nodes.some((n) => n.name === "combo-label-1-1")).toBe(true);
    // No single shared secondary axis in independent mode.
    expect(s.nodes.some((n) => n.name === "secondary-axis")).toBe(false);
  });

  it("shared axis (default) keeps one secondary scale", () => {
    const shared = buildChart({ ...cfg, secondaryAxis: true, combo: {} });
    expect(shared.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
  });
});

describe("combo waterfall base — shared-axis line overflow/underflow", () => {
  const mk = (line: number[]): ChartConfig => ({
    kind: "combo",
    ...DEFAULT_SIZE,
    combo: { columns: "waterfall" }, // shared axis (no secondaryAxis)
    data: {
      categories: ["A", "B", "C", "D"],
      series: [
        { name: "Cols", values: [100, 20, 30, 10] }, // running peak 160, trough 0
        { name: "Line", type: "line", values: line },
      ],
    },
    decorations: { segmentLabels: false, seriesLabels: false },
  });

  const lineYs = (s: ReturnType<typeof buildChart>) =>
    s.nodes
      .filter((n) => n.name?.startsWith("combo-marker-0-"))
      .map((n) => (n.kind === "rect" ? n.y : n.kind === "ellipse" ? n.cy : NaN))
      .filter((y) => !Number.isNaN(y));

  it("keeps a line that dips below the waterfall trough on-plot", () => {
    // #157 stretched the axis MAX for an overflowing line but left the floor at 0,
    // so a negative line point plotted off the bottom of the plot.
    const s = buildChart(mk([-80, -40, 10, 50]));
    const ys = lineYs(s);
    expect(ys.length).toBe(4);
    for (const y of ys) expect(y).toBeLessThanOrEqual(DEFAULT_SIZE.height + 1); // not off the bottom
  });

  it("still keeps a line that overshoots the peak on-plot (unchanged)", () => {
    const s = buildChart(mk([200, 40, 30, 10]));
    for (const y of lineYs(s)) expect(y).toBeGreaterThanOrEqual(-1); // not off the top
  });
});
