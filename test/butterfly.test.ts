import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig, Series } from "../src/core/types";

/** Butterfly / tornado — stacked flanks, narrow-frame geometry, value ticks. */

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

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */
const W = 480;

const H = 300;

const regionSeries = (n: number): Series[] =>
  REGIONS.slice(0, n).map((name, i) => ({ name, values: [1 + i, 2 + i, 3 + i] }));

/** Nodes that carry an x/w box and stick out past the right edge of the canvas. */
const offRight = (nodes: SceneNode[], namePrefix: string, width = W) =>
  nodes.filter((n) => {
    const b = n as Partial<RectNode>;
    return n.name?.startsWith(namePrefix) && b.x != null && b.x + (b.w ?? 0) > width;
  });

const REGIONS = ["Northern Europe", "Southern Europe", "North America", "Asia Pacific", "Latin America", "Middle East"];

describe("butterfly value ticks", () => {
  const base: ChartConfig = {
    kind: "butterfly",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "Left", values: [10, 20] },
        { name: "Right", values: [15, 25] },
      ],
    },
  };

  it("adds no axis chrome by default", () => {
    const plain = buildChart(base);
    expect(plain.nodes.some((n) => n.name?.startsWith("tick-"))).toBe(false);
    expect(plain.nodes.some((n) => n.name?.startsWith("gridline-"))).toBe(false);
  });

  it("valueAxis + gridlines draw mirrored ticks and gridlines on both flanks", () => {
    const axed = buildChart({ ...base, decorations: { valueAxis: true, gridlines: true } });
    expect(axed.nodes.some((n) => n.name?.startsWith("gridline-"))).toBe(true);
    const ticks = axed.nodes.filter((n): n is TextNode => !!n.name?.startsWith("tick-"));
    const lefts = ticks.filter((n) => n.name!.endsWith("-l"));
    const rights = ticks.filter((n) => n.name!.endsWith("-r"));
    expect(lefts.length).toBeGreaterThan(1);
    expect(lefts.length).toBe(rights.length); // mirrored
    expect(axed.nodes.some((n) => n.name === "tick-0-l")).toBe(true);
  });
});

describe("butterfly stacked flanks", () => {
  const cfg: ChartConfig = {
    kind: "butterfly",
    ...DEFAULT_SIZE,
    butterfly: { split: 2 },
    data: {
      categories: ["A", "B"],
      series: [
        { name: "L1", values: [10, 12] },
        { name: "L2", values: [6, 8] },
        { name: "R1", values: [14, 9] },
        { name: "R2", values: [4, 5] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("stacks split series on the left and the rest on the right", () => {
    const s = buildChart(cfg);
    // Four segments per category (2 left + 2 right).
    expect(s.nodes.filter((n) => n.name?.match(/^seg-\d+-0$/))).toHaveLength(4);
    const l1 = s.nodes.find((n): n is RectNode => n.name === "seg-0-0")!; // L1 innermost-left
    const l2 = s.nodes.find((n): n is RectNode => n.name === "seg-1-0")!; // L2 stacked further left
    expect(l2.x + l2.w).toBeCloseTo(l1.x, 1); // L2 sits just left of L1 (contiguous)
    const r1 = s.nodes.find((n): n is RectNode => n.name === "seg-2-0")!; // R1 innermost-right
    const r2 = s.nodes.find((n): n is RectNode => n.name === "seg-3-0")!;
    expect(r2.x).toBeCloseTo(r1.x + r1.w, 1); // R2 just right of R1
    // Stacked mode shows a legend of every series.
    expect(s.nodes.some((n) => n.name === "legend-3")).toBe(true);
  });

  it("default (no split) keeps the classic two-series butterfly", () => {
    const s = buildChart({ ...cfg, butterfly: undefined });
    expect(s.nodes.some((n) => n.name === "seg-2-0")).toBe(false); // series 2+ ignored
    expect(s.nodes.some((n) => n.name === "header-0")).toBe(true); // two headers, not a legend
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

describe("butterfly", () => {
  it("wraps the stacked-flank legend inside the canvas", () => {
    const cfg: ChartConfig = {
      kind: "butterfly",
      width: W,
      height: H,
      butterfly: { split: 2 },
      data: { categories: ["A", "B", "C"], series: regionSeries(5) },
    };
    // Before: legend-4 ("Latin America") spanned x 429.0 → 505.2.
    expect(offRight(buildChart(cfg).nodes, "legend")).toHaveLength(0);
  });

  it("floors the plot height instead of emitting negative-height bars", () => {
    const cfg: ChartConfig = {
      kind: "butterfly",
      width: 320,
      height: 48, // shorter than title + header + axis chrome
      title: "Headcount",
      data: {
        categories: ["Sales", "Ops", "Tech"],
        series: [
          { name: "2024", values: [40, 60, 30] },
          { name: "2025", values: [45, 55, 35] },
        ],
      },
      decorations: { segmentLabels: true, categoryAxis: true, valueAxis: true, seriesLabels: true },
    };
    // Before: every seg rect got h = -2.44 — dropped by SVG, clamped to a
    // 0.2pt sliver by Office.js, so preview and deck disagreed.
    const bad = buildChart(cfg).nodes.filter((n) => {
      const b = n as Partial<RectNode>;
      return (b.w != null && b.w < 0) || (b.h != null && b.h < 0);
    });
    expect(bad).toHaveLength(0);
  });
});
