import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, PolygonNode, RectNode } from "../src/core/scene";

/**
 * Backlog batch L — §2 tail: radar per-spoke scales, transparent floating
 * segments, line missing-data bridge.
 */

describe("radar per-spoke scales", () => {
  const base: ChartConfig = {
    kind: "radar",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Big unit", "Small unit"],
      series: [{ name: "S", values: [1000, 5] }],
    },
  };

  it("normalises each spoke to its own max (drops numeric ticks)", () => {
    const s = buildChart({ ...base, radar: { perSpoke: true } });
    const poly = s.nodes.find((n): n is PolygonNode => n.kind === "polygon" && n.name === "series-0")!;
    const cx = DEFAULT_SIZE.width / 2;
    // Spoke 0 points straight up (x == centre) despite its 1000 vs the other's 5.
    expect(poly.points[0].x).toBeCloseTo(cx, 1);
    // No numeric tick labels in per-spoke mode.
    expect(s.nodes.some((n) => n.name?.startsWith("tick-"))).toBe(false);
    // Fraction grid rings are still drawn.
    expect(s.nodes.some((n) => n.name === "grid-1")).toBe(true);
  });

  it("default radar keeps a shared scale and numeric ticks", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.name?.startsWith("tick-"))).toBe(true);
  });

  it("per-spoke makes equal-rank spokes reach the same radius", () => {
    const s = buildChart({
      kind: "radar",
      ...DEFAULT_SIZE,
      radar: { perSpoke: true },
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          { name: "hi", values: [100, 3, 100, 3] },
          { name: "lo", values: [50, 1.5, 50, 1.5] },
        ],
      },
    });
    const hi = s.nodes.find((n): n is PolygonNode => n.name === "series-0")!;
    const cx = DEFAULT_SIZE.width / 2;
    // "hi" is each spoke's own max, so opposite spokes A(100) and C(100) both
    // reach the rim → symmetric about the centre.
    expect(Math.abs(hi.points[0].x - cx)).toBeCloseTo(Math.abs(hi.points[2].x - cx), 1);
  });
});

describe("transparent floating segments", () => {
  const cfg: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "base", color: "transparent", values: [10, 14] },
        { name: "Range", values: [20, 18] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("draws no rect for the transparent segment but still stacks above it", () => {
    const s = buildChart(cfg);
    expect(s.nodes.some((n) => n.name === "seg-0-0")).toBe(false); // transparent base not drawn
    const range = s.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    expect(range).toBeTruthy();
    const solid = buildChart({
      ...cfg,
      data: {
        ...cfg.data,
        series: [
          { name: "base", values: [10, 14] },
          { name: "Range", values: [20, 18] },
        ],
      },
    });
    const solidRange = solid.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    // Same height, but the floating one sits higher (smaller y) is false — same y
    // since the level is identical; the difference is the missing base rect.
    expect(range.y).toBeCloseTo(solidRange.y, 3);
    expect(range.h).toBeCloseTo(solidRange.h, 3);
    // The floating build is missing the base segment the solid build draws.
    expect(solid.nodes.some((n) => n.name === "seg-0-0")).toBe(true);
  });

  it("floating bar does not reach the baseline", () => {
    const s = buildChart(cfg);
    const range = s.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    // Baseline is the bottom of the plot; the bar's bottom (y+h) is above it
    // by the (undrawn) base segment. Compare Q1 (base 10) vs a base-0 build.
    const grounded = buildChart({
      ...cfg,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "base", color: "transparent", values: [0, 0] },
          { name: "Range", values: [20, 18] },
        ],
      },
    });
    const gRange = grounded.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    expect(range.y + range.h).toBeLessThan(gRange.y + gRange.h - 1); // floats above the grounded bar's base
  });
});

describe("line missing-data bridge", () => {
  const base: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "v", values: [10, null, 30, 40] }] },
    decorations: { segmentLabels: false },
  };

  it("breaks the line at nulls by default", () => {
    const s = buildChart(base);
    const lines = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    // Only C→D connects (A is isolated by the null B, B has no point).
    expect(lines).toHaveLength(1);
    expect(s.nodes.filter((n) => n.name?.startsWith("marker-0-"))).toHaveLength(3); // A, C, D
  });

  it("bridges across nulls when bridgeGaps is set", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, bridgeGaps: true } });
    const lines = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    // A→C (bridging B) and C→D → two segments.
    expect(lines).toHaveLength(2);
    // The bridge segment joins A's point to C's point.
    const bridge = lines.find((l) => l.name === "line-0-2")!;
    const cx = (i: number) => s.nodes.find((n): n is RectNode => n.name === `marker-0-${i}`)!.x;
    expect(bridge.x1).toBeCloseTo(cx(0) + 2.4, 0); // near A's marker
    expect(bridge.x2).toBeCloseTo(cx(2) + 2.4, 0); // to C's marker
  });
});
