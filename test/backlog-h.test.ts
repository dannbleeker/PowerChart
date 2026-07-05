import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, LineNode, RectNode, SceneNode, TextNode } from "../src/core/scene";

/**
 * Backlog batch H — more §2 within-kind gaps: area with negative values,
 * scatter/bubble trajectory trail, boxplot jittered raw-data dots.
 */

describe("area with negative values", () => {
  const areaRects = (sc: { nodes: SceneNode[] }) =>
    sc.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("area-"));
  const hasNegAxisLabel = (sc: { nodes: SceneNode[] }) =>
    sc.nodes.some((n) => n.kind === "text" && n.name === "value-axis" && /^-/.test((n.text ?? "").trim()));
  // The value axis "0" tick sits at toY(0) — the zero baseline in screen y.
  const zeroY = (sc: { nodes: SceneNode[] }) => {
    const z = sc.nodes.find((n): n is TextNode => n.kind === "text" && n.name === "value-axis" && (n.text ?? "").trim() === "0")!;
    return z.y + z.h / 2;
  };
  const mk = (values: (number | null)[], cats = ["A", "B", "C"]): ChartConfig => ({
    kind: "area",
    ...DEFAULT_SIZE,
    data: { categories: cats, series: [{ name: "v", values }] },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true },
  });

  it("fills both above and below the zero baseline", () => {
    const s = buildChart(mk([8, -8], ["A", "B"]));
    const y0 = zeroY(s);
    const rects = areaRects(s);
    expect(rects.some((r) => r.y + r.h > y0 + 1)).toBe(true); // fill extending below zero
    expect(rects.some((r) => r.y < y0 - 1)).toBe(true); // fill extending above zero
  });

  it("extends the value axis into the negatives (not clamped to 0)", () => {
    expect(hasNegAxisLabel(buildChart(mk([10, -6, 8])))).toBe(true);
  });

  it("positive-only area stays entirely above the baseline", () => {
    const pos = buildChart(mk([10, 6, 8]));
    expect(hasNegAxisLabel(pos)).toBe(false);
    const y0 = zeroY(pos);
    expect(areaRects(pos).every((r) => r.y + r.h <= y0 + 1)).toBe(true); // nothing below zero
  });
});

describe("scatter/bubble trajectory", () => {
  const base: ChartConfig = {
    kind: "scatter",
    ...DEFAULT_SIZE,
    data: {
      categories: ["t1", "t2", "t3", "t4"],
      series: [
        { name: "X", values: [10, 20, 30, 40] },
        { name: "Y", values: [10, 30, 20, 45] },
      ],
    },
  };

  it("adds no trail by default", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.name?.startsWith("trajectory"))).toBe(false);
  });

  it("connects points in row order with a direction arrowhead per segment", () => {
    const s = buildChart({ ...base, decorations: { trajectory: true } });
    const lines = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("trajectory-"));
    const heads = s.nodes.filter((n) => n.kind === "arrowhead" && n.name?.startsWith("trajectory-head-"));
    expect(lines).toHaveLength(3); // 4 points → 3 segments
    expect(heads).toHaveLength(3);
    // Segment 0 joins the first two points; x increases with X.
    expect(lines[0].x2).toBeGreaterThan(lines[0].x1);
    // Trail sits before the markers in draw order (behind them).
    const firstTraj = s.nodes.findIndex((n) => n.name === "trajectory-0");
    const firstPoint = s.nodes.findIndex((n) => n.name === "point-0");
    expect(firstTraj).toBeLessThan(firstPoint);
  });

  it("works for bubble charts too", () => {
    const s = buildChart({
      kind: "bubble",
      ...DEFAULT_SIZE,
      decorations: { trajectory: true },
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [1, 2, 3] },
          { name: "Y", values: [3, 1, 2] },
          { name: "Size", values: [10, 20, 30] },
        ],
      },
    });
    expect(s.nodes.filter((n) => n.name?.startsWith("trajectory-head-"))).toHaveLength(2);
  });
});

describe("boxplot jittered dots", () => {
  const rows = Array.from({ length: 12 }, (_, i) => ({
    name: `s${i + 1}`,
    values: [3 + (i % 5), 5 + (i % 4), 4 + (i % 6)],
  }));
  const base: ChartConfig = {
    kind: "boxplot",
    ...DEFAULT_SIZE,
    data: { categories: ["North", "South", "East"], series: rows },
    decorations: { categoryAxis: true, valueAxis: true },
  };

  it("adds no dots by default", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.name?.startsWith("dot-"))).toBe(false);
  });

  it("overlays one jittered dot per observation, spread across the box", () => {
    const s = buildChart({ ...base, boxplot: { jitter: true } });
    const dots = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("dot-0-"));
    expect(dots).toHaveLength(12); // one per raw observation in category 0
    // Deterministic jitter spreads them horizontally (not all on center).
    const xs = new Set(dots.map((d) => Math.round(d.cx * 10)));
    expect(xs.size).toBeGreaterThan(4);
    // Jitter subsumes the separate outlier dots.
    expect(s.nodes.some((n) => n.name?.startsWith("outlier-"))).toBe(false);
  });

  it("is deterministic (same layout twice)", () => {
    const a = buildChart({ ...base, boxplot: { jitter: true } }).nodes.filter((n) => n.name?.startsWith("dot-"));
    const b = buildChart({ ...base, boxplot: { jitter: true } }).nodes.filter((n) => n.name?.startsWith("dot-"));
    expect(JSON.stringify(a)).toBe(JSON.stringify(b));
  });
});
