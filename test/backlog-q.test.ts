import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, PolygonNode, RectNode, WedgeNode } from "../src/core/scene";

/**
 * Batch Q — three new kinds (owner-requested, previously in §2/§4): sunburst,
 * violin, candlestick.
 */

describe("sunburst", () => {
  const cfg: ChartConfig = {
    kind: "sunburst",
    ...DEFAULT_SIZE,
    data: {
      categories: ["G1 | A", "G1 | B", "G2 | C"],
      series: [{ name: "V", values: [30, 10, 40] }],
    },
    decorations: { segmentLabels: false },
  };
  const s = buildChart(cfg);

  it("nests items on an outer ring inside group wedges on the inner ring", () => {
    const g0 = s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === "group-0")!;
    const item = s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === "slice-0")!;
    expect(g0).toBeTruthy();
    expect(item).toBeTruthy();
    // Inner ring (groups) sits inside the outer ring (items).
    expect(g0.r).toBeLessThan(item.r);
    expect(item.innerR).toBeGreaterThanOrEqual(g0.r - 0.01);
    // Group spans are proportional: G1 (40) vs G2 (40) → equal; item A (30) > B (10).
    const a = s.nodes.find((n): n is WedgeNode => n.name === "slice-0")!;
    const b = s.nodes.find((n): n is WedgeNode => n.name === "slice-1")!;
    expect(a.endAngle - a.startAngle).toBeGreaterThan(b.endAngle - b.startAngle);
  });

  it("flat data (no groups) makes a single-ring doughnut", () => {
    const flat = buildChart({ ...cfg, data: { categories: ["A", "B"], series: [{ name: "V", values: [1, 1] }] } });
    expect(flat.nodes.some((n) => n.name?.startsWith("group-"))).toBe(false);
    const slices = flat.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    expect(slices).toHaveLength(2);
    expect(slices.every((w) => w.innerR > 0)).toBe(true);
  });
});

describe("violin", () => {
  const rows = Array.from({ length: 10 }, (_, i) => ({
    name: `s${i + 1}`,
    values: [100 + i, 200 + i * 3, 500 + i * 8],
  }));
  const cfg: ChartConfig = {
    kind: "violin",
    ...DEFAULT_SIZE,
    data: { categories: ["Tight", "Mid", "Wide"], series: rows },
    decorations: { categoryAxis: true, valueAxis: true },
  };
  const s = buildChart(cfg);

  it("draws a density polygon and a median tick per category", () => {
    [0, 1, 2].forEach((c) => {
      expect(s.nodes.some((n) => n.kind === "polygon" && n.name === `violin-${c}`)).toBe(true);
      expect(s.nodes.some((n) => n.kind === "line" && n.name === `median-${c}`)).toBe(true);
    });
  });

  it("wider-spread data produces a taller violin", () => {
    const height = (c: number) => {
      const p = s.nodes.find((n): n is PolygonNode => n.name === `violin-${c}`)!;
      const ys = p.points.map((q) => q.y);
      return Math.max(...ys) - Math.min(...ys);
    };
    // "Wide" (500–572) spans more of the shared axis than "Tight" (100–109).
    expect(height(2)).toBeGreaterThan(height(0));
  });
});

describe("candlestick", () => {
  const cfg: ChartConfig = {
    kind: "candlestick",
    ...DEFAULT_SIZE,
    data: {
      categories: ["D1", "D2"],
      series: [
        { name: "Open", values: [40, 46] },
        { name: "High", values: [48, 47] },
        { name: "Low", values: [39, 42] },
        { name: "Close", values: [46, 43] }, // D1 rises, D2 falls
      ],
    },
  };
  const s = buildChart(cfg);

  it("draws a high–low wick and an open–close body per period", () => {
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && n.name === "wick-0")).toBe(true);
    expect(s.nodes.some((n): n is RectNode => n.kind === "rect" && n.name === "body-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "body-1")).toBe(true);
  });

  it("colours rising periods green and falling periods red", () => {
    const up = s.nodes.find((n): n is RectNode => n.name === "body-0")!; // close 46 > open 40
    const down = s.nodes.find((n): n is RectNode => n.name === "body-1")!; // close 43 < open 46
    expect(up.fill).not.toBe(down.fill);
    expect(up.fill).toBe("#1a9e6e"); // green
  });

  it("the wick spans the full high–low range", () => {
    const wick = s.nodes.find((n): n is LineNode => n.name === "wick-0")!;
    // High (48) maps above Low (39): y1 (high) < y2 (low) on screen.
    expect(wick.y1).toBeLessThan(wick.y2);
  });
});
