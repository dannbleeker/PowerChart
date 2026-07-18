import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { PolygonNode, RectNode } from "../src/core/scene";

/**
 * Backlog batch J — more §2 within-kind gaps: gantt auto-summary bars,
 * notched boxplots, radar min–max (peer range + us) band.
 */

describe("gantt auto-summary bars", () => {
  const cfg: ChartConfig = {
    kind: "gantt",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Phase 1", "> Research", "> Interviews", "Phase 2", "> Build"],
      series: [
        { name: "Start", values: [null, 1, 3, null, 6] },
        { name: "End", values: [null, 4, 8, null, 12] },
      ],
    },
    decorations: { summaryBars: true, segmentLabels: false },
  };

  it("draws a capped summary bar spanning the section's children", () => {
    const s = buildChart(cfg);
    const summary = s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "summary-0");
    expect(summary).toBeTruthy();
    // Children of Phase 1 are rows 1–2: min start = 1, max end = 8.
    const r1 = s.nodes.find((n): n is RectNode => n.name === "bar-1")!;
    const r2 = s.nodes.find((n): n is RectNode => n.name === "bar-2")!;
    expect(summary!.x).toBeCloseTo(r1.x, 1); // starts at the earliest child start
    expect(summary!.x + summary!.w).toBeCloseTo(r2.x + r2.w, 1); // ends at the latest child end
    expect(s.nodes.some((n) => n.name === "summary-cap-a-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "summary-cap-b-0")).toBe(true);
  });

  it("no summary bars without the decoration", () => {
    const s = buildChart({ ...cfg, decorations: { segmentLabels: false } });
    expect(s.nodes.some((n) => n.name?.startsWith("summary-"))).toBe(false);
  });
});

describe("notched boxplots", () => {
  const rows = Array.from({ length: 10 }, (_, i) => ({
    name: `o${i + 1}`,
    values: [10 + i, 12 + i * 1.5, 20 + i * 2],
  }));
  const base: ChartConfig = {
    kind: "boxplot",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: rows },
    decorations: { categoryAxis: true, valueAxis: true },
  };

  it("renders the box as a 10-point notched polygon in raw-sample mode", () => {
    const s = buildChart({ ...base, boxplot: { notch: true } });
    const box = s.nodes.find((n): n is PolygonNode => n.kind === "polygon" && n.name === "box-0");
    expect(box).toBeTruthy();
    expect(box!.points).toHaveLength(10);
    // The notch pinches inward at the median: mid-height x-extent < box extent.
    const xs = box!.points.map((p) => p.x);
    const boxExtent = Math.max(...xs) - Math.min(...xs);
    const notchXs = box!.points.filter((_, i) => i === 3 || i === 8).map((p) => p.x);
    expect(Math.abs(notchXs[0] - notchXs[1])).toBeLessThan(boxExtent);
  });

  it("plain box (rect) without the notch flag", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.kind === "rect" && n.name === "box-0")).toBe(true);
    expect(s.nodes.some((n) => n.kind === "polygon" && n.name === "box-0")).toBe(false);
  });

  it("precomputed boxes (no sample size) stay rectangular even with notch on", () => {
    const s = buildChart({
      kind: "boxplot",
      ...DEFAULT_SIZE,
      boxplot: { notch: true },
      data: {
        categories: ["X"],
        series: [
          { name: "Min", values: [2] },
          { name: "Q1", values: [4] },
          { name: "Median", values: [6] },
          { name: "Q3", values: [9] },
          { name: "Max", values: [12] },
        ],
      },
    });
    expect(s.nodes.some((n) => n.kind === "rect" && n.name === "box-0")).toBe(true);
  });
});

describe("radar min–max band (peer range + us)", () => {
  const cfg: ChartConfig = {
    kind: "radar",
    ...DEFAULT_SIZE,
    scale: { min: 0, max: 5 },
    data: {
      categories: ["A", "B", "C", "D"],
      series: [
        { name: "Peer 1", values: [3, 2, 4, 3] },
        { name: "Peer 2", values: [2, 3, 3, 4] },
        { name: "Us", values: [4, 4, 2, 5] },
      ],
    },
    decorations: { seriesLabels: true, radarBand: true },
  };

  it("shades the peer envelope and draws only the last series", () => {
    const s = buildChart(cfg);
    expect(s.nodes.some((n) => n.name?.startsWith("band-"))).toBe(true);
    expect(s.nodes.some((n) => n.name === "band-max")).toBe(true);
    expect(s.nodes.some((n) => n.name === "band-min")).toBe(true);
    // Only "us" (index 2) is drawn as a series polygon; peers are the band.
    expect(s.nodes.some((n) => n.name === "series-2")).toBe(true);
    expect(s.nodes.some((n) => n.name === "series-0")).toBe(false);
    expect(s.nodes.some((n) => n.name === "series-1")).toBe(false);
    // Legend collapses to "Peer range" + us.
    expect(s.nodes.some((n) => n.name === "legend-band")).toBe(true);
    expect(s.nodes.some((n) => n.name === "legend-us")).toBe(true);
  });

  it("draws one band quad per spoke plus min/max envelopes", () => {
    const s = buildChart(cfg);
    const quads = s.nodes.filter((n): n is PolygonNode => n.kind === "polygon" && /^band-\d+$/.test(n.name ?? ""));
    expect(quads).toHaveLength(4); // one per category/spoke
    expect(quads.every((q) => q.points.length === 4)).toBe(true);
    const max = s.nodes.find((n): n is PolygonNode => n.name === "band-max")!;
    const min = s.nodes.find((n): n is PolygonNode => n.name === "band-min")!;
    // Envelopes trace every spoke; max further from centre → larger bounding box.
    const bbox = (p: PolygonNode) => {
      const xs = p.points.map((q) => q.x);
      const ys = p.points.map((q) => q.y);
      return (Math.max(...xs) - Math.min(...xs)) * (Math.max(...ys) - Math.min(...ys));
    };
    expect(max.points).toHaveLength(4);
    expect(bbox(max)).toBeGreaterThan(bbox(min));
  });

  it("without radarBand, all series draw and no band appears", () => {
    const s = buildChart({ ...cfg, decorations: { seriesLabels: true } });
    expect(s.nodes.some((n) => n.name === "series-0")).toBe(true);
    expect(s.nodes.some((n) => n.name?.startsWith("band-"))).toBe(false);
  });
});
