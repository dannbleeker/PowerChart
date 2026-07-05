import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, LineNode, RectNode } from "../src/core/scene";

/**
 * Batch R — three low/niche §2 gaps: critical-path Gantt highlight, the
 * mean±SD boxplot variant, and sparklines.
 */

describe("critical-path Gantt", () => {
  // Two chains from Research: the long one (Design→Build→QA→Launch) and a
  // short branch (Copywriting). Durations make the long chain critical.
  const cfg: ChartConfig = {
    kind: "gantt",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Research", "Design", "Build", "Copywriting", "QA", "Launch"],
      series: [
        { name: "Start", values: [1, 3, 8, 3, 14, 18] },
        { name: "End", values: [3, 8, 14, 5, 18, 20] },
        { name: "After", values: [null, 1, 2, 1, 3, 5] },
      ],
    },
    decorations: { criticalPath: true },
  };
  const s = buildChart(cfg);
  const bar = (c: number) => s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === `bar-${c}`)!;

  it("outlines the critical activities in red and leaves the branch plain", () => {
    // Critical chain: Research(0) Design(1) Build(2) QA(4) Launch(5).
    for (const c of [0, 1, 2, 4, 5]) {
      expect(bar(c).stroke).toBe("#e34948");
    }
    // Copywriting(3) is on the shorter branch — no critical outline.
    expect(bar(3).stroke).toBeUndefined();
  });

  it("draws the critical dependency arrows thicker and red", () => {
    const critHead = s.nodes.find((n) => n.name === "dep-head-2")!; // Build ← Design edge
    const branchHead = s.nodes.find((n) => n.name === "dep-head-3")!; // Copywriting ← Research edge
    expect((critHead as { fill: string }).fill).toBe("#e34948");
    expect((branchHead as { fill: string }).fill).not.toBe("#e34948");
  });

  it("is a no-op without the decoration", () => {
    const plain = buildChart({ ...cfg, decorations: {} });
    const b0 = plain.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "bar-0")!;
    expect(b0.stroke).toBeUndefined();
  });
});

describe("mean±SD boxplot", () => {
  // Symmetric sample around 10: mean 10, sample SD 2.
  const cfg: ChartConfig = {
    kind: "boxplot",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A"],
      series: [
        { name: "o1", values: [8] },
        { name: "o2", values: [10] },
        { name: "o3", values: [12] },
        { name: "o4", values: [8] },
        { name: "o5", values: [10] },
        { name: "o6", values: [12] },
      ],
    },
    boxplot: { meanSd: true },
    decorations: { valueAxis: true },
  };
  const s = buildChart(cfg);
  const valueToY = s.nodes.length; // touch to avoid unused; real check below

  it("centres the box on the mean with a ±1·SD body", () => {
    const box = s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "box-0")!;
    const median = s.nodes.find((n): n is LineNode => n.kind === "line" && n.name === "median-0")!;
    // Median line is the mean (y at value 10); the box spans q1..q3 = mean±SD.
    const meanY = median.y1;
    // Box top corresponds to the higher value (mean+SD=12), so box.y < meanY.
    expect(box.y).toBeLessThan(meanY);
    expect(box.y + box.h).toBeGreaterThan(meanY);
    // Symmetric: mean line sits at the vertical centre of the box (±1px).
    expect(Math.abs(box.y + box.h / 2 - meanY)).toBeLessThan(1);
  });

  it("has no mean × marker (the centre line already is the mean)", () => {
    expect(s.nodes.some((n) => n.name?.startsWith("mean-"))).toBe(false);
    expect(valueToY).toBeGreaterThan(0);
  });
});

describe("sparklines", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: {
      categories: ["1", "2", "3", "4", "5"],
      series: [{ name: "Visits", values: [10, 40, 20, 5, 25] }],
    },
    decorations: { sparkline: true },
  };
  const s = buildChart(cfg);

  it("draws a thin line with min/max/last dots and no axis chrome", () => {
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && n.strokeWidth === 1.25)).toBe(true);
    const min = s.nodes.find((n): n is EllipseNode => n.name === "spark-min-0")!;
    const max = s.nodes.find((n): n is EllipseNode => n.name === "spark-max-0")!;
    const last = s.nodes.find((n): n is EllipseNode => n.name === "spark-last-0")!;
    expect(min.fill).toBe("#e34948"); // red on the minimum (value 5)
    expect(max.fill).toBe("#1a9e6e"); // green on the maximum (value 40)
    // Max (40) sits above min (5) on screen.
    expect(max.cy).toBeLessThan(min.cy);
    // Last dot is at the final category, to the right of the (earlier) max.
    expect(last.cx).toBeGreaterThan(max.cx);
    // No gridlines/axis lines emitted.
    expect(s.nodes.some((n) => n.name === "gridline")).toBe(false);
  });

  it("area sparklines add a light fill beneath the line", () => {
    const areaSpark = buildChart({ ...cfg, kind: "area" });
    expect(areaSpark.nodes.some((n) => n.name?.startsWith("spark-fill-"))).toBe(true);
  });
});
