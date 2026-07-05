import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, RectNode } from "../src/core/scene";

/**
 * Backlog batch P — the deferred §2-tail items: horizontal profile chart
 * (line/area) and radar vertex markers (verified in the scene → add-in path).
 */

describe("horizontal profile chart — line", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    horizontal: true,
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [20, 80, 50] }] },
    decorations: { segmentLabels: false, categoryAxis: true, valueAxis: true },
  };
  const s = buildChart(cfg);
  const marker = (c: number) => s.nodes.find((n): n is RectNode => n.name === `marker-0-${c}`)!;

  it("runs categories down the y axis (top to bottom)", () => {
    expect(marker(0).y).toBeLessThan(marker(1).y);
    expect(marker(1).y).toBeLessThan(marker(2).y);
  });

  it("maps larger values further right (value along x)", () => {
    expect(marker(1).x).toBeGreaterThan(marker(0).x); // 80 right of 20
    expect(marker(2).x).toBeGreaterThan(marker(0).x); // 50 right of 20
    expect(marker(1).x).toBeGreaterThan(marker(2).x); // 80 right of 50
  });

  it("connects the points with line segments", () => {
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
  });

  it("the vertical line chart is unaffected (no horizontal dispatch)", () => {
    const v = buildChart({ ...cfg, horizontal: undefined });
    // In the vertical chart categories run along x, so markers share no single column.
    const m0 = v.nodes.find((n): n is RectNode => n.name === "marker-0-0")!;
    const m1 = v.nodes.find((n): n is RectNode => n.name === "marker-0-1")!;
    expect(m1.x).toBeGreaterThan(m0.x); // categories advance along x when vertical
  });
});

describe("horizontal profile chart — stacked area", () => {
  const cfg: ChartConfig = {
    kind: "area",
    ...DEFAULT_SIZE,
    horizontal: true,
    data: {
      categories: ["N", "S"],
      series: [
        { name: "A", values: [40, 55] },
        { name: "B", values: [25, 20] },
      ],
    },
    decorations: { seriesLabels: true, categoryAxis: true, valueAxis: true },
  };
  const s = buildChart(cfg);

  it("stacks the series along x (second series to the right of the first)", () => {
    // Compare the same y-strip (slab k=0) of each series: B stacks to the
    // right of A there (globally the x-ranges overlap across categories).
    const a0 = s.nodes.find((n): n is RectNode => n.name === "area-0-0-0")!;
    const b0 = s.nodes.find((n): n is RectNode => n.name === "area-1-0-0")!;
    expect(a0).toBeTruthy();
    expect(b0).toBeTruthy();
    expect(b0.x).toBeGreaterThanOrEqual(a0.x + a0.w - 1);
  });
});

describe("radar vertex markers (scene → add-in path)", () => {
  it("emits a filled ellipse marker at every vertex", () => {
    const s = buildChart({
      kind: "radar",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B", "C"],
        series: [{ name: "S", values: [3, 4, 2] }],
      },
    });
    const markers = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.match(/^marker-\d+-\d+$/));
    // One marker per (series × category); rendered by powerpoint.ts's ellipse
    // case, so they appear in the live add-in as well as SVG/pptx.
    expect(markers).toHaveLength(3);
    expect(markers.every((m) => m.fill && m.fill !== "none")).toBe(true);
  });
});
