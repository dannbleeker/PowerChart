import { describe, expect, it } from "vitest";
import { buildChart, valueExtent, DEFAULT_SIZE } from "../src/core/chart";
import { isTotalToken } from "../src/core/layout/waterfall";
import { contrastInk, type Scene, type SceneNode, type TextNode } from "../src/core/scene";
import { resolveLabelCollisions } from "../src/core/collide";
import { sceneToSvg } from "../src/render/svg";
import type { ChartConfig, ChartData } from "../src/core/types";

const data2 = (a: number[], b: number[], cats = ["A", "B", "C"]): ChartData => ({
  categories: cats.slice(0, a.length),
  series: [
    { name: "S1", values: a },
    { name: "S2", values: b },
  ],
});

const cfg = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  data: data2([10, 20, 30], [5, 5, 5]),
  ...DEFAULT_SIZE,
  ...partial,
});

const texts = (s: Scene) => s.nodes.filter((n): n is TextNode => n.kind === "text");

describe("waterfall total token", () => {
  it("accepts e, =, and Σ in any case with padding", () => {
    for (const t of ["e", "E", " e ", "=", "Σ", "σ"]) expect(isTotalToken(t), t).toBe(true);
  });
  it("rejects data-looking cells", () => {
    for (const t of ["", "12", "e2", "==", "sum"]) expect(isTotalToken(t), t).toBe(false);
  });
});

describe("valueExtent (Same Scale)", () => {
  it("stacked sums positives and negatives per category", () => {
    const e = valueExtent(cfg({ data: data2([10, -4, 30], [5, -6, 5]) }));
    expect(e).toEqual({ min: -10, max: 35 });
  });
  it("clustered/line use the raw value range", () => {
    expect(valueExtent(cfg({ kind: "clustered" }))).toEqual({ min: 0, max: 30 });
    expect(valueExtent(cfg({ kind: "line", data: data2([-2, 8, 4], [1, 1, 1]) }))).toEqual({ min: -2, max: 8 });
  });
  it("area stacks positive values from zero", () => {
    expect(valueExtent(cfg({ kind: "area" }))).toEqual({ min: 0, max: 35 });
  });
  it("waterfall tracks the running level", () => {
    const c: ChartConfig = cfg({
      kind: "waterfall",
      data: { categories: ["Start", "Up", "Down", "End"], series: [{ name: "S", values: [50, 20, -30, 0] }] },
      waterfall: { totalIndices: [3] },
    });
    expect(valueExtent(c)).toEqual({ min: 0, max: 70 });
  });
  it("returns null when there is nothing to measure", () => {
    expect(valueExtent(cfg({ data: { categories: [], series: [] } }))).toBeNull();
    expect(valueExtent(cfg({ kind: "pie" }))).toBeNull();
  });
});

describe("horizontal bar chrome", () => {
  it("draws vertical gridlines and a bottom value axis", () => {
    const s = buildChart(
      cfg({
        horizontal: true,
        decorations: { gridlines: true, valueAxis: true, categoryAxis: true, segmentLabels: true },
      }),
    );
    const grid = s.nodes.filter((n) => n.name === "gridline");
    expect(grid.length).toBeGreaterThan(1);
    // Horizontal chart gridlines are vertical strokes spanning the plot.
    for (const g of grid) if (g.kind === "line") expect(g.x1).toBe(g.x2);
    expect(texts(s).some((t) => t.name === "value-axis")).toBe(true);
    expect(texts(s).some((t) => t.name?.startsWith("category-"))).toBe(true);
  });
});

describe("date-spaced line categories", () => {
  it("spaces category centers proportionally to time", () => {
    const c = cfg({
      kind: "line",
      data: {
        categories: ["2025-01", "2025-02", "2025-12"],
        series: [{ name: "S", values: [1, 2, 3] }],
      },
    });
    const s = buildChart(c);
    const pts = s.nodes.filter((n) => n.name?.startsWith("marker-0-"));
    expect(pts).toHaveLength(3);
    const xs = pts.map((p) => (p.kind === "rect" ? p.x : 0));
    // Jan→Feb gap must be far smaller than Feb→Dec.
    expect(xs[1] - xs[0]).toBeLessThan((xs[2] - xs[1]) / 5);
  });
});

describe("contrastInk", () => {
  it("expands 3-digit hex and picks readable ink", () => {
    expect(contrastInk("#fff")).toBe(contrastInk("#ffffff"));
    expect(contrastInk("#000")).not.toBe(contrastInk("#fff"));
  });
});

describe("difference arrow anchored to a value line", () => {
  const base = {
    from: 0,
    to: 2,
  };
  it("anchors at a mean value line", () => {
    const s = buildChart(
      cfg({
        decorations: {
          segmentLabels: true,
          valueLines: [{ mode: "mean" }],
          difference: { ...base, fromValueLine: 0 },
        },
      }),
    );
    expect(texts(s).some((t) => t.name === "diff-label")).toBe(true);
  });
  it("anchors at a fixed value line", () => {
    const s = buildChart(
      cfg({
        decorations: {
          segmentLabels: true,
          valueLines: [{ mode: "value", value: 20 }],
          difference: { ...base, fromValueLine: 0 },
        },
      }),
    );
    expect(texts(s).some((t) => t.name === "diff-label")).toBe(true);
  });
  it("normalizes the deprecated single valueLine field", () => {
    const s = buildChart(
      cfg({
        decorations: {
          segmentLabels: true,
          valueLine: { mode: "mean" },
          difference: { ...base, fromValueLine: 0 },
        },
      }),
    );
    expect(s.nodes.some((n) => n.name?.startsWith("value-line"))).toBe(true);
  });
});

describe("label collision resolution", () => {
  it("ignores unnamed and immovable nodes", () => {
    const overlapped: SceneNode[] = [
      {
        kind: "text",
        x: 0,
        y: 0,
        w: 40,
        h: 12,
        text: "a",
        fontSize: 10,
        color: "#000",
        align: "center",
        valign: "middle",
      },
      {
        kind: "text",
        x: 0,
        y: 0,
        w: 40,
        h: 12,
        text: "b",
        fontSize: 10,
        color: "#000",
        align: "center",
        valign: "middle",
        name: "segment-label-0-0",
      },
    ];
    const before = JSON.stringify(overlapped);
    resolveLabelCollisions(overlapped);
    expect(JSON.stringify(overlapped)).toBe(before);
  });
  it("nudges a movable total off a fixed label", () => {
    const nodes: SceneNode[] = [
      {
        kind: "text",
        x: 0,
        y: 0,
        w: 40,
        h: 12,
        text: "fixed",
        fontSize: 10,
        color: "#000",
        align: "center",
        valign: "middle",
        name: "segment-label-0-0",
      },
      {
        kind: "text",
        x: 0,
        y: 0,
        w: 40,
        h: 12,
        text: "42",
        fontSize: 10,
        color: "#000",
        align: "center",
        valign: "middle",
        name: "total-0",
      },
    ];
    resolveLabelCollisions(nodes);
    const total = nodes[1] as TextNode;
    expect(total.y).toBeLessThan(0);
  });
});

describe("SVG renderer options", () => {
  it("paints an explicit background", () => {
    const svg = sceneToSvg(buildChart(cfg({})), { background: "#ffffff" });
    expect(svg).toContain('fill="#ffffff"');
  });
  it("emits annular paths for doughnut wedges", () => {
    const c = cfg({
      kind: "doughnut",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
    });
    const svg = sceneToSvg(buildChart(c));
    // Two arcs per annular wedge: outer radius sweep + inner return.
    expect(svg.match(/<path/g)!.length).toBeGreaterThanOrEqual(2);
  });
});
