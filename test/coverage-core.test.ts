import { describe, expect, it } from "vitest";
import { buildChart, valueExtent, DEFAULT_SIZE } from "../src/core/chart";
import { isTotalToken } from "../src/core/layout/waterfall";
import { contrastInk, type Scene, type SceneNode, type TextNode } from "../src/core/scene";
import { resolveLabelCollisions } from "../src/core/collide";
import { toRgb, toHex6, alphaOf } from "../src/core/color";
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

  it("reads rgb()/hsl() fills, not as black", () => {
    // A hex-only parser returned NaN->0 (pure black) for functional colours, so
    // a near-white rgb() fill wrongly got WHITE ink. These forms are valid config
    // (Series.color etc. are plain strings the renderer's PAINT_OK admits).
    expect(contrastInk("rgb(250,250,250)")).toBe(contrastInk("#fafafa"));
    expect(contrastInk("rgb(250,250,250)")).toBe("#0b0b0b"); // dark ink on near-white
    expect(contrastInk("rgb(20,20,20)")).toBe("#ffffff"); // white ink on near-black
    expect(contrastInk("hsl(0,0%,100%)")).toBe("#0b0b0b"); // hsl white
  });
});

describe("toHex6 / alphaOf normalize colours for the PowerPoint renderers", () => {
  it("normalizes every allow-listed form to a 6-digit hex", () => {
    expect(toHex6("#4e79a7")).toBe("#4e79a7"); // identity for the hex the engine emits
    expect(toHex6("#abc")).toBe("#aabbcc");
    expect(toHex6("#4e79a780")).toBe("#4e79a7"); // alpha byte dropped
    expect(toHex6("rgb(78,121,167)")).toBe("#4e79a7");
    expect(toHex6("hsl(0,0%,100%)")).toBe("#ffffff");
    expect(toHex6("not-a-color")).toBe("#808080"); // named/unknown → grey, never black
  });

  it("reads the opacity a paint carries, 1 when opaque", () => {
    expect(alphaOf("#4e79a7")).toBe(1);
    expect(alphaOf("#4e79a780")).toBeCloseTo(128 / 255, 5);
    expect(alphaOf("rgba(1,2,3,0.5)")).toBe(0.5);
    expect(alphaOf("hsla(0,0%,0%,0.25)")).toBe(0.25);
    expect(alphaOf("rgb(1,2,3)")).toBe(1);
  });
});

describe("toRgb parses every allow-listed colour form", () => {
  it("matches hex for the equivalent rgb()/hsl(), strips alpha, expands short hex", () => {
    expect(toRgb("rgb(78,121,167)")).toEqual(toRgb("#4e79a7"));
    expect(toRgb("#abc")).toEqual(toRgb("#aabbcc"));
    expect(toRgb("#4e79a780")).toEqual(toRgb("#4e79a7")); // alpha byte dropped
    expect(toRgb("rgb(50%,50%,50%)")).toEqual([127, 127, 127]);
    expect(toRgb("hsl(0,0%,0%)")).toEqual([0, 0, 0]);
    // A malformed paint falls back to mid grey, never NaN/black.
    expect(toRgb("not-a-color")).toEqual([128, 128, 128]);
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
  it("nudges a movable total up off a fixed label when there is room", () => {
    const nodes: SceneNode[] = [
      {
        kind: "text",
        x: 0,
        y: 40,
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
        y: 40,
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
    expect(total.y).toBeLessThan(40); // moved up, clear of the fixed label
    expect(total.y).toBeGreaterThanOrEqual(0); // but still on the canvas
  });

  it("does not nudge a movable total off the top of the canvas", () => {
    // A total pinned against the top with no room (e.g. sharing the totals row
    // with the fixed grand-total label) must NOT escape the canvas: an
    // overlapping label reads, an off-canvas one is lost. It gives up in place.
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
    expect((nodes[1] as TextNode).y).toBeGreaterThanOrEqual(0);
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
