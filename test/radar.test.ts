import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { EllipseNode, PolygonNode, RectNode, SceneNode } from "../src/core/scene";
import type { ChartConfig, Series } from "../src/core/types";

/** Radar — per-spoke scales, min–max bands, vertex markers, stacked, legend wrap. */

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
    const markers = s.nodes.filter(
      (n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.match(/^marker-\d+-\d+$/),
    );
    // One marker per (series × category); rendered by powerpoint.ts's ellipse
    // case, so they appear in the live add-in as well as SVG/pptx.
    expect(markers).toHaveLength(3);
    expect(markers.every((m) => m.fill && m.fill !== "none")).toBe(true);
  });
});

describe("stacked radar", () => {
  const cfg: ChartConfig = {
    kind: "radar",
    ...DEFAULT_SIZE,
    data: {
      categories: ["P", "Q", "R"],
      series: [
        { name: "A", values: [2, 3, 1] },
        { name: "B", values: [3, 2, 4] },
        { name: "C", values: [1, 1, 2] },
      ],
    },
    radar: { stacked: true },
  };
  const s = buildChart(cfg);

  it("draws one filled band polygon per series", () => {
    for (const si of [0, 1, 2]) {
      expect(s.nodes.some((n): n is PolygonNode => n.kind === "polygon" && n.name === `series-${si}`)).toBe(true);
    }
    // No overlaid single-value markers in stacked mode.
    expect(s.nodes.some((n) => n.name?.startsWith("marker-"))).toBe(false);
  });

  it("scales to the per-spoke sums so the outer band reaches the total", () => {
    // Distance from centre of the outermost band's first-spoke point should
    // exceed that of the innermost band (cumulative stacking).
    const cxu = DEFAULT_SIZE.width / 2;
    const inner = s.nodes.find((n): n is PolygonNode => n.name === "series-0")!;
    const outer = s.nodes.find((n): n is PolygonNode => n.name === "series-2")!;
    const rOf = (p: PolygonNode) => Math.abs(p.points[0].x - cxu) + Math.abs(p.points[0].y);
    expect(rOf(outer)).not.toBe(rOf(inner));
  });
});

describe("radar legends wrap instead of marching off the canvas", () => {
  it("wraps the polygon radar's legend and reserves the extra row", () => {
    const cfg: ChartConfig = {
      kind: "radar",
      width: W,
      height: H,
      data: { categories: ["Speed", "Cost", "Quality"], series: regionSeries(6) },
    };
    const { nodes } = buildChart(cfg);
    // Before: legend-5 ("Middle East") spanned x 476.4 → 552.6.
    expect(offRight(nodes, "legend")).toHaveLength(0);
    const chipRows = new Set(nodes.filter((n) => n.name?.startsWith("legend-chip-")).map((n) => (n as RectNode).y));
    expect(chipRows.size).toBe(2);
  });

  it("wraps the radial-bar (coxcomb) legend too", () => {
    const cfg: ChartConfig = {
      kind: "radar",
      width: W,
      height: H,
      radar: { bars: true },
      data: { categories: ["Speed", "Cost", "Quality"], series: regionSeries(6) },
    };
    expect(offRight(buildChart(cfg).nodes, "legend")).toHaveLength(0);
  });

  it("keeps the band-mode swatch names while wrapping", () => {
    const cfg: ChartConfig = {
      kind: "radar",
      width: W,
      height: H,
      data: { categories: ["Speed", "Cost", "Quality"], series: regionSeries(4) },
      decorations: { radarBand: true },
    };
    const { nodes } = buildChart(cfg);
    expect(nodes.some((n) => n.name === "legend-band")).toBe(true);
    expect(nodes.some((n) => n.name === "legend-us")).toBe(true);
    expect(offRight(nodes, "legend")).toHaveLength(0);
  });
});

describe("radar perSpoke normalises against each spoke's own maximum", () => {
  it("puts a sub-1 spoke maximum on the rim", () => {
    const cfg: ChartConfig = {
      kind: "radar",
      width: W,
      height: H,
      radar: { perSpoke: true },
      // A rate expressed as 0–1 beside counts — the case perSpoke exists for.
      data: { categories: ["Revenue (m)", "Conversion rate", "NPS"], series: [{ name: "Us", values: [3, 0.5, 40] }] },
    };
    const { nodes } = buildChart(cfg);
    const rim = nodes.filter((n) => n.name?.startsWith("grid-")).at(-1) as { points: { x: number; y: number }[] };
    const shape = nodes.find((n) => n.name === "series-0") as { points: { x: number; y: number }[] };
    // Before: the 0.5 spoke normalised against the Math.max(1, …) floor and
    // drew at half the rim radius.
    shape.points.forEach((p, i) => {
      expect(p.x).toBeCloseTo(rim.points[i].x, 6);
      expect(p.y).toBeCloseTo(rim.points[i].y, 6);
    });
  });
});
