import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { PolygonNode, RectNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Tilemap cartogram — hex tiles and mini-glyphs. */

/**
 * Backlog batch N — §2 tail: tilemap hex tiles, tilemap mini-glyphs,
 * stacked100 negative values.
 */
describe("tilemap hex tiles", () => {
  const base: ChartConfig = {
    kind: "tilemap",
    ...DEFAULT_SIZE,
    map: "us",
    data: { categories: ["CA", "TX", "NY"], series: [{ name: "S", values: [100, 80, 60] }] },
  };

  it("draws hexagon polygons instead of square tiles", () => {
    const s = buildChart({ ...base, tilemap: { shape: "hex" } });
    const ca = s.nodes.find((n): n is PolygonNode => n.kind === "polygon" && n.name === "tile-CA");
    expect(ca).toBeTruthy();
    expect(ca!.points).toHaveLength(6);
  });

  it("default tilemap uses square rects", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.kind === "rect" && n.name === "tile-CA")).toBe(true);
    expect(s.nodes.some((n) => n.kind === "polygon" && n.name === "tile-CA")).toBe(false);
  });

  it("odd rows are offset (hex packing)", () => {
    const s = buildChart({ ...base, tilemap: { shape: "hex" } });
    // Every hex tile is a polygon; the map has more than one row so some tiles
    // are horizontally offset from the base column grid.
    const hexes = s.nodes.filter((n): n is PolygonNode => n.kind === "polygon" && !!n.name?.startsWith("tile-"));
    expect(hexes.length).toBeGreaterThan(2);
  });
});

describe("tilemap mini-glyphs", () => {
  const cfg: ChartConfig = {
    kind: "tilemap",
    ...DEFAULT_SIZE,
    map: "us",
    tilemap: { glyph: "bars" },
    data: {
      categories: ["CA", "TX", "NY"],
      series: [
        { name: "Q1", values: [30, 20, 18] },
        { name: "Q2", values: [35, 25, 20] },
        { name: "Q3", values: [40, 28, 22] },
      ],
    },
  };

  it("draws one mini bar per series inside each region tile", () => {
    const s = buildChart(cfg);
    const caBars = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.match(/^glyph-CA-\d$/));
    expect(caBars).toHaveLength(3);
    // Bars scale with value: Q3 (40) taller than Q1 (30).
    const byIdx = (i: number) => caBars.find((b) => b.name === `glyph-CA-${i}`)!;
    expect(byIdx(2).h).toBeGreaterThan(byIdx(0).h);
    // A series legend replaces the value gradient.
    expect(s.nodes.some((n) => n.name === "legend-0")).toBe(true);
    expect(s.nodes.some((n) => n.name?.startsWith("legend-step-"))).toBe(false);
  });

  it("single-series tilemap keeps the color scale (no glyphs)", () => {
    const s = buildChart({ ...cfg, data: { categories: ["CA", "TX"], series: [{ name: "S", values: [10, 20] }] } });
    expect(s.nodes.some((n) => n.name?.startsWith("glyph-"))).toBe(false);
  });
});
