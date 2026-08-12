import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { PolygonNode, RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Tilemap cartogram — hex tiles and mini-glyphs. */

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

describe("the legend's numbers are not cut off by the frame", () => {
  /**
   * The min and max labels under the gradient strip had their descenders cut by
   * the bottom of the frame on EVERY tilemap at EVERY size — 1.3pt at a 10pt
   * font, 1.9pt at 18 — because two reservations were each a little short.
   *
   * `legendH` reserved `fs * 2.4` for a legend whose last ink is `fs * 2.65`
   * below the grid (`fs * 0.5` of clearance, then the label box `fs * 0.95`
   * down and `fs * 1.2` tall). And `rowsBottom` adds a gutter PER ROW where the
   * height budget pays for the gaps BETWEEN rows — one gutter more than it is
   * given. Together they predict `fs * 0.0785 + 0.5` of overflow: 1.29 and 1.91
   * at those two fonts, which is the measurement to the tenth.
   */
  const overflow = (width: number, height: number, fontSize: number) => {
    const scene = buildChart({ ...sampleConfig("tilemap"), width, height, style: { fontSize } } as ChartConfig);
    const labels = scene.nodes.filter(
      (n): n is TextNode => n.kind === "text" && (n.name === "legend-min" || n.name === "legend-max"),
    );
    expect(labels.length, `${width}x${height} fs=${fontSize} drew no legend numbers`).toBe(2);
    // Where the glyphs land, not where the box is: the box is what was reserved
    // for, and the ink is what the frame cuts.
    return Math.max(
      ...labels.map((t) => {
        const base =
          t.valign === "top"
            ? t.y + t.fontSize
            : t.valign === "bottom"
              ? t.y + t.h - t.fontSize * 0.25
              : t.y + t.h / 2 + t.fontSize * 0.36;
        return base + t.fontSize * 0.21 - scene.height;
      }),
    );
  };

  it("keeps them inside the frame at every ordinary size", () => {
    for (const [w, h] of [
      [480, 300],
      [960, 540],
      [640, 400],
      [240, 160],
    ]) {
      for (const fs of [8, 10, 14, 18]) {
        expect(overflow(w, h, fs), `${w}x${h} fs=${fs} cut the legend numbers`).toBeLessThanOrEqual(0);
      }
    }
  });

  it("still draws the map and the legend it made room for", () => {
    // The reservation must not be paid by dropping what it reserves for, nor by
    // shrinking the map to nothing.
    const scene = buildChart({ ...sampleConfig("tilemap"), ...DEFAULT_SIZE } as ChartConfig);
    expect(scene.nodes.filter((n) => n.name?.startsWith("legend-step-")).length).toBeGreaterThan(10);
    const tiles = scene.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("tile-"));
    expect(tiles.length).toBeGreaterThan(20);
    expect(Math.min(...tiles.map((t) => t.h))).toBeGreaterThan(6);
  });
});
