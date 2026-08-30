import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { RectNode, SymbolNode, TextNode } from "../src/core/scene";
import { symbolPoints } from "../src/core/geometry";
import type { ChartConfig } from "../src/core/types";

/** Tilemap cartogram — hex tiles and mini-glyphs. */

describe("tilemap hex tiles", () => {
  const base: ChartConfig = {
    kind: "tilemap",
    ...DEFAULT_SIZE,
    map: "us",
    data: { categories: ["CA", "TX", "NY"], series: [{ name: "S", values: [100, 80, 60] }] },
  };

  it("draws hex tiles as FILLED symbols, never as polygons", () => {
    // A polygon is the one thing this tile must not be: Office.js has no
    // freeform fill and degrades a PolygonNode to its outline, which on a
    // choropleth discards the entire encoding. A SymbolNode names a preset the
    // host fills itself.
    const s = buildChart({ ...base, tilemap: { shape: "hex" } });
    const ca = s.nodes.find((n): n is SymbolNode => n.kind === "symbol" && n.name === "tile-CA");
    expect(ca).toBeTruthy();
    expect(ca!.shape).toBe("hexagon");
    expect(ca!.fill).toBeTruthy();
    expect(s.nodes.some((n) => n.kind === "polygon" && !!n.name?.startsWith("tile-"))).toBe(false);
  });

  it("default tilemap uses square rects", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.kind === "rect" && n.name === "tile-CA")).toBe(true);
    expect(s.nodes.some((n) => n.kind === "symbol" && n.name === "tile-CA")).toBe(false);
  });

  it("odd rows are offset half a cell — the thing that makes it a hex grid", () => {
    // This used to COUNT the tiles and call that packing, so it passed with the
    // offset deleted. Measured against the square layout instead, which shares
    // the same column grid and shifts nothing: if odd rows are offset, the hex
    // map resolves to strictly MORE distinct tile centres than there are
    // columns, and the square map is what says how many columns there are.
    const squares = buildChart(base).nodes.filter(
      (n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("tile-"),
    );
    const hexes = buildChart({ ...base, tilemap: { shape: "hex" } }).nodes.filter(
      (n): n is SymbolNode => n.kind === "symbol" && !!n.name?.startsWith("tile-"),
    );
    expect(hexes).toHaveLength(squares.length);
    const cols = new Set(squares.map((t) => Math.round(t.x * 100))).size;
    const centres = new Set(hexes.map((t) => Math.round(t.cx * 100))).size;
    expect(cols).toBeGreaterThan(1);
    expect(centres).toBeGreaterThan(cols);
  });

  it("no two hex tiles overlap, at any size the grid is drawn at", () => {
    /**
     * The size of a hex tile is `HEX_W` (√3/2) of its cell, and this is the
     * whole reason for that number — so it should not be possible to change it
     * and keep the suite green.
     *
     * The tile used to be a POINTY-top hexagon: narrow across, tall down, and
     * the grid shifts odd rows sideways, so its narrow axis faced the shift and
     * a full-cell height was free. The preset that replaced it points left and
     * right, which swaps the narrow axis onto the one that matters. At a full
     * cell its flat top and bottom reach the offset row below.
     *
     * Separating-axis over the real outlines rather than bounding boxes: the
     * boxes of two hexes an odd row apart DO overlap, at the corners where
     * neither shape has any ink. A box test would fail a correct grid, which is
     * how you end up loosening a test until it stops meaning anything.
     */
    const overlaps = (a: { x: number; y: number }[], b: { x: number; y: number }[]) => {
      for (const poly of [a, b]) {
        for (let i = 0; i < poly.length; i++) {
          const p = poly[i];
          const q = poly[(i + 1) % poly.length];
          const nx = -(q.y - p.y);
          const ny = q.x - p.x;
          const proj = (pts: { x: number; y: number }[]) => pts.map((t) => t.x * nx + t.y * ny);
          const pa = proj(a);
          const pb = proj(b);
          // A shared edge is not an overlap: tiles are allowed to touch.
          if (Math.min(...pa) >= Math.max(...pb) - 1e-9 || Math.min(...pb) >= Math.max(...pa) - 1e-9) return false;
        }
      }
      return true;
    };

    for (const [w, h] of [
      [480, 300],
      [960, 540],
      [320, 240],
      [240, 160],
    ]) {
      const tiles = buildChart({ ...base, width: w, height: h, tilemap: { shape: "hex" } }).nodes.filter(
        (n): n is SymbolNode => n.kind === "symbol" && !!n.name?.startsWith("tile-"),
      );
      expect(tiles.length).toBeGreaterThan(20);
      const outlines = tiles.map((t) => ({ name: t.name, pts: symbolPoints(t.shape, t.cx, t.cy, t.size) }));
      for (let i = 0; i < outlines.length; i++) {
        for (let j = i + 1; j < outlines.length; j++) {
          expect(
            overlaps(outlines[i].pts, outlines[j].pts),
            `${w}x${h}: ${outlines[i].name} overlaps ${outlines[j].name}`,
          ).toBe(false);
        }
      }
    }
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

describe("the gradient legend's two ends on a small frame", () => {
  /**
   * `legend-min` and `legend-max` each own HALF the colour bar and are anchored
   * to its outer edges, so their ink meets in the middle as soon as a number is
   * wider than `lw / 2` — and the bar is `min(gridW * 0.5, fs * 12)`, a few
   * points wide on a thumbnail.
   *
   * Pinned here rather than in `frame-fit`'s overlap sweep because that gate
   * covers 160x120 upward and this shows at 120x90. Nothing overflows, so no
   * frame gate could ever have seen it: the two numbers are simply drawn on top
   * of each other, inside the chart.
   */
  const ink = (t: TextNode) => {
    // The label is anchored to an edge of its half, so its ink is what collides,
    // not its box. `w` is the half-bar; the text is as wide as the text.
    const est = t.text.length * t.fontSize * 0.55;
    const x = t.align === "right" ? t.x + t.w - est : t.x;
    return { x0: x, x1: x + est };
  };

  it("never draws the two scale ends on top of each other", () => {
    for (const [w, h] of [
      [120, 90],
      [160, 120],
      [200, 150],
    ] as [number, number][]) {
      const nodes = buildChart({ ...sampleConfig("tilemap"), width: w, height: h } as ChartConfig).nodes;
      const lo = nodes.find((n): n is TextNode => n.kind === "text" && n.name === "legend-min");
      const hi = nodes.find((n): n is TextNode => n.kind === "text" && n.name === "legend-max");
      // Dropping the pair is a legitimate outcome — a scale whose ends cannot be
      // read is not a scale — but they go together or not at all.
      expect(!!lo, `only one end drawn at ${w}x${h}`).toBe(!!hi);
      if (!lo || !hi) continue;
      expect(ink(lo).x1, `the scale ends overlap at ${w}x${h}`).toBeLessThanOrEqual(ink(hi).x0 + 0.5);
    }
  });
});
