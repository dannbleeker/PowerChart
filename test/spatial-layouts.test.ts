import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { detectLayout } from "../src/core/layout/tilemap-layouts";
import type { ChartConfig } from "../src/core/types";
import type { PolygonNode, RectNode, TextNode } from "../src/core/scene";

/**
 * Bug-hunt guards for the spatial layouts (tilemap, calendar heatmap, waffle,
 * mekko): geometry and encoding faults an all-positive, long-range,
 * denominator-consistent sample never reaches.
 */

const iso = (n: number) => new Date(Date.UTC(2026, 0, 5) + n * 86400000).toISOString().slice(0, 10);

describe("hex tilemap survives the Office.js polygon degradation", () => {
  const cfg: ChartConfig = {
    kind: "tilemap",
    ...DEFAULT_SIZE,
    width: 480,
    height: 300,
    map: "us",
    tilemap: { shape: "hex" },
    data: { categories: ["CA", "TX", "NY", "FL"], series: [{ name: "Sales", values: [10, 40, 25, 5] }] },
  };

  it("outlines each hex in its own value color, not the background", () => {
    // Office.js has no freeform fill and draws a polygon as `stroke ?? fill`
    // (scene.ts's parity contract), so a background-colored stroke rendered the
    // whole cartogram white-on-white in the add-in.
    const tiles = buildChart(cfg).nodes.filter(
      (n): n is PolygonNode => n.kind === "polygon" && !!n.name?.startsWith("tile-"),
    );
    expect(tiles.length).toBeGreaterThan(2);
    for (const t of tiles) expect(t.stroke).toBe(t.fill);
    // …and the tiles still differ by value, so the outline carries the encoding.
    const ca = tiles.find((t) => t.name === "tile-CA")!;
    const tx = tiles.find((t) => t.name === "tile-TX")!;
    expect(ca.stroke).not.toBe(tx.stroke);
  });
});

describe("tilemap mini-glyph bars encode negative values", () => {
  const cfg: ChartConfig = {
    kind: "tilemap",
    ...DEFAULT_SIZE,
    map: "us",
    tilemap: { glyph: "bars" },
    data: {
      categories: ["CA", "TX", "NY"],
      series: [
        { name: "2024", values: [10, -20, 30] },
        { name: "2025", values: [12, 25, -5] },
      ],
    },
  };

  it("a decline draws below the zero line instead of collapsing to zero height", () => {
    const bars = buildChart(cfg).nodes.filter(
      (n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("glyph-"),
    );
    const at = (name: string) => bars.find((b) => b.name === name)!;
    // -20 and -5 used to be h=0, pixel-identical to a value of exactly 0.
    expect(at("glyph-TX-0").h).toBeGreaterThan(0);
    expect(at("glyph-NY-1").h).toBeGreaterThan(0);
    // Within a tile the positive bar's floor IS the negative bar's ceiling:
    // both hang off one shared zero line.
    expect(at("glyph-NY-0").y + at("glyph-NY-0").h).toBeCloseTo(at("glyph-NY-1").y, 6);
    expect(at("glyph-TX-1").y + at("glyph-TX-1").h).toBeCloseTo(at("glyph-TX-0").y, 6);
    // Magnitude still reads: -20 is a taller bar than -5.
    expect(at("glyph-TX-0").h).toBeGreaterThan(at("glyph-NY-1").h);
  });

  it("all-positive data keeps the original baseline scale", () => {
    const s = buildChart({
      ...cfg,
      data: {
        categories: ["CA", "TX"],
        series: [
          { name: "Q1", values: [30, 20] },
          { name: "Q2", values: [40, 28] },
        ],
      },
    });
    const bars = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("glyph-CA-"));
    const bottoms = bars.map((b) => b.y + b.h);
    // Every bar still sits on one shared floor and grows with its value.
    expect(Math.max(...bottoms) - Math.min(...bottoms)).toBeLessThan(0.001);
    expect(bars[1].h).toBeGreaterThan(bars[0].h);
  });
});

describe("calendar heatmap legend stays inside the frame", () => {
  const calendar = (days: number, width = 480, height = 300): ChartConfig => {
    const categories = Array.from({ length: days }, (_, i) => iso(i));
    return {
      kind: "heatmap",
      ...DEFAULT_SIZE,
      width,
      height,
      heatmap: { calendar: true },
      data: { categories, series: [{ name: "commits", values: categories.map((_, i) => (i * 7) % 13) }] },
    };
  };

  // `cell` grows as the range shortens while the legend band is a constant
  // reserve, so short calendars spilled the swatches off the bottom and right.
  for (const [days, w, h] of [
    [7, 480, 300],
    [30, 480, 300],
    [30, 480, 400],
    [30, 320, 300],
    [365, 480, 300],
  ] as const) {
    it(`${days} days at ${w}x${h} keeps every legend rect on canvas`, () => {
      const s = buildChart(calendar(days, w, h));
      const swatches = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("legend-step-"));
      expect(swatches.length).toBeGreaterThan(0);
      for (const r of swatches) {
        expect(r.y + r.h, `${r.name} bottom`).toBeLessThanOrEqual(h + 0.001);
        expect(r.x + r.w, `${r.name} right`).toBeLessThanOrEqual(w + 0.001);
      }
      const more = s.nodes.find((n): n is TextNode => n.kind === "text" && n.name === "legend-more")!;
      expect(more.x + more.w).toBeLessThanOrEqual(w + 0.001);
    });
  }
});

describe("waffle honors an inconsistent 100%= denominator", () => {
  it("equal values get equal areas that match the printed shares", () => {
    // hundredPercent 80 is smaller than the parts (50 + 50): the quotas summed
    // to 124 cells and the 100-cell cap truncated whichever category came last.
    const s = buildChart({
      kind: "waffle",
      ...DEFAULT_SIZE,
      width: 480,
      height: 300,
      data: {
        categories: ["Region A", "Region B"],
        series: [{ name: "V", values: [50, 50] }],
        hundredPercent: [80],
      },
    });
    const byFill = new Map<string, number>();
    for (const n of s.nodes) {
      if (n.kind === "rect" && n.name?.startsWith("waffle-cell-")) byFill.set(n.fill, (byFill.get(n.fill) ?? 0) + 1);
    }
    const counts = [...byFill.values()].sort((a, b) => b - a);
    expect(counts).toEqual([50, 50]);
    // The legend agrees with the picture.
    const labels = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("legend-label-"));
    expect(labels.map((l) => l.text)).toEqual(["Region A  50%", "Region B  50%"]);
  });

  it("a denominator larger than the parts still leaves no-data cells", () => {
    const s = buildChart({
      kind: "waffle",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Done", "Open"],
        series: [{ name: "V", values: [20, 30] }],
        hundredPercent: [200],
      },
    });
    const cells = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("waffle-cell-"));
    const colored = new Set(cells.map((c) => c.fill));
    expect(colored.size).toBe(3); // two categories + gray remainder
  });
});

describe("tilemap auto-detection refuses an ambiguous code set", () => {
  it("codes valid in two grids fall through to the 'set map' hint", () => {
    // DE/MD/ME/MT/AL are ISO-2 (Germany, Moldova, Montenegro, Malta, Albania)
    // AND US postal codes — first-match-wins painted Europe onto the US map.
    expect(detectLayout(["DE", "MD", "ME", "MT", "AL"])).toBeNull();
    const s = buildChart({
      kind: "tilemap",
      ...DEFAULT_SIZE,
      data: { categories: ["DE", "MD", "ME", "MT", "AL"], series: [{ name: "Sales", values: [1, 2, 3, 4, 5] }] },
    });
    expect(s.nodes.some((n) => n.name === "tilemap-error")).toBe(true);
    expect(s.nodes.some((n) => n.name?.startsWith("tile-"))).toBe(false);
  });

  it("unambiguous sets still resolve, EU before the wider Europe grid", () => {
    expect(detectLayout(["CA", "TX", "NY", "FL"])).toBe("us");
    expect(detectLayout(["DK", "SE", "DE", "FR"])).toBe("eu");
    expect(detectLayout(["DK", "SE", "NO", "CH"])).toBe("europe");
  });
});

describe("mekko honors the transparent floating-segment idiom", () => {
  it('color "transparent" occupies the column but draws no rect', () => {
    const s = buildChart({
      kind: "mekko",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Base", values: [10, 12, 14], color: "transparent" },
          { name: "Range", values: [20, 18, 22] },
        ],
      },
    });
    // A literal "transparent" fill reaches both PowerPoint sinks as a paint they
    // cannot express (mid grey in pptx, rejected by Office.js setSolidColor).
    expect(s.nodes.some((n) => n.kind === "rect" && n.fill === "transparent")).toBe(false);
    expect(s.nodes.some((n) => n.name === "seg-0-0")).toBe(false);
    // The visible series above it is unmoved: it still starts where the
    // invisible base ended (two thirds up a 10/20 column).
    const seg = s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "seg-1-0")!;
    expect(seg).toBeTruthy();
    const solid = buildChart({
      kind: "mekko",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Base", values: [10, 12, 14] },
          { name: "Range", values: [20, 18, 22] },
        ],
      },
    });
    const ref = solid.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "seg-1-0")!;
    expect(seg.y).toBeCloseTo(ref.y, 6);
    expect(seg.h).toBeCloseTo(ref.h, 6);
  });
});
