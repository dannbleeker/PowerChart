import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode } from "../src/core/scene";

/** §3 treemap: squarified rect packing (area ∝ value), with 2-level grouping. */

describe("treemap (flat)", () => {
  const cfg: ChartConfig = {
    kind: "treemap",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "V", values: [100, 50, 30, 20] }] },
    decorations: { segmentLabels: true },
  };
  const s = buildChart(cfg);
  const tile = (i: number) => s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === `tile-${i}`)!;

  it("draws one tile per category with area proportional to value", () => {
    [0, 1, 2, 3].forEach((i) => expect(tile(i)).toBeTruthy());
    const area = (i: number) => tile(i).w * tile(i).h;
    // A (100) roughly double B (50), and > C (30) > D (20).
    expect(area(0)).toBeGreaterThan(area(1));
    expect(area(1)).toBeGreaterThan(area(2));
    expect(area(2)).toBeGreaterThan(area(3));
    expect(area(0) / area(1)).toBeCloseTo(2, 0);
  });

  it("tiles pack without large gaps (cover most of the plot)", () => {
    const totalArea = [0, 1, 2, 3].reduce((a, i) => a + tile(i).w * tile(i).h, 0);
    const plot = DEFAULT_SIZE.width * DEFAULT_SIZE.height;
    expect(totalArea).toBeGreaterThan(plot * 0.6); // fills the bulk of the canvas
  });

  it("aspect ratios stay reasonable (squarified, not slivers)", () => {
    [0, 1, 2, 3].forEach((i) => {
      const t = tile(i);
      const ar = Math.max(t.w / t.h, t.h / t.w);
      expect(ar).toBeLessThan(6);
    });
  });
});

describe("treemap (two-level)", () => {
  const cfg: ChartConfig = {
    kind: "treemap",
    ...DEFAULT_SIZE,
    data: {
      categories: ["G1 | A", "G1 | B", "G2 | C", "G2 | D"],
      series: [{ name: "V", values: [40, 20, 30, 10] }],
    },
    decorations: { segmentLabels: true },
  };
  const s = buildChart(cfg);

  it("nests items inside group cells", () => {
    // Group boxes + headers exist.
    expect(s.nodes.some((n) => n.name === "group-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "group-label-0")).toBe(true);
    // Each item is a tile inside its group.
    [0, 1, 2, 3].forEach((i) => expect(s.nodes.some((n) => n.name === `tile-${i}`)).toBe(true));
    // G1 (total 60) cell is larger than G2 (total 40).
    const g0 = s.nodes.find((n): n is RectNode => n.name === "group-0")!;
    const g1 = s.nodes.find((n): n is RectNode => n.name === "group-1")!;
    expect(g0.w * g0.h).toBeGreaterThan(g1.w * g1.h);
    // Items sit within their group's bounds.
    const a = s.nodes.find((n): n is RectNode => n.name === "tile-0")!;
    expect(a.x).toBeGreaterThanOrEqual(g0.x - 1);
    expect(a.x + a.w).toBeLessThanOrEqual(g0.x + g0.w + 1);
  });
});

describe("treemap (two-level edge cases)", () => {
  // Ungrouped labels mixed in with grouped ones fall into an unnamed "" group,
  // and a zero-value member squarifies to nothing — the grouped path's guards
  // (groupOf/labelOf's no-"|" branch, the skipped-tile branch) that a clean
  // all-grouped fixture never reaches.
  const cfg: ChartConfig = {
    kind: "treemap",
    ...DEFAULT_SIZE,
    data: {
      categories: ["North | Apples", "North | Pears", "loose", "South | Figs", "South | Zero"],
      series: [{ name: "V", values: [80, 40, 25, 30, 0] }],
    },
    decorations: { segmentLabels: true },
  };
  const s = buildChart(cfg);

  it("renders a group for the loose (ungrouped) label alongside the named groups", () => {
    // Three groups: "North", "" (loose), "South" — each gets a box.
    expect(s.nodes.some((n) => n.name === "group-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "group-1")).toBe(true);
    expect(s.nodes.some((n) => n.name === "group-2")).toBe(true);
  });

  it("draws a tile for every non-zero member and skips the zero-value one", () => {
    // The four non-zero members each get a tile; the zero member (index 4) has
    // no area, so squarify hands it no rect and no tile is drawn.
    [0, 1, 2, 3].forEach((i) => expect(s.nodes.some((n) => n.name === `tile-${i}`)).toBe(true));
    expect(s.nodes.some((n) => n.name === "tile-4")).toBe(false);
  });

  it("never throws and produces only finite geometry", () => {
    for (const n of s.nodes) {
      if (n.kind === "rect") {
        expect(Number.isFinite(n.w)).toBe(true);
        expect(Number.isFinite(n.h)).toBe(true);
        expect(n.w).toBeGreaterThanOrEqual(0);
        expect(n.h).toBeGreaterThanOrEqual(0);
      }
    }
  });
});
