import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { layoutViolin } from "../src/core/layout/violin";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { PolygonNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Violin plots — density shape, palette, sample memoization. */

describe("violin", () => {
  const rows = Array.from({ length: 10 }, (_, i) => ({
    name: `s${i + 1}`,
    values: [100 + i, 200 + i * 3, 500 + i * 8],
  }));
  const cfg: ChartConfig = {
    kind: "violin",
    ...DEFAULT_SIZE,
    data: { categories: ["Tight", "Mid", "Wide"], series: rows },
    decorations: { categoryAxis: true, valueAxis: true },
  };
  const s = buildChart(cfg);

  it("draws a density polygon and a median tick per category", () => {
    [0, 1, 2].forEach((c) => {
      expect(s.nodes.some((n) => n.kind === "polygon" && n.name === `violin-${c}`)).toBe(true);
      expect(s.nodes.some((n) => n.kind === "line" && n.name === `median-${c}`)).toBe(true);
    });
  });

  it("wider-spread data produces a taller violin", () => {
    const height = (c: number) => {
      const p = s.nodes.find((n): n is PolygonNode => n.name === `violin-${c}`)!;
      const ys = p.points.map((q) => q.y);
      return Math.max(...ys) - Math.min(...ys);
    };
    // "Wide" (500–572) spans more of the shared axis than "Tight" (100–109).
    expect(height(2)).toBeGreaterThan(height(0));
  });
});

describe("violin honours a custom palette length", () => {
  it("colours a category beyond the default 8 by the palette, not modulo 8", () => {
    // A 5-colour palette: category index 9 should map to 9 % 5 = 4, not (9 % 8) % 5 = 1.
    const palette = ["#111111", "#222222", "#333333", "#444444", "#555555"];
    const cats = Array.from({ length: 10 }, (_, i) => `C${i}`);
    const cfg: ChartConfig = {
      kind: "violin",
      width: 900,
      height: 300,
      style: { ...DEFAULT_STYLE, palette },
      data: {
        categories: cats,
        series: Array.from({ length: 6 }, (_, s) => ({
          name: `o${s}`,
          values: cats.map((_, c) => 10 + c + s * 2),
        })),
      },
    };
    const res = layoutViolin(cfg, { ...DEFAULT_STYLE, palette }, DEFAULT_DECOR);
    // The body stroke carries the raw palette color (the fill is lightened).
    const body9 = res.nodes.find((n) => n.name === "violin-9") as { stroke?: string } | undefined;
    expect(body9).toBeDefined();
    // palette[9 % 5] = "#555555"; the old (9 % 8) % 5 = 1 would give "#222222".
    expect(body9!.stroke).toBe("#555555");
  });
});

describe("violin sample memoization", () => {
  const cfg: ChartConfig = {
    kind: "violin",
    width: 640,
    height: 320,
    data: {
      categories: ["A", "B", "C"],
      series: Array.from({ length: 8 }, (_, s) => ({
        name: `o${s}`,
        values: [10 + s, 40 + s * 2, 70 + s],
      })),
    },
  };

  it("renders a violin body per populated category", () => {
    const res = layoutViolin(cfg, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(res.nodes.filter((n) => n.name?.startsWith("violin-"))).toHaveLength(3);
  });

  it("produces identical output on repeated layout calls (cache is per-call)", () => {
    const a = JSON.stringify(layoutViolin(cfg, DEFAULT_STYLE, DEFAULT_DECOR).nodes);
    const b = JSON.stringify(layoutViolin(cfg, DEFAULT_STYLE, DEFAULT_DECOR).nodes);
    expect(a).toBe(b);
  });
});
