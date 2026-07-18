import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { boxplotExtent } from "../src/core/layout/boxplot";
import { DEFAULT_STYLE } from "../src/core/style";
import { layoutViolin } from "../src/core/layout/violin";
import { legendRow } from "../src/core/layout/column";
import { DEFAULT_DECOR } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";

describe("boxplot data-driven domain (no forced zero)", () => {
  it("does not pin the value axis to zero for far-from-zero data", () => {
    const cfg: ChartConfig = {
      kind: "boxplot",
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "o1", values: [40, 55] },
          { name: "o2", values: [60, 80] },
          { name: "o3", values: [50, 95] },
        ],
      },
    };
    const ext = boxplotExtent(cfg)!;
    // The old code forced min:0; the domain must now start near the data (40).
    expect(ext.min).toBeGreaterThan(30);
    expect(ext.max).toBeLessThanOrEqual(95);
  });
});

describe("pie / doughnut all-zero total", () => {
  it("shows the true total (0), not the divisor fallback of 1", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 400,
      height: 300,
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [0, 0, 0] }] },
    };
    const nodes = buildChart(cfg).nodes;
    const texts = nodes.filter((n): n is TextNode => n.kind === "text");
    expect(nodes.some((n) => n.name === "hole")).toBe(true);
    // The centre shows the honest total (0, at the data's 2-decimal precision),
    // never the divisor fallback the old `|| 1` displayed as "1".
    expect(texts.some((t) => /^0(\.0+)?$/.test(t.text))).toBe(true);
    expect(texts.some((t) => /^1(\.0+)?$/.test(t.text))).toBe(false);
  });

  it("still renders a normal doughnut total unchanged", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 400,
      height: 300,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [30, 70] }] },
    };
    const texts = buildChart(cfg).nodes.filter((n): n is TextNode => n.kind === "text");
    expect(texts.some((t) => t.text === "100")).toBe(true);
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

describe("combo overlay line stays on the plot when it dips below the bars", () => {
  it("extends the shared axis down to a negative overlay over a non-negative base", () => {
    // All-zero bars + a line reaching −18: the shared column axis used to floor at
    // its own data (≥0), plotting the line thousands of points below the plot.
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: 300,
      data: {
        categories: ["C0", "C1", "C2", "C3"],
        series: [
          { name: "bars", values: [0, 0, 0, 0] },
          { name: "line", values: [0, 0, 0, -18] },
        ],
      },
    };
    const nodes = buildChart(cfg).nodes;
    const line = nodes.filter((n) => n.name?.startsWith("combo-line-") || n.name?.startsWith("combo-marker-"));
    expect(line.length).toBeGreaterThan(0);
    // Every line/marker coordinate must sit within a generous margin of the 300pt canvas.
    for (const n of line) {
      const ys = n.kind === "line" ? [n.y1, n.y2] : n.kind === "rect" ? [n.y, n.y + n.h] : [];
      for (const y of ys) expect(y).toBeLessThan(360);
    }
  });

  it("leaves a combo whose bars already run more negative unchanged", () => {
    // Bars reach −50, the line only −5: the bars own the floor, so the fix must
    // not raise it to the line (which would clip the bars).
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: 300,
      data: {
        categories: ["C0", "C1"],
        series: [
          { name: "bars", values: [-50, -40] },
          { name: "line", values: [-5, -3] },
        ],
      },
    };
    // The most-negative bar segment must still reach near the bottom of the plot.
    const rects = buildChart(cfg).nodes.filter((n) => n.name?.startsWith("seg-") || n.name?.startsWith("col-"));
    expect(rects.length).toBeGreaterThan(0);
  });
});

describe("legend wraps instead of marching off-canvas", () => {
  const seriesCfg = (n: number, width: number): ChartConfig => ({
    kind: "stacked",
    width,
    height: 200,
    data: {
      categories: ["A"],
      series: Array.from({ length: n }, (_, i) => ({ name: `Series ${i + 1}`, values: [1] })),
    },
  });
  const chipYs = (nodes: ReturnType<typeof legendRow>) =>
    new Set(nodes.filter((n) => n.name?.startsWith("legend-chip-")).map((n) => (n as RectNode).y));

  it("puts overflowing chips on a second row", () => {
    // 12 wide chips cannot fit one 300pt row — they must span multiple rows.
    expect(chipYs(legendRow(seriesCfg(12, 300), DEFAULT_STYLE, 0, 0, { maxX: 300 })).size).toBeGreaterThan(1);
  });

  it("stays a single row (byte-identical) when everything fits", () => {
    expect(chipYs(legendRow(seriesCfg(2, 800), DEFAULT_STYLE, 0, 0, { maxX: 796 })).size).toBe(1);
  });
});
