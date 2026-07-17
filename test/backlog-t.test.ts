import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, RectNode, TextNode } from "../src/core/scene";

/**
 * Batch T — heatmap cell-size encoding, heatmap row clustering (dendrogram),
 * and the stacked-area combo base.
 */

describe("heatmap cell-size encoding", () => {
  const cfg: ChartConfig = {
    kind: "heatmap",
    ...DEFAULT_SIZE,
    data: {
      categories: ["X"],
      series: [
        { name: "big", values: [1] },
        { name: "small", values: [0.25] },
      ],
    },
    heatmap: { sizeEncode: true },
  };
  const s = buildChart(cfg);
  const cell = (ri: number) => s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === `cell-${ri}-0`)!;

  it("sizes each cell by its magnitude (area ∝ |value|)", () => {
    const big = cell(0);
    const small = cell(1);
    // √(0.25) = 0.5 → the small cell's side is about half the big one's.
    expect(small.w).toBeLessThan(big.w);
    expect(small.w).toBeGreaterThan(0);
    expect(Math.abs(small.w / big.w - 0.5)).toBeLessThan(0.05);
    // Cells stay square and centred.
    expect(big.w).toBeCloseTo(big.h, 5);
  });

  it("plain heatmaps still fill the whole cell", () => {
    const plain = buildChart({ ...cfg, heatmap: {} });
    const full = plain.nodes.find((n): n is RectNode => n.name === "cell-0-0")!;
    const sized = cell(0);
    expect(full.w).toBeGreaterThan(sized.w);
  });
});

describe("heatmap row clustering", () => {
  const cfg: ChartConfig = {
    kind: "heatmap",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "Hi1", values: [80, 82, 85] },
        { name: "Lo1", values: [20, 22, 19] },
        { name: "Hi2", values: [78, 80, 83] },
        { name: "Lo2", values: [22, 18, 24] },
      ],
    },
    heatmap: { cluster: true },
  };
  const s = buildChart(cfg);
  const rowName = (ri: number) => (s.nodes.find((n): n is TextNode => n.kind === "text" && n.name === `row-${ri}`)!).text;

  it("reorders rows so similar rows are adjacent and draws a dendrogram", () => {
    const order = [0, 1, 2, 3].map(rowName);
    const hiPositions = [order.indexOf("Hi1"), order.indexOf("Hi2")].sort((a, b) => a - b);
    const loPositions = [order.indexOf("Lo1"), order.indexOf("Lo2")].sort((a, b) => a - b);
    // Each similar pair ends up adjacent.
    expect(hiPositions[1] - hiPositions[0]).toBe(1);
    expect(loPositions[1] - loPositions[0]).toBe(1);
    // The dendrogram tree is drawn.
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && n.name === "dendro-v")).toBe(true);
    expect(s.nodes.some((n) => n.name === "dendro-h")).toBe(true);
  });

  it("leaves rows in sheet order without the flag", () => {
    const plain = buildChart({ ...cfg, heatmap: {} });
    const name0 = (plain.nodes.find((n): n is TextNode => n.name === "row-0")!).text;
    expect(name0).toBe("Hi1");
    expect(plain.nodes.some((n) => n.name === "dendro-v")).toBe(false);
  });
});

describe("combo stacked-area base", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Jan", "Feb", "Mar"],
      series: [
        { name: "Cloud", values: [20, 24, 28] },
        { name: "Licenses", values: [15, 14, 16] },
        { name: "Margin %", type: "line", values: [22, 24, 26] },
      ],
    },
    combo: { columns: "area" },
    secondaryAxis: true,
  };
  const s = buildChart(cfg);

  it("draws a stacked area base with the line overlaid", () => {
    // The area base emits filled area slabs...
    expect(s.nodes.some((n) => n.name?.startsWith("area-"))).toBe(true);
    // ...and a secondary axis for the line series.
    expect(s.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
    // ...and line segments over the top.
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("combo-line-"))).toBe(true);
  });
});

describe("heatmap sign marks", () => {
  const hm = (heatmap: ChartConfig["heatmap"], vals: number[][] = [[8, -7], [-6, 9]], extra: Partial<ChartConfig> = {}): ChartConfig => ({
    kind: "heatmap",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2"],
      series: vals.map((v, i) => ({ name: `R${i}`, values: v })),
    },
    heatmap,
    ...extra,
  });
  const signs = (c: ChartConfig) => buildChart(c).nodes.filter((n) => n.name?.startsWith("cell-sign-"));
  const signAt = (c: ChartConfig, ri: number, col: number) =>
    buildChart(c).nodes.find((n) => n.name === `cell-sign-${ri}-${col}`);

  it("is off by default — the heatmap it has always drawn", () => {
    expect(signs(hm({ sizeEncode: true }))).toHaveLength(0);
    expect(buildChart(hm({ sizeEncode: true })).nodes).toEqual(buildChart(hm({ sizeEncode: true, symbols: undefined })).nodes);
  });

  it("marks + on positives and − on negatives", () => {
    const c = hm({ sizeEncode: true, symbols: "sign" });
    expect(signs(c)).toHaveLength(4);
    // The plus is preset geometry; the minus IS that plus's horizontal arm, so
    // the pair matches by construction — a rect, not a second preset.
    expect(signAt(c, 0, 0)!.kind).toBe("symbol"); // +8
    expect((signAt(c, 0, 0) as { shape: string }).shape).toBe("plus");
    expect(signAt(c, 0, 1)!.kind).toBe("rect"); // -7
    expect(signAt(c, 1, 0)!.kind).toBe("rect"); // -6
    expect(signAt(c, 1, 1)!.kind).toBe("symbol"); // +9
  });

  it("leaves a zero cell unmarked — the scale paints it neutral, and zero has no sign", () => {
    // The colour says "neither" (divergingScale's exact-zero is white). A "+"
    // over that white would have the two channels contradict each other on one
    // cell — and to the greyscale reader, the only one the mark exists for, the
    // glyph is the whole story. Reachable without any opt-out: labels are only
    // drawn if they FIT, so a dense diverging matrix hits this by default.
    const c = hm({ symbols: "sign" }, [[0, -7], [-6, 9]], { decorations: { segmentLabels: false } });
    const nodes = buildChart(c).nodes;
    expect((nodes.find((n) => n.name === "cell-0-0") as RectNode).fill).toBe("#ffffff");
    expect(nodes.find((n) => n.name === "cell-sign-0-0")).toBeUndefined();
    // ...while its signed neighbours are still marked.
    expect(nodes.find((n) => n.name === "cell-sign-0-1")!.kind).toBe("rect"); // -7
    expect(nodes.find((n) => n.name === "cell-sign-1-1")!.kind).toBe("symbol"); // +9
  });

  it("marks a calendar heatmap, which never draws labels at all", () => {
    // The calendar layout returns before the grid's cell loop, so it silently
    // ignored the option — on the one layout where colour is the ONLY sign
    // carrier, i.e. exactly where the marks matter most.
    const cal = (symbols?: "sign"): ChartConfig => ({
      kind: "heatmap",
      ...DEFAULT_SIZE,
      data: {
        categories: Array.from({ length: 40 }, (_, i) => `2026-03-${String((i % 28) + 1).padStart(2, "0")}`),
        series: [{ name: "Net flow", values: Array.from({ length: 40 }, (_, i) => (i % 2 ? 1 : -1) * (i + 1)) }],
      },
      heatmap: { calendar: true, symbols },
    });
    const signs = (c: ChartConfig) => buildChart(c).nodes.filter((n) => n.name?.startsWith("cell-sign-"));
    expect(signs(cal("sign")).length).toBeGreaterThan(10);
    // Still off by default.
    expect(signs(cal(undefined))).toHaveLength(0);
  });

  it("stays inert on one-signed data, where every mark would be identical", () => {
    expect(signs(hm({ sizeEncode: true, symbols: "sign" }, [[8, 7], [6, 9]]))).toHaveLength(0);
    expect(signs(hm({ sizeEncode: true, symbols: "sign" }, [[-8, -7], [-6, -9]]))).toHaveLength(0);
  });

  it("yields to the value label, which already prints the sign", () => {
    // Labels on and fitting => no marks: two signs in one cell is noise.
    expect(signs(hm({ symbols: "sign" }, [[8, -7], [-6, 9]], { decorations: { segmentLabels: true } }))).toHaveLength(0);
    // Labels off => the marks are the only sign carrier, so they appear.
    expect(signs(hm({ symbols: "sign" }, [[8, -7], [-6, 9]], { decorations: { segmentLabels: false } }))).toHaveLength(4);
  });

  it("inks each mark against its own cell, like the labels do", () => {
    for (const n of signs(hm({ sizeEncode: true, symbols: "sign" }))) {
      expect((n as { fill: string }).fill).toMatch(/^#(0b0b0b|ffffff)$/);
    }
  });

  it("keeps every mark inside the cell as DRAWN, however small sizeEncode makes it", () => {
    // A near-zero value shrinks its cell to almost nothing; a mark sized off the
    // slot rather than the drawn square would spill across its neighbours.
    const c = hm({ sizeEncode: true, symbols: "sign" }, [[100, -0.4], [-0.2, 100]]);
    const nodes = buildChart(c).nodes;
    for (const ri of [0, 1]) {
      for (const col of [0, 1]) {
        const cell = nodes.find((n) => n.name === `cell-${ri}-${col}`) as RectNode;
        const s = nodes.find((n) => n.name === `cell-sign-${ri}-${col}`);
        if (!s) continue;
        const b =
          s.kind === "symbol"
            ? { x: s.cx - s.size, y: s.cy - s.size, w: s.size * 2, h: s.size * 2 }
            : { x: (s as RectNode).x, y: (s as RectNode).y, w: (s as RectNode).w, h: (s as RectNode).h };
        expect(b.x).toBeGreaterThanOrEqual(cell.x - 1e-9);
        expect(b.y).toBeGreaterThanOrEqual(cell.y - 1e-9);
        expect(b.x + b.w).toBeLessThanOrEqual(cell.x + cell.w + 1e-9);
        expect(b.y + b.h).toBeLessThanOrEqual(cell.y + cell.h + 1e-9);
      }
    }
  });

  it("skips a cell too small to carry a legible mark, and cells with no value", () => {
    const c = hm({ sizeEncode: true, symbols: "sign" }, [[100, 0], [-0.001, 100]]);
    // A zero-magnitude cell has no square to draw in.
    expect(buildChart(c).nodes.find((n) => n.name === "cell-sign-0-1")).toBeUndefined();
    const nulls: ChartConfig = {
      kind: "heatmap",
      ...DEFAULT_SIZE,
      data: { categories: ["Q1", "Q2"], series: [{ name: "R0", values: [8, null] }, { name: "R1", values: [-6, 9] }] },
      heatmap: { sizeEncode: true, symbols: "sign" },
    };
    expect(buildChart(nulls).nodes.find((n) => n.name === "cell-sign-0-1")).toBeUndefined();
  });
});
