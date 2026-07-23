import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { LineNode, RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Heatmap — cell-size encoding, clustering, calendar, marginal totals, sign marks. */

describe("heatmap marginal totals", () => {
  const heat: ChartConfig = {
    kind: "heatmap",
    ...DEFAULT_SIZE,
    heatmap: { totals: "both" },
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [10, 40] },
        { name: "South", values: [20, 30] },
      ],
    },
  };

  it("adds row and column sum strips outside the color scale", () => {
    const s = buildChart(heat);
    expect((s.nodes.find((n) => n.name === "row-total-0") as TextNode).text).toBe("50");
    expect((s.nodes.find((n) => n.name === "col-total-1") as TextNode).text).toBe("70");
    // Totals sit outside the matrix: right of the last cell / below the last row.
    const cell = s.nodes.find((n) => n.name === "cell-0-1") as RectNode;
    const rowTotal = s.nodes.find((n) => n.name === "row-total-bg-0") as RectNode;
    expect(rowTotal.x).toBeGreaterThanOrEqual(cell.x + cell.w);
    // Neutral fill, not the value color scale.
    expect(rowTotal.fill).toBe("#f0efec");
  });

  it("row-only mode omits the column strip; default has neither", () => {
    const rowOnly = buildChart({ ...heat, heatmap: { totals: "row" } });
    expect(rowOnly.nodes.some((n) => n.name === "row-total-0")).toBe(true);
    expect(rowOnly.nodes.some((n) => n.name === "col-total-0")).toBe(false);
    const plain = buildChart({ ...heat, heatmap: {} });
    expect(plain.nodes.some((n) => n.name?.includes("total"))).toBe(false);
  });
});

describe("heatmap calendar layout", () => {
  const days = Array.from({ length: 21 }, (_, i) => {
    const d = new Date(Date.UTC(2025, 0, 6 + i));
    return d.toISOString().slice(0, 10);
  });
  const base: ChartConfig = {
    kind: "heatmap",
    ...DEFAULT_SIZE,
    heatmap: { calendar: true },
    data: { categories: days, series: [{ name: "Commits", values: days.map((_, i) => (i % 7) + 1) }] },
  };

  it("lays days out on a weekday × week grid", () => {
    const s = buildChart(base);
    expect(s.nodes.filter((n) => n.name?.match(/^cell-\d+$/))).toHaveLength(21);
    expect(s.nodes.some((n) => n.name?.startsWith("weekday-"))).toBe(true);
    expect(s.nodes.some((n) => n.name?.startsWith("month-"))).toBe(true);
    // 2025-01-06 is a Monday → first cell at the top-left of the grid.
    const c0 = s.nodes.find((n): n is RectNode => n.name === "cell-0")!;
    const c1 = s.nodes.find((n): n is RectNode => n.name === "cell-1")!; // Tuesday, same week
    expect(c1.x).toBeCloseTo(c0.x, 3); // same column (week)
    expect(c1.y).toBeGreaterThan(c0.y); // next weekday row down
  });

  it("advances a column across the week boundary", () => {
    const s = buildChart(base);
    const c0 = s.nodes.find((n): n is RectNode => n.name === "cell-0")!; // Mon wk0
    const c7 = s.nodes.find((n): n is RectNode => n.name === "cell-7")!; // +7 days = next Mon (wk1)
    expect(c7.x).toBeGreaterThan(c0.x);
    expect(c7.y).toBeCloseTo(c0.y, 3);
  });

  it("falls back to the matrix layout without date categories", () => {
    const s = buildChart({ ...base, data: { categories: ["Q1", "Q2"], series: [{ name: "R", values: [1, 2] }] } });
    expect(s.nodes.some((n) => n.name?.startsWith("weekday-"))).toBe(false);
    expect(s.nodes.some((n) => n.name === "cell-0-0")).toBe(true); // matrix cell naming
  });
});

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
  const rowName = (ri: number) => s.nodes.find((n): n is TextNode => n.kind === "text" && n.name === `row-${ri}`)!.text;

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
    const name0 = plain.nodes.find((n): n is TextNode => n.name === "row-0")!.text;
    expect(name0).toBe("Hi1");
    expect(plain.nodes.some((n) => n.name === "dendro-v")).toBe(false);
  });
});

describe("heatmap sign marks", () => {
  const hm = (
    heatmap: ChartConfig["heatmap"],
    vals: number[][] = [
      [8, -7],
      [-6, 9],
    ],
    extra: Partial<ChartConfig> = {},
  ): ChartConfig => ({
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
    expect(buildChart(hm({ sizeEncode: true })).nodes).toEqual(
      buildChart(hm({ sizeEncode: true, symbols: undefined })).nodes,
    );
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
    const c = hm(
      { symbols: "sign" },
      [
        [0, -7],
        [-6, 9],
      ],
      { decorations: { segmentLabels: false } },
    );
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
    expect(
      signs(
        hm({ sizeEncode: true, symbols: "sign" }, [
          [8, 7],
          [6, 9],
        ]),
      ),
    ).toHaveLength(0);
    expect(
      signs(
        hm({ sizeEncode: true, symbols: "sign" }, [
          [-8, -7],
          [-6, -9],
        ]),
      ),
    ).toHaveLength(0);
  });

  it("yields to the value label, which already prints the sign", () => {
    // Labels on and fitting => no marks: two signs in one cell is noise.
    expect(
      signs(
        hm(
          { symbols: "sign" },
          [
            [8, -7],
            [-6, 9],
          ],
          { decorations: { segmentLabels: true } },
        ),
      ),
    ).toHaveLength(0);
    // Labels off => the marks are the only sign carrier, so they appear.
    expect(
      signs(
        hm(
          { symbols: "sign" },
          [
            [8, -7],
            [-6, 9],
          ],
          { decorations: { segmentLabels: false } },
        ),
      ),
    ).toHaveLength(4);
  });

  it("inks each mark against its own cell, like the labels do", () => {
    for (const n of signs(hm({ sizeEncode: true, symbols: "sign" }))) {
      expect((n as { fill: string }).fill).toMatch(/^#(0b0b0b|ffffff)$/);
    }
  });

  it("keeps every mark inside the cell as DRAWN, however small sizeEncode makes it", () => {
    // A near-zero value shrinks its cell to almost nothing; a mark sized off the
    // slot rather than the drawn square would spill across its neighbours.
    const c = hm({ sizeEncode: true, symbols: "sign" }, [
      [100, -0.4],
      [-0.2, 100],
    ]);
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
    const c = hm({ sizeEncode: true, symbols: "sign" }, [
      [100, 0],
      [-0.001, 100],
    ]);
    // A zero-magnitude cell has no square to draw in.
    expect(buildChart(c).nodes.find((n) => n.name === "cell-sign-0-1")).toBeUndefined();
    const nulls: ChartConfig = {
      kind: "heatmap",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "R0", values: [8, null] },
          { name: "R1", values: [-6, 9] },
        ],
      },
      heatmap: { sizeEncode: true, symbols: "sign" },
    };
    expect(buildChart(nulls).nodes.find((n) => n.name === "cell-sign-0-1")).toBeUndefined();
  });
});

describe("heatmap forced-diverging on single-signed data draws no phantom zero", () => {
  it("omits the 0 tick when zero is outside the data range", () => {
    const cfg: ChartConfig = {
      kind: "heatmap",
      width: 640,
      height: 400,
      heatmap: { mode: "diverging" },
      data: { categories: ["A", "B", "C"], series: [{ name: "r", values: [10, 55, 100] }] },
    };
    expect(buildChart(cfg).nodes.some((n) => n.name === "legend-zero")).toBe(false);
  });

  it("still draws the 0 tick, on-strip, when the data spans zero", () => {
    const cfg: ChartConfig = {
      kind: "heatmap",
      width: 640,
      height: 400,
      heatmap: { mode: "diverging" },
      data: { categories: ["A", "B", "C"], series: [{ name: "r", values: [-40, 0, 60] }] },
    };
    const zero = buildChart(cfg).nodes.find((n) => n.name === "legend-zero") as TextNode | undefined;
    expect(zero).toBeDefined();
    expect(zero!.x).toBeGreaterThan(0);
  });
});

/**
 * PR8 replaced two hot loops with cached equivalents — the heatmap's row
 * clustering now reads a precomputed distance matrix instead of recomputing
 * each Euclidean distance on every merge scan, and the violin memoizes each
 * category's samples. Both are meant to be byte-identical; these tests pin the
 * observable behaviour so a future edit that changes the arithmetic is caught.
 */
describe("heatmap row clustering (precomputed distance matrix)", () => {
  // Two obvious groups: rows 0/1/2 are ~flat-low, rows 3/4/5 are ~flat-high.
  // Interleaved on input; average-linkage clustering must pull each group's
  // members adjacent in the rendered row order.
  const cfg: ChartConfig = {
    kind: "heatmap",
    width: 600,
    height: 400,
    heatmap: { cluster: true },
    data: {
      categories: ["c0", "c1", "c2", "c3"],
      series: [
        { name: "lowA", values: [1, 2, 1, 2] },
        { name: "highA", values: [90, 91, 89, 92] },
        { name: "lowB", values: [2, 1, 2, 1] },
        { name: "highB", values: [88, 90, 91, 89] },
        { name: "lowC", values: [1, 1, 2, 2] },
        { name: "highC", values: [92, 88, 90, 90] },
      ],
    },
  };

  it("reorders rows so similar rows are adjacent and draws the dendrogram", () => {
    const nodes = buildChart(cfg).nodes;
    const rowOrder = nodes
      .filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("row-"))
      .sort((a, b) => Number(a.name!.slice(4)) - Number(b.name!.slice(4)))
      .map((t) => t.text);
    // Every "low*" row must sit in one contiguous block and every "high*" in
    // the other — the clustering separated the two groups.
    const lowIdx = rowOrder.map((n, i) => (n.startsWith("low") ? i : -1)).filter((i) => i >= 0);
    const highIdx = rowOrder.map((n, i) => (n.startsWith("high") ? i : -1)).filter((i) => i >= 0);
    const contiguous = (idx: number[]) => idx[idx.length - 1] - idx[0] === idx.length - 1;
    expect(lowIdx).toHaveLength(3);
    expect(highIdx).toHaveLength(3);
    expect(contiguous(lowIdx)).toBe(true);
    expect(contiguous(highIdx)).toBe(true);
    // The dendrogram gutter is drawn.
    expect(nodes.some((n) => n.name === "dendro-v")).toBe(true);
  });

  it("is deterministic across rebuilds (no shared mutable state)", () => {
    expect(JSON.stringify(buildChart(cfg).nodes)).toBe(JSON.stringify(buildChart(cfg).nodes));
  });
});
