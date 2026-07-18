import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { layoutViolin } from "../src/core/layout/violin";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

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
