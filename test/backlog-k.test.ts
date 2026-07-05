import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";

/**
 * Backlog batch K — more §2 within-kind gaps: automatic "Other" bucket,
 * heatmap calendar layout, butterfly stacked flanks.
 */

describe("automatic Other bucket", () => {
  const base: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Y1", "Y2"],
      series: [
        { name: "Big", values: [50, 60] },
        { name: "Mid", values: [30, 32] },
        { name: "Small A", values: [8, 9] },
        { name: "Small B", values: [5, 6] },
        { name: "Small C", values: [3, 4] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("keeps max−1 largest series and sums the rest into Other", () => {
    const s = buildChart({ ...base, otherBucket: { max: 3 } });
    const labels = s.nodes.filter((n): n is TextNode => n.kind === "text" && n.name?.startsWith("series-label") === true).map((n) => n.text);
    // Series labels aren't on; assert by segment count instead: 3 series → 3 segments per category.
    const segs0 = s.nodes.filter((n) => n.name?.match(/^seg-\d+-0$/));
    expect(segs0).toHaveLength(3); // Big, Mid, Other
    void labels;
  });

  it("Other sums the collapsed tail", () => {
    const s = buildChart({ ...base, otherBucket: { max: 3 }, decorations: { seriesLabels: true, segmentLabels: false } });
    // Other = Small A+B+C at Y1 = 8+5+3 = 16; it is the top (last) segment.
    const other = s.nodes.find((n): n is TextNode => n.kind === "text" && n.text === "Other");
    expect(other).toBeTruthy();
    // Top segment height corresponds to 16 vs Big 50 → ratio ~0.32.
    const segNames = [0, 1, 2].map((i) => `seg-${i}-0`);
    const heights = segNames.map((nm) => (buildChart({ ...base, otherBucket: { max: 3 } }).nodes.find((n): n is RectNode => n.name === nm) as RectNode).h);
    // seg-2 is Other (16), seg-0 is Big (50).
    expect(heights[2] / heights[0]).toBeCloseTo(16 / 50, 1);
  });

  it("no-op when already within budget, or for non-column kinds", () => {
    const within = buildChart({ ...base, otherBucket: { max: 8 } });
    expect(within.nodes.some((n) => n.name === "seg-4-0")).toBe(true); // all 5 kept
    const line = buildChart({ ...base, kind: "line", otherBucket: { max: 2 } });
    expect(line.nodes.some((n) => n.name?.startsWith("marker-4-"))).toBe(true); // 5 lines intact
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

describe("butterfly stacked flanks", () => {
  const cfg: ChartConfig = {
    kind: "butterfly",
    ...DEFAULT_SIZE,
    butterfly: { split: 2 },
    data: {
      categories: ["A", "B"],
      series: [
        { name: "L1", values: [10, 12] },
        { name: "L2", values: [6, 8] },
        { name: "R1", values: [14, 9] },
        { name: "R2", values: [4, 5] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("stacks split series on the left and the rest on the right", () => {
    const s = buildChart(cfg);
    // Four segments per category (2 left + 2 right).
    expect(s.nodes.filter((n) => n.name?.match(/^seg-\d+-0$/))).toHaveLength(4);
    const l1 = s.nodes.find((n): n is RectNode => n.name === "seg-0-0")!; // L1 innermost-left
    const l2 = s.nodes.find((n): n is RectNode => n.name === "seg-1-0")!; // L2 stacked further left
    expect(l2.x + l2.w).toBeCloseTo(l1.x, 1); // L2 sits just left of L1 (contiguous)
    const r1 = s.nodes.find((n): n is RectNode => n.name === "seg-2-0")!; // R1 innermost-right
    const r2 = s.nodes.find((n): n is RectNode => n.name === "seg-3-0")!;
    expect(r2.x).toBeCloseTo(r1.x + r1.w, 1); // R2 just right of R1
    // Stacked mode shows a legend of every series.
    expect(s.nodes.some((n) => n.name === "legend-3")).toBe(true);
  });

  it("default (no split) keeps the classic two-series butterfly", () => {
    const s = buildChart({ ...cfg, butterfly: undefined });
    expect(s.nodes.some((n) => n.name === "seg-2-0")).toBe(false); // series 2+ ignored
    expect(s.nodes.some((n) => n.name === "header-0")).toBe(true); // two headers, not a legend
  });
});
