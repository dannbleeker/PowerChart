import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Category sorting and the automatic Other bucket. */

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

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
    const labels = s.nodes
      .filter((n): n is TextNode => n.kind === "text" && n.name?.startsWith("series-label") === true)
      .map((n) => n.text);
    // Series labels aren't on; assert by segment count instead: 3 series → 3 segments per category.
    const segs0 = s.nodes.filter((n) => n.name?.match(/^seg-\d+-0$/));
    expect(segs0).toHaveLength(3); // Big, Mid, Other
    void labels;
  });

  it("Other sums the collapsed tail", () => {
    const s = buildChart({
      ...base,
      otherBucket: { max: 3 },
      decorations: { seriesLabels: true, segmentLabels: false },
    });
    // Other = Small A+B+C at Y1 = 8+5+3 = 16; it is the top (last) segment.
    const other = s.nodes.find((n): n is TextNode => n.kind === "text" && n.text === "Other");
    expect(other).toBeTruthy();
    // Top segment height corresponds to 16 vs Big 50 → ratio ~0.32.
    const segNames = [0, 1, 2].map((i) => `seg-${i}-0`);
    const heights = segNames.map(
      (nm) =>
        (buildChart({ ...base, otherBucket: { max: 3 } }).nodes.find((n): n is RectNode => n.name === nm) as RectNode)
          .h,
    );
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

describe("categorySort excludes carried rows", () => {
  it("ranks by real stack totals, ignoring Error/Target rows", () => {
    // Category B has the larger real total (30 vs 20) but a huge Target on A.
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      categorySort: "descending",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Value", values: [20, 30] },
          { name: "Target", values: [999, 1] },
        ],
      },
    };
    const scene = buildChart(cfg);
    // Descending by real total → B (30) before A (20). Find the category axis
    // labels in order.
    const labels = scene.nodes
      .filter((n) => n.kind === "text" && (n as any).name?.startsWith("category-"))
      .map((n) => (n as any).text);
    expect(labels[0]).toBe("B");
    expect(labels[1]).toBe("A");
  });
});

describe("category sorting", () => {
  it("sorts categories by total, descending", () => {
    const scene = buildChart(
      cfg({
        categorySort: "descending",
        data: {
          categories: ["Small", "Big", "Mid"],
          series: [{ name: "S", values: [10, 100, 50] }],
        },
        decorations: { categoryAxis: true },
      }),
    );
    const cats = byName(scene.nodes, "category-") as TextNode[];
    expect(cats.map((c) => c.text)).toEqual(["Big", "Mid", "Small"]);
  });

  it("leaves order-sensitive kinds untouched", () => {
    const scene = buildChart(
      cfg({
        kind: "waterfall",
        categorySort: "descending",
        data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 100] }] },
      }),
    );
    const cats = byName(scene.nodes, "category-") as TextNode[];
    expect(cats.map((c) => c.text)).toEqual(["A", "B"]);
  });
});
