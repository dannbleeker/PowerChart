import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Category sorting and the automatic Other bucket. */

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

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

/**
 * A reorder moves `colors` because "a highlight belongs to a data point, not to
 * a screen position" (test/palette.test.ts states the same contract). The
 * decorations were copied through untouched, so a callout and its arrows stayed
 * on the SLOT the category used to occupy — and the arrows, which read
 * `columnValue` at those indices, printed a number for a pair the author never
 * named.
 */
describe("decorations travel with a reordered category", () => {
  const data = { categories: ["Small", "Big", "Mid"], series: [{ name: "S", values: [10, 90, 50] }] };
  const centerOf = (scene: ReturnType<typeof buildChart>, name: string) => {
    const n = scene.nodes.find((x) => x.name === name) as RectNode;
    return n.x + n.w / 2;
  };
  const categoryCenter = (scene: ReturnType<typeof buildChart>, text: string) => {
    const n = (byName(scene.nodes, "category-") as TextNode[]).find((c) => c.text === text)!;
    return n.x + n.w / 2;
  };

  it.each(["descending", "ascending"] as const)("keeps a callout on its own column (%s)", (categorySort) => {
    const scene = buildChart(
      cfg({
        categorySort,
        data,
        decorations: { categoryAxis: true, callouts: [{ category: 1, text: "peak" }] },
      }),
    );
    // Category 1 is "Big" — wherever the sort puts it, the bubble goes with it.
    expect(centerOf(scene, "callout-box-0")).toBeCloseTo(categoryCenter(scene, "Big"), 0);
  });

  it("computes CAGR from the columns the author named, not the slots they vacated", () => {
    const scene = buildChart(cfg({ categorySort: "descending", data, decorations: { cagr: { from: 0, to: 2 } } }));
    // Categories 0 and 2 are Small (10) and Mid (50). Descending puts Mid before
    // Small, and the arrow reads left-to-right, so the rate is 50 → 10 over one
    // period. Reading slots 0 and 2 instead gave the rate between Big and Small.
    const label = (scene.nodes as TextNode[]).find((n) => n.name === "cagr-label")!;
    expect(label.text).toBe("-80.0% p.a.");
  });

  it("moves a Pareto chart's callout with the ranked bar", () => {
    const scene = buildChart(
      cfg({
        kind: "clustered",
        pareto: true,
        data,
        decorations: { categoryAxis: true, callouts: [{ category: 0, text: "tail" }] },
      }),
    );
    // Category 0 is "Small" — last in a Pareto ranking.
    expect(centerOf(scene, "callout-box-0")).toBeCloseTo(categoryCenter(scene, "Small"), 0);
  });
});

/**
 * `pie.explode` and `pie.breakout` are category indices too — they just live on
 * the top-level config rather than under `decorations`, so `remapDecorations`
 * never saw them. Unpermuted, the highlight followed the SCREEN SLOT: with
 * `categorySort` on, `explode: [0]` offset whichever category the sort put
 * first, and `breakout: [0, 2]` collapsed a different pair of data points than
 * the author named. Pie and doughnut are both in `SORTABLE`, so this is
 * reachable from the pane.
 */
describe("a pie highlight belongs to its data point, not its slot", () => {
  const cfg = (extra: Record<string, unknown>): ChartConfig =>
    ({
      kind: "pie",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Alpha", "Beta", "Gamma", "Delta"],
        series: [{ name: "Share", values: [10, 40, 20, 30] }],
      },
      ...extra,
    }) as unknown as ChartConfig;

  /** Index of the one wedge that is offset from the common centre. */
  const explodedSlot = (c: ChartConfig) => {
    const wedges = buildChart(c).nodes.filter((n) => n.kind === "wedge") as unknown as { cx: number }[];
    const common = wedges.map((w) => w.cx).sort((a, b) => a - b)[Math.floor(wedges.length / 2)];
    return wedges.findIndex((w) => Math.abs(w.cx - common) > 0.5);
  };

  it("explodes the same category before and after sorting", () => {
    // Descending by value is Beta(40), Delta(30), Gamma(20), Alpha(10) — so
    // Alpha, named as index 0, is drawn in slot 3.
    expect(explodedSlot(cfg({ pie: { explode: [0] } }))).toBe(0);
    expect(explodedSlot(cfg({ pie: { explode: [0] }, categorySort: "descending" }))).toBe(3);
  });

  it("breaks out the same categories before and after sorting", () => {
    const broken = (c: ChartConfig) =>
      buildChart(c)
        .nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("breakout-label"))
        .map((n) => n.text);
    const flat = broken(cfg({ pie: { breakout: [0, 2] } }));
    const sorted = broken(cfg({ pie: { breakout: [0, 2] }, categorySort: "descending" }));
    expect(flat.length, "no breakout labels, so this proves nothing").toBeGreaterThan(0);
    expect(sorted).toEqual(flat);
  });
});
