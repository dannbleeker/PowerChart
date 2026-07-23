import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart, valueExtent } from "../src/core/chart";
import { textWidth } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Cross-kind value-extent / auto-scale / layout-indexing invariants. */

/** Helper: the vertical span of every rect in a scene. */
const rectSpan = (scene: { nodes: any[] }) => {
  const ys = scene.nodes.flatMap((n) => (n.kind === "rect" ? [n.y, n.y + n.h] : []));
  return { top: Math.min(...ys), bottom: Math.max(...ys) };
};

describe("degenerate inputs do not throw", () => {
  it("boxplot with an empty series", () => {
    const cfg: ChartConfig = {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [] }] },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
  it("area chart with a single empty category set", () => {
    const cfg: ChartConfig = {
      kind: "area",
      ...DEFAULT_SIZE,
      decorations: { seriesLabels: true },
      data: { categories: [], series: [{ name: "S", values: [] }] },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
});

describe("value extents and auto-scales", () => {
  it("area extent honours the negative stack, like stacked (was floored at 0)", () => {
    const data = {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "A", values: [10, -40, 20] },
        { name: "B", values: [5, -30, 10] },
      ],
    };
    expect(valueExtent({ kind: "area", ...DEFAULT_SIZE, data } as ChartConfig)).toEqual({ min: -70, max: 30 });
    // identical data must give the same extent as the stacked sibling
    expect(valueExtent({ kind: "area", ...DEFAULT_SIZE, data } as ChartConfig)).toEqual(
      valueExtent({ kind: "stacked", ...DEFAULT_SIZE, data } as ChartConfig),
    );
  });

  it("negative small-multiples areas stay inside the scene", () => {
    const scene = buildChart({
      kind: "area",
      width: 480,
      height: 300,
      multiples: {},
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "A", values: [10, -40, 20] },
          { name: "B", values: [5, -30, 10] },
        ],
      },
    } as ChartConfig);
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1);
    expect(bottom).toBeLessThanOrEqual(301); // was ~775 — rendered far below the canvas
  });

  it("a target below the data range widens the scale down, not just up", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Actual", values: [10, 20, 30] },
          { name: "Target", values: [-50, -50, -50] },
        ],
      },
    } as ChartConfig);
    const ticks = scene.nodes.filter((n: any) => n.kind === "line" && n.name?.startsWith("target-"));
    expect(ticks.length).toBe(3);
    for (const t of ticks as any[]) {
      expect(t.y1).toBeGreaterThanOrEqual(0);
      expect(t.y1).toBeLessThanOrEqual(300); // target tick used to land outside the plot
    }
  });

  it("a multi-series waterfall combo keeps its columns on canvas", () => {
    const scene = buildChart({
      kind: "combo",
      ...DEFAULT_SIZE,
      combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "W1", values: [10, 10, 10] },
          { name: "W2", values: [10, 10, 10] },
          { name: "L", type: "line", values: [40, 40, 40] },
        ],
      },
    } as ChartConfig);
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1); // was -129 — stacked peak understated
    expect(bottom).toBeLessThanOrEqual(scene.height + 1);
  });

  it("small-multiples panels share one category order under categorySort", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 480,
      height: 300,
      multiples: { columns: 2 },
      categorySort: "descending",
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "S1", values: [1, 9, 5] },
          { name: "S2", values: [9, 1, 5] },
        ],
      },
      decorations: { categoryAxis: true },
    } as ChartConfig);
    const axisOf = (p: string) =>
      scene.nodes
        .filter((n: any) => n.kind === "text" && n.name?.startsWith(`${p}-category-`))
        .sort((a: any, b: any) => a.x - b.x)
        .map((n: any) => n.text);
    const p0 = axisOf("p0"),
      p1 = axisOf("p1");
    expect(p0.length).toBeGreaterThan(0);
    expect(p1).toEqual(p0); // panels used to rank by their own series and disagree
  });
});

describe("layout indexing", () => {
  it("grouped treemap tiles get the rectangle for their own value", () => {
    const scene = buildChart({
      kind: "treemap",
      width: 600,
      height: 400,
      data: { categories: ["G | a", "G | b", "G | c"], series: [{ name: "S", values: [1, 50, 100] }] },
    } as ChartConfig);
    const area = (name: string) => {
      const t = scene.nodes.find((n: any) => n.kind === "rect" && n.name === name) as any;
      return t ? t.w * t.h : 0;
    };
    // values 1 < 50 < 100 must give areas tile-0 < tile-1 < tile-2.
    // (Members are listed ascending, so the sort reorders them — which is
    // exactly when the post-sort-index lookup handed over the wrong rect.)
    expect(area("tile-0")).toBeLessThan(area("tile-1"));
    expect(area("tile-1")).toBeLessThan(area("tile-2"));
  });

  it("stacked connectors join the same series across a zero segment", () => {
    const scene = buildChart({
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [0, 10] },
          { name: "S2", values: [20, 30] },
        ],
      },
      decorations: { connectors: true },
    } as ChartConfig);
    const conns = scene.nodes.filter((n: any) => n.kind === "line" && n.name?.startsWith("connector-")) as any[];
    // S1 is absent in category A, so only S2's boundary can be joined: exactly one
    // connector, and it must link the two S2 tops (20 -> 30+... ), never S1<->S2.
    expect(conns).toHaveLength(1);
    expect(conns[0].name).toBe("connector-0-1"); // series index 1 = S2, not push-order 0
  });

  it("a tight-spread violin category still renders", () => {
    const scene = buildChart({
      kind: "violin",
      width: 480,
      height: 300,
      data: {
        categories: ["Wide", "Tight"],
        series: [
          { name: "o1", values: [0, 3] },
          { name: "o2", values: [30, 3.1] },
          { name: "o3", values: [60, 3] },
          { name: "o4", values: [90, 3.05] },
          { name: "o5", values: [120, 3.02] },
        ],
      },
    } as ChartConfig);
    // Both categories have valid observations; neither may be silently dropped.
    expect(scene.nodes.some((n: any) => n.name === "violin-0")).toBe(true);
    expect(scene.nodes.some((n: any) => n.name === "violin-1")).toBe(true);
  });

  it("horizontal mekko hides labels in rows thinner than the font", () => {
    const cats = Array.from({ length: 16 }, (_, i) => `C${i + 1}`);
    const cfg = {
      kind: "mekko",
      width: 600,
      height: 160,
      horizontal: true,
      data: {
        categories: cats,
        series: [
          { name: "S1", values: cats.map(() => 50) },
          { name: "S2", values: cats.map(() => 50) },
        ],
      },
      decorations: { segmentLabels: true },
    } as ChartConfig;
    const scene = buildChart(cfg);
    const segs = scene.nodes.filter((n: any) => n.kind === "rect" && n.name?.startsWith("seg-")) as any[];
    const labels = scene.nodes.filter((n: any) => n.kind === "text" && n.name?.startsWith("label-")) as any[];
    // Rows here are ~6.6pt thick — far under the 11pt font. The old gate measured
    // the segment's 250pt value-axis length instead, so it stamped a label into
    // every one of them.
    expect(segs.length).toBe(32);
    expect(segs.every((s) => s.h < 11 * 1.25)).toBe(true);
    expect(labels).toHaveLength(0);
    // Vertical mekko is unaffected: its gate always measured r.h.
    const tall = buildChart({ ...cfg, horizontal: false } as ChartConfig);
    expect(tall.nodes.some((n: any) => n.kind === "text" && n.name?.startsWith("label-"))).toBe(true);
  });

  it("horizontal mekko keeps labels out of segments shorter than the font", () => {
    // Thick rows, hairline-short segments — the case the thickness gate alone
    // does not catch. The fit check tolerates 2pt of bleed (the text box is
    // drawn 4pt wider than the segment), which is fine for a vertical mekko's
    // wide columns but let three 3.8pt segments each print a 5.4pt "4",
    // overlapping their neighbours by 1.6pt.
    const scene = buildChart({
      kind: "mekko",
      width: 400,
      height: 300,
      horizontal: true,
      data: {
        categories: ["EMEA"],
        series: [
          { name: "A", values: [4] },
          { name: "B", values: [4] },
          { name: "C", values: [4] },
          { name: "D", values: [300] },
        ],
      },
    } as ChartConfig);
    const ink = (n: any) => {
      const w = textWidth(n.text, n.fontSize);
      const cx = n.x + n.w / 2;
      return { lo: cx - w / 2, hi: cx + w / 2 };
    };
    const labels = scene.nodes.filter((n: any) => n.kind === "text" && n.name?.startsWith("label-")) as any[];
    // Only D is wide enough to carry a label.
    expect(labels.map((l) => l.text)).toEqual(["300"]);
    // And no two labels may ever overlap.
    const inks = labels.map(ink).sort((a, b) => a.lo - b.lo);
    for (let i = 1; i < inks.length; i++) expect(inks[i].lo).toBeGreaterThanOrEqual(inks[i - 1].hi);
  });
});

describe("valueExtent reports what the layout draws", () => {
  /** What "Same scale" does: write the extent back as a hard scale override. */
  const underSameScale = (cfg: ChartConfig) => {
    const e = valueExtent(cfg)!;
    return buildChart({ ...cfg, scale: { min: e.min < 0 ? e.min : undefined, max: e.max } });
  };
  const inkSpan = (scene: { nodes: any[] }) => {
    const ys = scene.nodes.flatMap((n) =>
      n.kind === "line" ? [n.y1, n.y2] : n.kind === "rect" ? [n.y, n.y + n.h] : [],
    );
    return { top: Math.min(...ys), bottom: Math.max(...ys) };
  };

  it("treats an Error row as a whisker, not as a data point", () => {
    const cfg = {
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A"],
        series: [
          { name: "S", values: [10] },
          { name: "Error", values: [30] },
        ],
      },
    } as ChartConfig;
    // The whisker spans 10±30. The old extent was {0,30}: 30 was the Error row's
    // own magnitude mistaken for a value — neither the data range nor the drawn one.
    const ext = valueExtent(cfg)!;
    expect(ext.max).toBeGreaterThanOrEqual(40);
    expect(ext.min).toBeLessThanOrEqual(-20);
    const { top, bottom } = inkSpan(underSameScale(cfg));
    expect(top).toBeGreaterThanOrEqual(-1); // whisker used to reach y=-83
    expect(bottom).toBeLessThanOrEqual(301); // and y=465
  });

  it("covers a waterfall's Target row", () => {
    const cfg = {
      kind: "waterfall",
      width: 480,
      height: 300,
      waterfall: { totalIndices: [2] },
      data: {
        categories: ["Start", "Up", "End"],
        series: [
          { name: "V", values: [100, 20, 0] },
          { name: "Target", values: [null, null, 200] },
        ],
      },
    } as ChartConfig;
    expect(valueExtent(cfg)!.max).toBeGreaterThanOrEqual(200); // was 120 — the running total only
    const { bottom } = inkSpan(underSameScale(cfg));
    expect(bottom).toBeLessThanOrEqual(301);
  });

  it("does not sum a Target row into a stack", () => {
    const cfg = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A"],
        series: [
          { name: "S", values: [10] },
          { name: "Target", values: [5] },
        ],
      },
    } as ChartConfig;
    // The Target is a tick at 5, not another 5pt of stack: the column totals 10.
    expect(valueExtent(cfg)).toEqual({ min: 0, max: 10 });
  });

  it("covers an explicit threshold line above the data", () => {
    const cfg = {
      kind: "clustered",
      width: 480,
      height: 300,
      decorations: { valueLines: [{ mode: "value", value: 500 }] },
      data: { categories: ["A"], series: [{ name: "S", values: [10] }] },
    } as ChartConfig;
    expect(valueExtent(cfg)!.max).toBeGreaterThanOrEqual(500); // was 10
  });

  it("a mean value line needs no widening — it is inside the data by construction", () => {
    const plain = {
      kind: "clustered",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 20] }] },
    } as ChartConfig;
    const withMean = { ...plain, decorations: { valueLines: [{ mode: "mean" }] } } as ChartConfig;
    expect(valueExtent(withMean)).toEqual(valueExtent(plain));
  });
});
