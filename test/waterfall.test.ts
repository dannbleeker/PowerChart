import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart, valueExtent } from "../src/core/chart";
import { layoutWaterfall } from "../src/core/layout/waterfall";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { LineNode, RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Waterfall — chains, budget-vs-actual, grouping spacers, stacked, deltas across zero. */

/** Helper: the vertical span of every rect in a scene. */
const rectSpan = (scene: { nodes: any[] }) => {
  const ys = scene.nodes.flatMap((n) => (n.kind === "rect" ? [n.y, n.y + n.h] : []));
  return { top: Math.min(...ys), bottom: Math.max(...ys) };
};

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */
const W = 480;

const H = 300;

const texts = (nodes: SceneNode[], namePrefix: string) =>
  nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith(namePrefix)).map((n) => n.text);

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

/**
 * Two combo line-overlay scale bugs found by a layout-modes hunt: both put a line
 * series off the canvas (finite but wrong coordinates). Byte-identical for the
 * valid cases (positive line / plain-chain waterfall) — snapshots unchanged — and
 * only rescue the broken configs.
 */
const HEIGHT = 300;

const markerYs = (cfg: ChartConfig) =>
  buildChart(cfg)
    .nodes.filter((n) => n.name?.startsWith("combo-marker") || n.name?.startsWith("combo-line"))
    .flatMap((n) => (n.kind === "rect" ? [n.y, n.y + n.h] : n.kind === "line" ? [n.y1, n.y2] : []));

/** Backlog batch D: waterfall gap-to-target, bands, fillBetween, heatmap totals. */
describe("waterfall budget-vs-actual (Target row)", () => {
  const cfg: ChartConfig = {
    kind: "waterfall",
    ...DEFAULT_SIZE,
    data: {
      categories: ["FY23", "Volume", "Price", "FY24"],
      series: [
        { name: "Delta", values: [80, 10, 6, 0] },
        { name: "Target", values: [null, null, null, 110] },
      ],
    },
    waterfall: { totalIndices: [3] },
  };
  const s = buildChart(cfg);

  it("draws the target tick and a hatched gap segment on the final total", () => {
    const tick = s.nodes.find((n) => n.name === "target-3") as LineNode;
    expect(tick).toBeDefined();
    const gap = s.nodes.find((n) => n.name === "target-gap-3") as RectNode;
    expect(gap.pattern).toBe("diagonal");
    // Gap spans actual (96) → target (110): its bottom meets the bar top.
    const bar = s.nodes.find((n) => n.name === "bar-3") as RectNode;
    expect(gap.y + gap.h).toBeCloseTo(bar.y, 1);
    const label = s.nodes.find((n) => n.name === "target-gap-label-3") as TextNode;
    expect(label.text).toBe("Gap +14");
  });

  it("Target row never renders as waterfall bars and the scale covers it", () => {
    expect(s.nodes.filter((n) => n.name?.startsWith("bar-"))).toHaveLength(4);
    const gap = s.nodes.find((n) => n.name === "target-gap-3") as RectNode;
    expect(gap.y).toBeGreaterThan(0); // inside the plot — scale widened
  });

  it("a missed target (actual above target) hatches downward with a negative gap", () => {
    const over = buildChart({
      ...cfg,
      data: { ...cfg.data, series: [cfg.data.series[0], { name: "Target", values: [null, null, null, 90] }] },
    });
    expect((over.nodes.find((n) => n.name === "target-gap-label-3") as TextNode).text).toBe("Gap -6");
  });
});

describe("waterfall grouping spacers", () => {
  const cfg: ChartConfig = {
    kind: "waterfall",
    ...DEFAULT_SIZE,
    data: {
      categories: ["FY23", "Volume", "Price", "", "Cost", "FX", "", "FY24"],
      series: [{ name: "Delta", values: [86, 14, 9, null, -12, -4, null, 0] }],
    },
    waterfall: { totalIndices: [7], spacerIndices: [3, 6] },
    decorations: { categoryAxis: true },
  };
  const s = buildChart(cfg);

  it("draws no bar in a spacer slot", () => {
    expect(s.nodes.some((n) => n.name === "bar-3")).toBe(false);
    expect(s.nodes.some((n) => n.name === "bar-6")).toBe(false);
    expect(s.nodes.some((n) => n.name === "bar-1")).toBe(true);
    expect(s.nodes.some((n) => n.name === "bar-4")).toBe(true);
  });

  it("carries the running total across the gap (final total unaffected)", () => {
    // 86 + 14 + 9 − 12 − 4 = 93.
    const total = s.nodes.find((n): n is TextNode => n.kind === "text" && n.name === "label-7");
    expect(total?.text).toBe("93");
  });

  it("bridges the connector across the spacer slot", () => {
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && n.name === "spacer-bridge-3")).toBe(true);
    expect(s.nodes.some((n) => n.name === "spacer-bridge-6")).toBe(true);
  });

  it("without spacers, those categories render normally", () => {
    const plain = buildChart({ ...cfg, waterfall: { totalIndices: [7] } });
    expect(plain.nodes.some((n) => n.name?.startsWith("spacer-bridge"))).toBe(false);
  });
});

/**
 * Backlog batch M — combo family: waterfall/mekko base under a line overlay,
 * and independent per-line-series axes.
 */
describe("combo waterfall base", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    secondaryAxis: true,
    combo: { columns: "waterfall" },
    waterfall: { totalIndices: [3] },
    data: {
      categories: ["Start", "Up", "Down", "End"],
      series: [
        { name: "Delta", values: [100, 20, -15, 0] },
        { name: "Margin %", type: "line", values: [40, 42, 38, 41] },
      ],
    },
    decorations: { segmentLabels: false, seriesLabels: false },
  };

  it("draws waterfall bars with a line overlaid on the secondary axis", () => {
    const s = buildChart(cfg);
    // Waterfall base bars.
    expect(s.nodes.some((n) => n.name === "bar-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "bar-3")).toBe(true);
    // Line overlay + secondary axis ticks.
    expect(s.nodes.filter((n) => n.name?.startsWith("combo-marker-0-"))).toHaveLength(4);
    expect(s.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
  });
});

describe("combo waterfall base — shared-axis line overflow/underflow", () => {
  const mk = (line: number[]): ChartConfig => ({
    kind: "combo",
    ...DEFAULT_SIZE,
    combo: { columns: "waterfall" }, // shared axis (no secondaryAxis)
    data: {
      categories: ["A", "B", "C", "D"],
      series: [
        { name: "Cols", values: [100, 20, 30, 10] }, // running peak 160, trough 0
        { name: "Line", type: "line", values: line },
      ],
    },
    decorations: { segmentLabels: false, seriesLabels: false },
  });

  const lineYs = (s: ReturnType<typeof buildChart>) =>
    s.nodes
      .filter((n) => n.name?.startsWith("combo-marker-0-"))
      .map((n) => (n.kind === "rect" ? n.y : n.kind === "ellipse" ? n.cy : NaN))
      .filter((y) => !Number.isNaN(y));

  it("keeps a line that dips below the waterfall trough on-plot", () => {
    // #157 stretched the axis MAX for an overflowing line but left the floor at 0,
    // so a negative line point plotted off the bottom of the plot.
    const s = buildChart(mk([-80, -40, 10, 50]));
    const ys = lineYs(s);
    expect(ys.length).toBe(4);
    for (const y of ys) expect(y).toBeLessThanOrEqual(DEFAULT_SIZE.height + 1); // not off the bottom
  });

  it("still keeps a line that overshoots the peak on-plot (unchanged)", () => {
    const s = buildChart(mk([200, 40, 30, 10]));
    for (const y of lineYs(s)) expect(y).toBeGreaterThanOrEqual(-1); // not off the top
  });
});

describe("waterfall extent walks the same chain the layout draws", () => {
  const stacked = {
    kind: "waterfall",
    ...DEFAULT_SIZE,
    waterfall: { totalIndices: [3] },
    data: {
      categories: ["FY23", "Organic", "M&A", "FY24"],
      series: [
        { name: "Europe", values: [50, 8, 5, 0] },
        { name: "Americas", values: [36, 6, 9, 0] },
      ],
    },
  } as ChartConfig;

  it("counts every stacked series, not just the first", () => {
    // The bridge runs (50+36) + (8+6) + (5+9) = 114. A second chain that added
    // only series[0] reported 63 — the first series' total, called the chart's.
    expect(valueExtent(stacked)).toEqual({ min: 0, max: 114 });
  });

  it("keeps a stacked bridge on the shape under Same scale", () => {
    const e = valueExtent(stacked)!;
    const scene = buildChart({ ...stacked, scale: { min: e.min < 0 ? e.min : undefined, max: e.max } });
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1); // was -213.8 on a 300pt canvas
    expect(bottom).toBeLessThanOrEqual(scene.height + 1);
  });

  it("carries the running total across a spacer, as the bars do", () => {
    // The old extent ignored spacerIndices entirely; it only got away with it
    // because a spacer's cell is usually null.
    const withSpacer = {
      kind: "waterfall",
      ...DEFAULT_SIZE,
      waterfall: { totalIndices: [4], spacerIndices: [2] },
      data: {
        categories: ["Start", "Up", "", "Up2", "End"],
        series: [{ name: "V", values: [40, 10, 999, 10, 0] }],
      },
    } as ChartConfig;
    // The spacer draws no bar and must not advance the total: 40+10+10 = 60.
    expect(valueExtent(withSpacer)).toEqual({ min: 0, max: 60 });
  });
});

describe('waterfall "of which" detail groups', () => {
  /** Cost (-12) decomposed into Labour/Freight/Energy (-7/-3/-2). */
  const bridge = (detail: boolean) =>
    ({
      kind: "waterfall",
      width: 560,
      height: 300,
      data: {
        categories: ["FY23", "Volume", "Cost", "> Labour", "> Freight", "> Energy", "FX", "FY24"],
        series: [{ name: "Delta", values: [86, 14, -12, -7, -3, -2, -4, 0] }],
      },
      waterfall: {
        totalIndices: [7],
        ...(detail ? { detailGroups: [{ of: 2, indices: [3, 4, 5] }] } : {}),
      },
    }) as ChartConfig;
  const node = (c: ChartConfig, name: string) => buildChart(c).nodes.find((n) => n.name === name) as any;

  it("keeps detail columns off the chain, so the totals stay right", () => {
    // 86 + 14 - 12 - 4 = 84. Without the grouping the details join the walk and
    // the same rows total 72 — the breakdown counted twice.
    expect(node(bridge(true), "label-7").text).toBe("84");
    expect(node(bridge(false), "label-7").text).toBe("72");
  });

  it("decomposes the parent's delta, from the parent's own base", () => {
    const c = bridge(true);
    const cost = node(c, "bar-2");
    const energy = node(c, "bar-5"); // the last detail
    const labour = node(c, "bar-3"); // the first
    // The sub-bridge starts where Cost starts and ends where Cost ends: the
    // group IS that column taken apart, not more steps in the walk.
    expect(labour.y).toBeCloseTo(cost.y, 6);
    expect(energy.y + energy.h).toBeCloseTo(cost.y + cost.h, 6);
  });

  it("steps the connector over the group, without burying it", () => {
    const c = bridge(true);
    const names = buildChart(c)
      .nodes.filter((n) => n.name?.startsWith("connector-"))
      .map((n) => n.name);
    // A detail has no outgoing level to carry, so it draws no connector.
    expect(names).toEqual(["connector-0", "connector-1", "connector-2", "connector-6"]);
    // The parent's connector reaches the next CHAIN column, not the next index.
    const conn = node(c, "connector-2");
    const fx = node(c, "bar-6");
    expect(conn.x2).toBeLessThanOrEqual(fx.x + 0.01);
    expect(conn.x2 - conn.x1).toBeGreaterThan(150); // it spans the whole group
    // Anchoring the sub-bridge at the parent's base (rather than at the level
    // the chain carries) is what keeps this line clear of the bars it skips.
    for (const i of [3, 4, 5]) {
      const b = node(c, `bar-${i}`);
      const through = conn.y1 > b.y && conn.y1 < b.y + b.h;
      expect(through, `bar-${i}`).toBe(false);
    }
  });

  it("renders a group that does not sum to its parent exactly as authored", () => {
    // The engine draws your numbers; it does not reconcile them. The chain is
    // unaffected either way, so the totals cannot silently drift.
    const c = bridge(true);
    (c.data.series[0].values as (number | null)[])[3] = -1; // 1+3+2 != 12
    expect(node(c, "label-7").text).toBe("84");
    expect(node(c, "bar-3")).toBeTruthy();
  });

  it("covers the detail bars in the value extent", () => {
    expect(valueExtent(bridge(true))!.max).toBeGreaterThanOrEqual(100);
  });
});

describe("waterfall deltas keep their sign across a zero crossing", () => {
  it("signs a mid-chain delta whose incoming running total is exactly 0", () => {
    const cfg: ChartConfig = {
      kind: "waterfall",
      width: W,
      height: H,
      data: { categories: ["Opening", "Repayment", "New loan"], series: [{ name: "S", values: [100, -100, 50] }] },
    };
    // Before: ["100", "-100", "50"] — the +50 lost its sign because its stack
    // happened to start at 0, which the old gate read as a base column.
    expect(texts(buildChart(cfg).nodes, "label-")).toEqual(["100", "-100", "+50"]);
  });
});

describe("stacked waterfall", () => {
  const c = cfg({
    kind: "waterfall",
    data: {
      categories: ["Base", "Growth", "End"],
      series: [
        { name: "EU", values: [50, 8, 0] },
        { name: "US", values: [30, -6, 0] },
      ],
    },
    waterfall: { totalIndices: [2] },
  });
  const { nodes, anchors } = layoutWaterfall(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("stacks per-series contributions and moves the running level by the column sum", () => {
    expect(anchors.columnValue).toEqual([80, 82, 82]); // 50+30, +8-6, total
    const segs = byName(nodes, "bar-1-s") as RectNode[];
    expect(segs).toHaveLength(2);
  });

  it("colors segments by series in stacked mode", () => {
    const [eu] = byName(nodes, "bar-1-s0") as RectNode[];
    expect(eu.fill).toBe(DEFAULT_STYLE.palette[0]);
  });
});

describe("combo-waterfall with an overflowing line", () => {
  it("stretches the shared axis so the tall line stays inside the plot", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      ...DEFAULT_SIZE,
      combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "W", values: [10, 10, 10] }, // cumulative peak = 30
          { name: "L", type: "line", values: [5, 60, 5] }, // spikes to 60
        ],
      },
    };
    const scene = buildChart(cfg);
    const markers = scene.nodes.filter((n) => (n.name ?? "").startsWith("combo-marker"));
    expect(markers.length).toBe(3);
    // None of the markers escapes above the top of the canvas.
    expect(Math.min(...markers.map((m) => (m as { y: number }).y))).toBeGreaterThanOrEqual(0);
  });
});

describe("combo waterfall wfMax ignores off-chain columns", () => {
  it("keeps a tall shared-axis line on-plot over a detail-group waterfall", () => {
    // detailGroups[2,3] decompose column 1 in place — they do NOT advance the
    // chain (peak = 150). The old wfMax summed them (200), matched the line's 200,
    // so the overflow stretch never fired and the line clipped off the top.
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: HEIGHT,
      combo: { columns: "waterfall" },
      waterfall: { detailGroups: [{ of: 1, indices: [2, 3] }] },
      data: {
        categories: ["a", "b", "c", "d", "e"],
        series: [
          { name: "delta", values: [100, 50, 30, 20, null] },
          { name: "line", type: "line", values: [200, 200, 200, 200, 200] },
        ],
      },
    };
    const ys = markerYs(cfg);
    expect(ys.length).toBeGreaterThan(0);
    for (const y of ys) expect(y).toBeGreaterThanOrEqual(-1);
  });
});
