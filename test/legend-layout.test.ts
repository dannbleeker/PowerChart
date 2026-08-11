import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { legendRow } from "../src/core/layout/column";
import { DEFAULT_STYLE } from "../src/core/style";
import type { RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Legend layout — wrapping, reserved rows, chip mirrors the mark. */

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */
const W = 480;

const H = 300;

describe("legend chips are a miniature of the mark they label", () => {
  const paintedCfg: ChartConfig = {
    kind: "stacked",
    horizontal: true,
    width: W,
    height: H,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "Retail", values: [10, 20], color: "#2a78d6" },
        { name: "Online", values: [5, 8], color: "#2a78d6", pattern: "diagonal" },
        { name: "Plan", values: [4, 6], scenario: "PL" },
      ],
    },
    decorations: { seriesLabels: true },
  };

  it("carries pattern and the IBCS scenario restyle onto the chip", () => {
    const { nodes } = buildChart(paintedCfg);
    const chip = (i: number) => nodes.find((n) => n.name === `legend-chip-${i}`) as RectNode;
    const seg = (i: number) => nodes.find((n) => n.name === `seg-${i}-0`) as RectNode;
    // Before: three identical solid squares — two same-coloured series told
    // apart only by a hatch, and a hollow PL bar keyed by a solid block.
    // (The chip carries no separator stroke, so only the scenario restyle's
    // own outline is compared against the segment.)
    for (const i of [0, 1, 2]) {
      expect(chip(i).fill).toBe(seg(i).fill);
      expect(chip(i).pattern).toBe(seg(i).pattern);
    }
    expect(chip(0).stroke).toBeUndefined();
    expect(chip(1).pattern).toBe("diagonal");
    expect(chip(2).fill).toBe("none");
    expect(chip(2).stroke).toBe(seg(2).stroke);
    expect(chip(2).strokeWidth).toBe(seg(2).strokeWidth);
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

describe("a wrapped legend reserves its rows and never overlaps the plot", () => {
  // Horizontal bars, a narrow canvas, and six long series names: the legend
  // cannot fit one row, so #139's wrap advances it downward. Until the frame
  // reserved those extra rows, the wrapped rows painted on top of the bars.
  const cfg: ChartConfig = {
    kind: "stacked",
    horizontal: true,
    width: 360,
    height: 320,
    data: {
      categories: ["North", "South"],
      series: [
        { name: "Northern Europe Wholesale", values: [12, 9] },
        { name: "Central Europe Retail", values: [8, 11] },
        { name: "Southern Europe Online", values: [6, 7] },
        { name: "Nordics Direct-to-Consumer", values: [5, 4] },
        { name: "Baltics Partner Channel", values: [3, 6] },
        { name: "Iberia Marketplace", values: [4, 5] },
      ],
    },
    decorations: { seriesLabels: true, categoryAxis: true },
  };
  const nodes = buildChart(cfg).nodes;
  const box = (n: RectNode | TextNode) => ({ x: n.x, y: n.y, r: n.x + n.w, b: n.y + n.h });
  const overlaps = (a: ReturnType<typeof box>, c: ReturnType<typeof box>) =>
    a.x < c.r && c.x < a.r && a.y < c.b && c.y < a.b;
  const legendNodes = nodes.filter(
    (n): n is RectNode | TextNode => (n.kind === "rect" || n.kind === "text") && !!n.name?.startsWith("legend-"),
  );
  const plotRects = nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("seg-"));

  it("actually wraps to more than one row (non-vacuous)", () => {
    const chipRows = new Set(legendNodes.filter((n) => n.name?.startsWith("legend-chip-")).map((n) => n.y));
    expect(chipRows.size).toBeGreaterThan(1);
    expect(plotRects.length).toBeGreaterThan(0);
  });

  it("puts no legend chip or label on top of a bar segment", () => {
    for (const l of legendNodes) {
      for (const p of plotRects) {
        expect(overlaps(box(l), box(p))).toBe(false);
      }
    }
  });

  it("keeps every legend node inside the canvas height", () => {
    for (const l of legendNodes) {
      expect(l.y + l.h).toBeLessThanOrEqual(cfg.height + 1e-6);
    }
  });
});

/**
 * A horizontal line/area chart drew its legend TWICE: `horizontalChrome` emits
 * the shared `legendRow` under `decor.seriesLabels`, and the line layout had a
 * hand-rolled copy of its own under the same condition. The two sat 2.5pt
 * apart, so every series name rendered as a smear — visible in the shipped
 * showcase deck ("Horizontal profile chart (stacked area)"), where "Retail" and
 * "Online" were each drawn over themselves.
 */
describe("a sideways line/area chart legends its series once", () => {
  const cfg = (kind: "line" | "area"): ChartConfig => ({
    kind,
    horizontal: true,
    width: W,
    height: H,
    data: {
      categories: ["North", "South", "East"],
      series: [
        { name: "Retail", values: [10, 20, 15] },
        { name: "Online", values: [5, 8, 12] },
      ],
    },
    decorations: { seriesLabels: true },
  });

  for (const kind of ["line", "area"] as const) {
    it(`emits one chip and one label per series (${kind})`, () => {
      const { nodes } = buildChart(cfg(kind));
      for (const si of [0, 1]) {
        expect(
          nodes.filter((n) => n.name === `legend-chip-${si}`),
          `chips for series ${si}`,
        ).toHaveLength(1);
        expect(
          nodes.filter((n) => n.name === `legend-${si}`),
          `labels for series ${si}`,
        ).toHaveLength(1);
      }
      // Non-vacuous: the legend is actually drawn, not merely not duplicated.
      const labels = nodes.filter((n): n is TextNode => n.kind === "text" && /^legend-\d+$/.test(n.name ?? ""));
      expect(labels.map((n) => n.text)).toEqual(["Retail", "Online"]);
    });
  }
});

/**
 * `seriesLabelNodes` pushes overlapping labels DOWN, then — if the last one
 * overflows the plot — shifts the whole stack up and re-propagates upward, with
 * nothing clamping the top. Twelve series on a 240×160 chart put two labels at
 * negative y; thirty on a 400×300 put the topmost at −123. `collide.ts` refuses
 * to nudge a label off the top for the same reason, but it only moves labels
 * UP, so it cannot rescue one already emitted above the canvas.
 */
describe("right-hand series labels stay on the canvas", () => {
  const sized = (w: number, h: number, count: number): ChartConfig =>
    ({
      kind: "line",
      width: w,
      height: h,
      decorations: { seriesLabels: true },
      data: {
        categories: ["Q1", "Q2"],
        series: Array.from({ length: count }, (_, i) => ({ name: `S${i}`, values: [i + 1, i + 2] })),
      },
    }) as ChartConfig;

  for (const [w, h, count] of [
    [240, 160, 12],
    [400, 300, 30],
    [480, 300, 24],
  ] as const) {
    it(`fits ${count} labels on a ${w}×${h} chart`, () => {
      const labels = buildChart(sized(w, h, count)).nodes.filter(
        (n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("series-label"),
      );
      expect(labels.length, "no labels drawn, so this proves nothing").toBeGreaterThan(count / 2);
      for (const l of labels) {
        expect(l.y, "a label was drawn above the canvas").toBeGreaterThanOrEqual(-0.01);
        expect(l.y + l.h, "a label was drawn below the canvas").toBeLessThanOrEqual(h + 0.01);
      }
    });
  }

  it("leaves a chart with room to spare exactly where it was", () => {
    // The negative control: spreading unconditionally would move every label on
    // every chart, and these three are nowhere near needing it.
    const ys = buildChart(sized(480, 300, 3))
      .nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("series-label"))
      .map((n) => Math.round(n.y * 10) / 10);
    expect(new Set(ys).size).toBe(3);
    expect(Math.min(...ys)).toBeGreaterThan(0);
  });
});
