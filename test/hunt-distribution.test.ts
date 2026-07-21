import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig, Series } from "../src/core/types";
import type { RectNode, SceneNode, TextNode } from "../src/core/scene";

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */

const W = 480;
const H = 300;
const REGIONS = ["Northern Europe", "Southern Europe", "North America", "Asia Pacific", "Latin America", "Middle East"];
const regionSeries = (n: number): Series[] =>
  REGIONS.slice(0, n).map((name, i) => ({ name, values: [1 + i, 2 + i, 3 + i] }));

/** Nodes that carry an x/w box and stick out past the right edge of the canvas. */
const offRight = (nodes: SceneNode[], namePrefix: string, width = W) =>
  nodes.filter((n) => {
    const b = n as Partial<RectNode>;
    return n.name?.startsWith(namePrefix) && b.x != null && b.x + (b.w ?? 0) > width;
  });

const texts = (nodes: SceneNode[], namePrefix: string) =>
  nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith(namePrefix)).map((n) => n.text);

describe("radar legends wrap instead of marching off the canvas", () => {
  it("wraps the polygon radar's legend and reserves the extra row", () => {
    const cfg: ChartConfig = {
      kind: "radar",
      width: W,
      height: H,
      data: { categories: ["Speed", "Cost", "Quality"], series: regionSeries(6) },
    };
    const { nodes } = buildChart(cfg);
    // Before: legend-5 ("Middle East") spanned x 476.4 → 552.6.
    expect(offRight(nodes, "legend")).toHaveLength(0);
    const chipRows = new Set(nodes.filter((n) => n.name?.startsWith("legend-chip-")).map((n) => (n as RectNode).y));
    expect(chipRows.size).toBe(2);
  });

  it("wraps the radial-bar (coxcomb) legend too", () => {
    const cfg: ChartConfig = {
      kind: "radar",
      width: W,
      height: H,
      radar: { bars: true },
      data: { categories: ["Speed", "Cost", "Quality"], series: regionSeries(6) },
    };
    expect(offRight(buildChart(cfg).nodes, "legend")).toHaveLength(0);
  });

  it("keeps the band-mode swatch names while wrapping", () => {
    const cfg: ChartConfig = {
      kind: "radar",
      width: W,
      height: H,
      data: { categories: ["Speed", "Cost", "Quality"], series: regionSeries(4) },
      decorations: { radarBand: true },
    };
    const { nodes } = buildChart(cfg);
    expect(nodes.some((n) => n.name === "legend-band")).toBe(true);
    expect(nodes.some((n) => n.name === "legend-us")).toBe(true);
    expect(offRight(nodes, "legend")).toHaveLength(0);
  });
});

describe("radar perSpoke normalises against each spoke's own maximum", () => {
  it("puts a sub-1 spoke maximum on the rim", () => {
    const cfg: ChartConfig = {
      kind: "radar",
      width: W,
      height: H,
      radar: { perSpoke: true },
      // A rate expressed as 0–1 beside counts — the case perSpoke exists for.
      data: { categories: ["Revenue (m)", "Conversion rate", "NPS"], series: [{ name: "Us", values: [3, 0.5, 40] }] },
    };
    const { nodes } = buildChart(cfg);
    const rim = nodes.filter((n) => n.name?.startsWith("grid-")).at(-1) as { points: { x: number; y: number }[] };
    const shape = nodes.find((n) => n.name === "series-0") as { points: { x: number; y: number }[] };
    // Before: the 0.5 spoke normalised against the Math.max(1, …) floor and
    // drew at half the rim radius.
    shape.points.forEach((p, i) => {
      expect(p.x).toBeCloseTo(rim.points[i].x, 6);
      expect(p.y).toBeCloseTo(rim.points[i].y, 6);
    });
  });
});

describe("butterfly", () => {
  it("wraps the stacked-flank legend inside the canvas", () => {
    const cfg: ChartConfig = {
      kind: "butterfly",
      width: W,
      height: H,
      butterfly: { split: 2 },
      data: { categories: ["A", "B", "C"], series: regionSeries(5) },
    };
    // Before: legend-4 ("Latin America") spanned x 429.0 → 505.2.
    expect(offRight(buildChart(cfg).nodes, "legend")).toHaveLength(0);
  });

  it("floors the plot height instead of emitting negative-height bars", () => {
    const cfg: ChartConfig = {
      kind: "butterfly",
      width: 320,
      height: 48, // shorter than title + header + axis chrome
      title: "Headcount",
      data: {
        categories: ["Sales", "Ops", "Tech"],
        series: [
          { name: "2024", values: [40, 60, 30] },
          { name: "2025", values: [45, 55, 35] },
        ],
      },
      decorations: { segmentLabels: true, categoryAxis: true, valueAxis: true, seriesLabels: true },
    };
    // Before: every seg rect got h = -2.44 — dropped by SVG, clamped to a
    // 0.2pt sliver by Office.js, so preview and deck disagreed.
    const bad = buildChart(cfg).nodes.filter((n) => {
      const b = n as Partial<RectNode>;
      return (b.w != null && b.w < 0) || (b.h != null && b.h < 0);
    });
    expect(bad).toHaveLength(0);
  });
});

describe("decoration anchors never fall back to a value the scale excludes", () => {
  it("candlestick: a blank High anchors on the period's own prices", () => {
    const cfg: ChartConfig = {
      kind: "candlestick",
      width: W,
      height: H,
      data: {
        categories: ["Mon", "Tue", "Wed"],
        series: [
          { name: "Open", values: [100, 102, 103] },
          { name: "High", values: [105, null, 106] },
          { name: "Low", values: [99, 101, 102] },
          { name: "Close", values: [102, 103, 105] },
        ],
      },
      decorations: { callouts: [{ text: "gap", category: 1 }] },
    };
    // Before: toY(0) on a zero-free OHLC scale put the callout box at y ≈ 3452.
    const box = buildChart(cfg).nodes.find((n) => n.name === "callout-box-0") as RectNode;
    expect(box.y).toBeGreaterThanOrEqual(0);
    expect(box.y + box.h).toBeLessThanOrEqual(H);
  });

  it("violin: a category with no observations anchors on the plot floor", () => {
    const cfg: ChartConfig = {
      kind: "violin",
      width: W,
      height: H,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "o1", values: [50, null] },
          { name: "o2", values: [60, null] },
          { name: "o3", values: [70, null] },
        ],
      },
      decorations: { valueAxis: true, callouts: [{ text: "x", category: 1 }] },
    };
    // Before: Math.max(...[0]) on a 50–70 domain put the callout at y ≈ 915.
    const box = buildChart(cfg).nodes.find((n) => n.name === "callout-box-0") as RectNode;
    expect(box.y).toBeGreaterThanOrEqual(0);
    expect(box.y + box.h).toBeLessThanOrEqual(H);
  });
});

describe("funnel conversion marker follows the direction of the step", () => {
  it("marks a rise with ▴ and a fall with ▾", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: W,
      height: H,
      // Ascending — the pyramid ordering funnel.ts recommends.
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 50, 0] }] },
    };
    // Before: "▾ 500.0%" — a down arrow on a 5x increase.
    expect(texts(buildChart(cfg).nodes, "conversion-")).toEqual(["▴ 500.0%", "▾ 0.0%"]);
  });

  it("drops the marker when the stage is unchanged", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: W,
      height: H,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 10] }] },
    };
    expect(texts(buildChart(cfg).nodes, "conversion-")).toEqual(["100.0%"]);
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
