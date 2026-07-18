import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { layoutColumns } from "../src/core/layout/column";
import { layoutMekko } from "../src/core/layout/mekko";
import { layoutButterfly } from "../src/core/layout/butterfly";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";
import { dataToSheet, sheetToData, transposeSheet } from "../src/taskpane/datasheet";

const rects = (nodes: { kind: string }[]) => nodes.filter((n): n is RectNode => n.kind === "rect");
const byName = (nodes: { name?: string }[], prefix: string) => nodes.filter((n) => n.name?.startsWith(prefix));

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

describe("horizontal (bar) orientation", () => {
  const c = cfg({
    horizontal: true,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "S1", values: [10, 20] },
        { name: "S2", values: [5, 15] },
      ],
    },
  });
  const { nodes, anchors } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("bars extend along x and stack contiguously", () => {
    const segs = byName(rects(nodes), "seg-") as RectNode[];
    expect(segs).toHaveLength(4);
    const colA = segs.filter((s) => s.name!.endsWith("-0"));
    const [s1, s2] = colA;
    // Same row (y), s2 starts where s1 ends.
    expect(s1.y).toBeCloseTo(s2.y, 5);
    expect(s2.x).toBeCloseTo(s1.x + s1.w, 5);
  });

  it("bar lengths are proportional to totals", () => {
    const lenA = anchors.columnTop[0] - anchors.baselineY; // x extent, total 15
    const lenB = anchors.columnTop[1] - anchors.baselineY; // total 35
    expect(lenB / lenA).toBeCloseTo(35 / 15, 5);
  });

  it("builds through buildChart without decorations blowing up", () => {
    const scene = buildChart({ ...c, decorations: { cagr: { from: 0, to: 1 }, totals: true } });
    // CAGR must be skipped in horizontal orientation.
    expect(scene.nodes.find((n) => n.name === "cagr-label")).toBeUndefined();
  });
});

describe("butterfly", () => {
  const c = cfg({
    kind: "butterfly",
    data: {
      categories: ["A", "B"],
      series: [
        { name: "L", values: [10, 20] },
        { name: "R", values: [20, 40] },
      ],
    },
  });
  const { nodes } = layoutButterfly(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("left series extends left, right series extends right, shared scale", () => {
    const segs = byName(rects(nodes), "seg-") as RectNode[];
    const l = segs.find((s) => s.name === "seg-0-1")!; // L=20
    const r = segs.find((s) => s.name === "seg-1-1")!; // R=40
    expect(r.w / l.w).toBeCloseTo(2, 3);
    expect(l.x + l.w).toBeLessThanOrEqual(r.x); // separated by the gutter
  });

  it("puts category labels in the center gutter", () => {
    const cats = byName(nodes, "category-") as TextNode[];
    expect(cats).toHaveLength(2);
  });
});

describe("100%= row", () => {
  it("leaves columns short of full height when series sum < denominator", () => {
    const c = cfg({
      kind: "stacked100",
      data: {
        categories: ["A"],
        series: [{ name: "S", values: [50] }],
        hundredPercent: [100],
      },
    });
    const { nodes, anchors } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const seg = byName(rects(nodes), "seg-")[0] as RectNode;
    expect(seg.h).toBeCloseTo(anchors.plot.h / 2, 1);
    const label = byName(nodes, "label-")[0] as TextNode;
    expect(label.text).toBe("50%");
  });
});

describe("mekko with units (X extent)", () => {
  it("takes widths from xExtent and heights from totals", () => {
    const c = cfg({
      kind: "mekko",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [10, 5] },
          { name: "S2", values: [10, 5] },
        ],
        xExtent: [30, 10],
      },
    });
    const { anchors } = layoutMekko(c, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(anchors.categoryWidth[0] / anchors.categoryWidth[1]).toBeCloseTo(3, 3);
    // Column B total (10) is half of A (20) → its top sits lower.
    const hA = anchors.plot.y + anchors.plot.h - anchors.columnTop[0];
    const hB = anchors.plot.y + anchors.plot.h - anchors.columnTop[1];
    expect(hB / hA).toBeCloseTo(0.5, 3);
  });
});

describe("level difference arrow", () => {
  it("compares cumulative stack levels of the given series", () => {
    const c = cfg({
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [100, 150] },
          { name: "S2", values: [50, 50] },
        ],
      },
      decorations: { difference: { from: 0, to: 1, series: 0 } },
    });
    const scene = buildChart(c);
    const label = scene.nodes.find((n) => n.name === "diff-label") as TextNode;
    expect(label.text).toBe("+50%"); // 150 vs 100 at series 0's level
  });
});

describe("multiple value lines", () => {
  it("draws one line per entry", () => {
    const c = cfg({
      data: { categories: ["A", "B"], series: [{ name: "S", values: [50, 100] }] },
      decorations: { valueLines: [{ mode: "mean" }, { mode: "value", value: 90 }] },
    });
    const scene = buildChart(c);
    expect(byName(scene.nodes, "value-line-label")).toHaveLength(2);
    expect(scene.nodes.filter((n) => n.kind === "line" && n.name?.startsWith("value-line-"))).toHaveLength(2);
  });

  it("normalizes the legacy single valueLine", () => {
    const c = cfg({
      data: { categories: ["A"], series: [{ name: "S", values: [50] }] },
      decorations: { valueLine: { mode: "value", value: 25 } },
    });
    const scene = buildChart(c);
    expect(byName(scene.nodes, "value-line-label")).toHaveLength(1);
  });
});

describe("datasheet special rows and transpose", () => {
  it("round-trips 100%= and X extent rows", () => {
    const data = {
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] as (number | null)[] }],
      hundredPercent: [10, 20] as (number | null)[],
      xExtent: [3, 4] as (number | null)[],
    };
    const back = sheetToData(dataToSheet(data));
    expect(back.hundredPercent).toEqual([10, 20]);
    expect(back.xExtent).toEqual([3, 4]);
    expect(back.series).toHaveLength(1);
    expect(back.series[0].values).toEqual([1, 2]);
  });

  it("transposes rows and columns", () => {
    const t = transposeSheet({
      cells: [
        ["", "A", "B"],
        ["S1", "1", "2"],
        ["S2", "3", "4"],
      ],
    });
    expect(t.cells).toEqual([
      ["", "S1", "S2"],
      ["A", "1", "3"],
      ["B", "2", "4"],
    ]);
    const data = sheetToData(t);
    expect(data.categories).toEqual(["S1", "S2"]);
    expect(data.series.map((s) => s.name)).toEqual(["A", "B"]);
  });
});
