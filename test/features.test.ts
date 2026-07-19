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

describe("grand total label (think-cell 14)", () => {
  const base = cfg({
    kind: "stacked",
    data: {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "A", values: [10, 20, 30] },
        { name: "B", values: [5, 10, 15] },
      ],
    },
  });
  const grand = (nodes: { name?: string }[]) => nodes.find((n) => n.name === "grand-total") as TextNode | undefined;

  it("emits one top-right label summing every category total", () => {
    const { nodes } = layoutColumns(base, DEFAULT_STYLE, { ...DEFAULT_DECOR, totals: true, grandTotal: true });
    const g = grand(nodes);
    expect(g).toBeTruthy();
    expect(g!.text).toBe("90"); // 15 + 30 + 45
    // Top-right: right-aligned, near the top of the frame, past the last column's centre.
    expect(g!.align).toBe("right");
    const totals = byName(nodes, "total-") as TextNode[];
    for (const t of totals) expect(g!.y).toBeLessThanOrEqual(t.y + 1); // grand total sits at/above the column totals row
  });

  it("is independent of per-column totals (can show without them)", () => {
    const { nodes } = layoutColumns(
      { ...base, decorations: { ...DEFAULT_DECOR, totals: false, grandTotal: true } },
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, totals: false, grandTotal: true },
    );
    expect(grand(nodes)).toBeTruthy();
    expect(byName(nodes, "total-")).toHaveLength(0); // no per-column totals
  });

  it("clustered sums across every series and category", () => {
    const { nodes } = layoutColumns(
      { ...base, kind: "clustered", decorations: { ...DEFAULT_DECOR, grandTotal: true } },
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, grandTotal: true },
    );
    expect(grand(nodes)!.text).toBe("90");
  });

  it("is suppressed on a 100% chart (every column totals the same) and when horizontal", () => {
    const pct = layoutColumns(
      { ...base, kind: "stacked100", decorations: { ...DEFAULT_DECOR, grandTotal: true } },
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, grandTotal: true },
    );
    expect(grand(pct.nodes)).toBeUndefined();
    const horiz = layoutColumns(
      { ...base, horizontal: true, decorations: { ...DEFAULT_DECOR, grandTotal: true } },
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, grandTotal: true },
    );
    expect(grand(horiz.nodes)).toBeUndefined();
  });
});

describe("IBCS scenario notation (Series.scenario)", () => {
  const seg = (nodes: { name?: string }[], si: number) =>
    nodes.find((n) => n.name === `seg-${si}-0`) as RectNode | undefined;
  const build = (scenarios: (string | undefined)[]) =>
    layoutColumns(
      cfg({
        kind: "clustered",
        data: {
          categories: ["Q1"],
          series: scenarios.map((sc, i) => ({
            name: "Sales",
            color: "#3b6ea5",
            values: [80 + i],
            ...(sc ? { scenario: sc as "AC" | "PY" | "PL" | "BU" | "FC" } : {}),
          })),
        },
      }),
      DEFAULT_STYLE,
      DEFAULT_DECOR,
    ).nodes;

  it("AC is a solid fill of the series colour", () => {
    const r = seg(build(["AC"]), 0)!;
    expect(r.fill).toBe("#3b6ea5");
    expect(r.pattern).toBeUndefined();
  });

  it("PY is a lighter solid (not the base, not white/none)", () => {
    const r = seg(build(["PY"]), 0)!;
    expect(r.fill).not.toBe("#3b6ea5");
    expect(r.fill).not.toBe("none");
    expect(r.fill).toMatch(/^#[0-9a-f]{6}$/i);
  });

  it("PL / BU are outlined/hollow: no fill, coloured border", () => {
    for (const sc of ["PL", "BU"]) {
      const r = seg(build([sc]), 0)!;
      expect(r.fill).toBe("none");
      expect(r.stroke).toBe("#3b6ea5");
      expect(r.strokeWidth).toBeGreaterThan(0);
    }
  });

  it("FC is hatched (diagonal pattern) over the fill", () => {
    const r = seg(build(["FC"]), 0)!;
    expect(r.pattern).toBe("diagonal");
    expect(r.fill).toBe("#3b6ea5");
  });

  it("appends the two-letter scenario code to the legend label", () => {
    const nodes = layoutColumns(
      cfg({
        kind: "clustered",
        data: { categories: ["Q1"], series: [{ name: "Sales", scenario: "AC", values: [80] }] },
      }),
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, seriesLabels: true },
    ).nodes;
    const labels = nodes.filter((n): n is TextNode => n.kind === "text").map((n) => n.text);
    expect(labels.some((t) => t.includes("Sales (AC)"))).toBe(true);
  });
});
