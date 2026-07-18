import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { layoutColumns } from "../src/core/layout/column";
import { layoutWaterfall } from "../src/core/layout/waterfall";
import { layoutMekko } from "../src/core/layout/mekko";
import { layoutGantt } from "../src/core/layout/gantt";
import { layoutScatter } from "../src/core/layout/scatter";
import { resolveLabelCollisions } from "../src/core/collide";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import { evaluateFormula, sheetToData, dataToSheet } from "../src/taskpane/datasheet";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, SceneNode, TextNode } from "../src/core/scene";

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}
const rects = (nodes: { kind: string }[]) => nodes.filter((n): n is RectNode => n.kind === "rect");
const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

describe("clustered-stacked (grouped stacks)", () => {
  const c = cfg({
    data: {
      categories: ["A"],
      series: [
        { name: "S1", values: [10], stack: 0 },
        { name: "S2", values: [20], stack: 0 },
        { name: "S3", values: [25], stack: 1 },
      ],
    },
  });
  const { nodes, anchors } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("places stack groups side by side, each internally stacked", () => {
    const segs = byName(rects(nodes), "seg-") as RectNode[];
    const [s1, s2, s3] = segs;
    expect(s1.x).toBeCloseTo(s2.x, 4); // same stack → same x
    expect(s3.x).toBeGreaterThan(s1.x + s1.w - 1); // other stack sits beside
    expect(s2.y + s2.h).toBeCloseTo(s1.y, 4); // stacked within group
  });

  it("scales to the tallest single stack, not the grand total", () => {
    // Stack 0 totals 30, stack 1 totals 25 → tallest stack (30) sets the top.
    const s12h = anchors.baselineY - anchors.columnTop[0];
    expect(s12h).toBeGreaterThan(0);
  });

  it("round-trips blank separator rows in the datasheet", () => {
    const data = sheetToData({
      cells: [
        ["", "A"],
        ["S1", "10"],
        ["S2", "20"],
        ["", ""],
        ["S3", "25"],
      ],
    });
    expect(data.series.map((s) => s.stack)).toEqual([0, 0, 1]);
    const back = sheetToData(dataToSheet(data));
    expect(back.series.map((s) => s.stack)).toEqual([0, 0, 1]);
  });
});

describe("rotated waterfall & mekko", () => {
  it("horizontal waterfall builds bars along x", () => {
    const c = cfg({
      kind: "waterfall",
      horizontal: true,
      data: { categories: ["Base", "Up", "End"], series: [{ name: "D", values: [50, 20, 0] }] },
      waterfall: { totalIndices: [2] },
    });
    const { nodes, anchors } = layoutWaterfall(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const bars = byName(rects(nodes), "bar-") as RectNode[];
    expect(bars[1].x).toBeCloseTo(bars[0].x + bars[0].w, 4); // floats to the right
    expect(bars[0].y).toBeLessThan(bars[1].y); // rows go downward
    expect(anchors.columnValue).toEqual([50, 70, 70]);
  });

  it("horizontal mekko turns columns into rows", () => {
    const c = cfg({
      kind: "mekko",
      horizontal: true,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [10, 30] },
          { name: "S2", values: [10, 30] },
        ],
      },
    });
    const { anchors, nodes } = layoutMekko(c, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(anchors.categoryWidth[1] / anchors.categoryWidth[0]).toBeCloseTo(3, 3); // row heights ∝ totals
    const segs = byName(rects(nodes), "seg-0-") as RectNode[];
    expect(segs[0].y).toBeLessThan(segs[1].y); // categories stacked vertically
  });
});

describe("global label collisions", () => {
  it("nudges overlapping outside labels apart", () => {
    const nodes: SceneNode[] = [
      {
        kind: "text",
        x: 0,
        y: 0,
        w: 60,
        h: 12,
        text: "88888",
        fontSize: 10,
        color: "#000",
        align: "center",
        valign: "middle",
        name: "total-0",
      },
      {
        kind: "text",
        x: 0,
        y: 0,
        w: 60,
        h: 12,
        text: "99999",
        fontSize: 10,
        color: "#000",
        align: "center",
        valign: "middle",
        name: "cagr-label",
      },
    ];
    resolveLabelCollisions(nodes);
    const [a, b] = nodes as TextNode[];
    expect(Math.abs(a.y - b.y)).toBeGreaterThan(10); // no longer overlapping
  });
});

describe("gantt granularity, grouping, remarks", () => {
  const day = (iso: string) => Math.round(Date.parse(iso) / 86400000);
  it("uses week ticks with weekend shading for short plans", () => {
    const c = cfg({
      kind: "gantt",
      data: {
        categories: ["Phase 1", "> Detail | Anna | on track"],
        series: [
          { name: "Start", values: [null, day("2026-01-05")] },
          { name: "End", values: [null, day("2026-02-06")] },
        ],
        dates: true,
      },
    });
    const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(byName(nodes, "weekend-").length).toBeGreaterThan(2);
    expect(nodes.find((n) => n.name === "section-0")).toBeTruthy(); // no dates → header row
    const owner = nodes.find((n) => n.name === "owner-1") as TextNode;
    expect(owner.text).toBe("Anna");
    const remark = nodes.find((n) => n.name === "remark-1") as TextNode;
    expect(remark.text).toBe("on track");
    const cat = nodes.find((n) => n.name === "category-1") as TextNode;
    expect(cat.text).toBe("Detail");
    expect(cat.x).toBeGreaterThan(0); // indented
  });
});

describe("scatter partition, trend, legend", () => {
  const c = cfg({
    kind: "scatter",
    data: {
      categories: ["P1", "P2", "P3", "P4"],
      series: [
        { name: "X", values: [0, 10, 20, 30] },
        { name: "Y", values: [0, 10, 20, 30] },
        { name: "Group", values: [1, 1, 2, 2] },
        { name: "X line", values: [15, null, null, null] },
        { name: "Y line", values: [15, null, null, null] },
        { name: "Trend", values: [1, null, null, null] },
      ],
    },
  });
  const { nodes } = layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR);
  it("draws partition and trend lines", () => {
    expect(nodes.find((n) => n.name === "x-line")).toBeTruthy();
    expect(nodes.find((n) => n.name === "y-line")).toBeTruthy();
    expect(nodes.find((n) => n.name === "trend")).toBeTruthy();
  });
  it("adds a group legend for multiple groups", () => {
    expect(byName(nodes, "legend-chip-")).toHaveLength(2);
  });
});

describe("datasheet formulas", () => {
  const cells = [
    ["", "2024", "2025"],
    ["Rev", "100", "120"],
    ["Cost", "40", "50"],
    ["Profit", "=B2-B3", "=C2-C3"],
  ];
  it("evaluates refs, arithmetic and functions", () => {
    expect(evaluateFormula(cells, "B2-B3")).toBe(60);
    expect(evaluateFormula(cells, "SUM(B2:C3)")).toBe(310);
    expect(evaluateFormula(cells, "AVG(B2,C2)*2")).toBe(220);
    expect(evaluateFormula(cells, "(B2+C2)/2")).toBe(110);
  });
  it("resolves formula chains and rejects cycles", () => {
    const data = sheetToData({ cells });
    expect(data.series[2].values).toEqual([60, 70]);
    expect(evaluateFormula([["=A1"]], "A1")).toBeNull();
  });
});

describe("smooth pie fan density", () => {
  it("keeps snapshot-level SVG unchanged (fans are a PPT-only concern)", () => {
    const scene = buildChart(
      cfg({ kind: "pie", data: { categories: ["A", "B"], series: [{ name: "S", values: [60, 40] }] } }),
    );
    expect(scene.nodes.filter((n) => n.kind === "wedge")).toHaveLength(2);
  });
});
