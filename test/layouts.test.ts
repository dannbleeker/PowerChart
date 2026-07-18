import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { layoutColumns } from "../src/core/layout/column";
import { layoutWaterfall } from "../src/core/layout/waterfall";
import { layoutMekko } from "../src/core/layout/mekko";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";

const rects = (nodes: { kind: string }[]) => nodes.filter((n): n is RectNode => n.kind === "rect");
const byName = (nodes: { name?: string }[], prefix: string) => nodes.filter((n) => n.name?.startsWith(prefix));

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

describe("stacked layout", () => {
  const c = cfg({
    data: {
      categories: ["A", "B"],
      series: [
        { name: "S1", values: [10, 20] },
        { name: "S2", values: [5, 15] },
      ],
    },
  });
  const { nodes, anchors } = layoutColumns(c, DEFAULT_STYLE, { ...DEFAULT_DECOR, totals: true });

  it("stacks segments contiguously from the baseline", () => {
    const segs = byName(rects(nodes), "seg-") as RectNode[];
    expect(segs).toHaveLength(4);
    const colA = segs.filter((s) => s.name!.endsWith("-0"));
    // Segment tops/bottoms meet: bottom of upper segment == top of lower segment.
    const [s1, s2] = colA;
    expect(s2.y + s2.h).toBeCloseTo(s1.y, 5);
    expect(s1.y + s1.h).toBeCloseTo(anchors.baselineY, 5);
  });

  it("column heights are proportional to totals", () => {
    const hA = anchors.baselineY - anchors.columnTop[0]; // total 15
    const hB = anchors.baselineY - anchors.columnTop[1]; // total 35
    expect(hB / hA).toBeCloseTo(35 / 15, 5);
  });

  it("emits totals above columns", () => {
    const totals = byName(nodes, "total-") as TextNode[];
    expect(totals.map((t) => t.text)).toEqual(["15", "35"]);
  });
});

describe("negative values", () => {
  it("draws negatives below the baseline", () => {
    const c = cfg({
      data: { categories: ["A"], series: [{ name: "S", values: [-10] }] },
    });
    const { nodes, anchors } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const seg = byName(rects(nodes), "seg-")[0] as RectNode;
    expect(seg.y).toBeCloseTo(anchors.baselineY, 5);
    expect(seg.h).toBeGreaterThan(0);
  });
});

describe("stacked100 layout", () => {
  it("normalizes every column to full plot height", () => {
    const c = cfg({
      kind: "stacked100",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [30, 300] },
          { name: "S2", values: [70, 700] },
        ],
      },
    });
    const { nodes, anchors } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const segs = byName(rects(nodes), "seg-") as RectNode[];
    for (const colIdx of [0, 1]) {
      const col = segs.filter((s) => s.name!.endsWith(`-${colIdx}`));
      const total = col.reduce((a, s) => a + s.h, 0);
      expect(total).toBeCloseTo(anchors.plot.h, 1);
    }
  });
});

describe("waterfall layout", () => {
  const c = cfg({
    kind: "waterfall",
    data: {
      categories: ["Start", "Up", "Down", "End"],
      series: [{ name: "D", values: [100, 30, -20, 0] }],
    },
    waterfall: { totalIndices: [3] },
  });
  const { nodes, anchors } = layoutWaterfall(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("computes the running total for 'e' columns", () => {
    expect(anchors.columnValue[3]).toBe(110); // 100 + 30 - 20
  });

  it("floats delta bars at the running level", () => {
    const bars = byName(rects(nodes), "bar-") as RectNode[];
    const [start, up, , end] = bars;
    // "Up" starts where "Start" ends.
    expect(up.y + up.h).toBeCloseTo(start.y, 5);
    // Total bar reaches from baseline to running total.
    expect(end.y + end.h).toBeCloseTo(anchors.baselineY, 5);
  });

  it("draws connectors between consecutive bars", () => {
    expect(byName(nodes, "connector-")).toHaveLength(3);
  });

  it("colors decreases with the negative fill", () => {
    const bars = byName(rects(nodes), "bar-") as RectNode[];
    expect(bars[2].fill).toBe(DEFAULT_STYLE.negative);
    expect(bars[3].fill).toBe(DEFAULT_STYLE.neutral);
  });
});

describe("mekko layout", () => {
  it("makes column widths proportional to totals", () => {
    const c = cfg({
      kind: "mekko",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [10, 30] },
          { name: "S2", values: [10, 30] },
        ],
      },
    });
    const { anchors } = layoutMekko(c, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(anchors.categoryWidth[1] / anchors.categoryWidth[0]).toBeCloseTo(3, 3);
  });
});

describe("decorations", () => {
  it("adds a CAGR arrow with the computed rate", () => {
    const c = cfg({
      data: { categories: ["Y1", "Y2", "Y3"], series: [{ name: "S", values: [100, 110, 121] }] },
      decorations: { cagr: { from: 0, to: 2 } },
    });
    const scene = buildChart(c);
    const label = scene.nodes.find((n) => n.name === "cagr-label") as TextNode;
    expect(label.text).toContain("+10.0%");
  });

  it("adds a difference arrow with percent delta", () => {
    const c = cfg({
      data: { categories: ["A", "B"], series: [{ name: "S", values: [100, 150] }] },
      decorations: { difference: { from: 0, to: 1 } },
    });
    const scene = buildChart(c);
    const label = scene.nodes.find((n) => n.name === "diff-label") as TextNode;
    expect(label.text).toBe("+50%");
  });
});

describe("samples build cleanly", () => {
  for (const kind of ["stacked", "clustered", "stacked100", "waterfall", "mekko", "line", "area"] as const) {
    it(`builds ${kind}`, () => {
      const scene = buildChart(sampleConfig(kind));
      expect(scene.nodes.length).toBeGreaterThan(5);
      // Everything stays inside the frame horizontally (with label slack).
      for (const n of scene.nodes) {
        if (n.kind === "rect") {
          expect(n.x).toBeGreaterThanOrEqual(-1);
          expect(n.x + n.w).toBeLessThanOrEqual(scene.width + 1);
        }
      }
    });
  }
});
