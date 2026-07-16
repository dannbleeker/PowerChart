import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { layoutScatter } from "../src/core/layout/scatter";
import { layoutGantt } from "../src/core/layout/gantt";
import { layoutColumns } from "../src/core/layout/column";
import { placeLabels } from "../src/core/labels";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, RectNode, TextNode } from "../src/core/scene";

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

describe("label placer", () => {
  const bounds = { x: 0, y: 0, w: 200, h: 200 };
  it("places non-conflicting labels and skips impossible ones", () => {
    const reqs = [
      { cx: 100, cy: 100, r: 3, w: 30, h: 10 },
      { cx: 100, cy: 100, r: 3, w: 30, h: 10 }, // same anchor → different slot
      { cx: 100, cy: 100, r: 3, w: 500, h: 10 }, // wider than bounds → hidden
    ];
    const placed = placeLabels(reqs, bounds);
    expect(placed).toHaveLength(2);
    expect(placed[0].slot).not.toBe(placed[1].slot);
  });
  it("respects obstacles", () => {
    const placed = placeLabels(
      [{ cx: 10, cy: 10, r: 2, w: 20, h: 8 }],
      bounds,
      [{ x: 0, y: 0, w: 200, h: 200 }], // everything blocked
    );
    expect(placed).toHaveLength(0);
  });
});

describe("scatter & bubble", () => {
  it("maps X/Y rows to point positions", () => {
    const c = cfg({
      kind: "scatter",
      data: {
        categories: ["P1", "P2"],
        series: [
          { name: "X", values: [0, 100] },
          { name: "Y", values: [0, 100] },
        ],
      },
    });
    const { nodes, anchors } = layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const pts = nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("point-"));
    expect(pts).toHaveLength(2);
    // P2 is right of and above P1.
    expect(pts[1].cx).toBeGreaterThan(pts[0].cx);
    expect(pts[1].cy).toBeLessThan(pts[0].cy);
    expect(anchors.plot.w).toBeGreaterThan(0);
  });

  it("scales bubble area by the Size row", () => {
    const c = cfg({
      kind: "bubble",
      data: {
        categories: ["Small", "Big"],
        series: [
          { name: "X", values: [10, 90] },
          { name: "Y", values: [10, 90] },
          { name: "Size", values: [25, 100] },
        ],
      },
    });
    const { nodes } = layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR);
    // By name, not by position: markers paint back-to-front (largest first), so
    // the emission order is deliberately not the datasheet order.
    const pt = (i: number) =>
      nodes.find((n): n is EllipseNode => n.kind === "ellipse" && n.name === `point-${i}`)!;
    // Area ∝ size → radius ratio = sqrt(100/25) = 2.
    expect(pt(1).rx / pt(0).rx).toBeCloseTo(2, 1);
  });

  it("paints bubbles back-to-front so a big one cannot bury a small one", () => {
    const c = cfg({
      kind: "bubble",
      data: {
        categories: ["Big", "Small"],
        series: [
          // Same point: the small bubble sits inside the big one's disc.
          { name: "X", values: [50, 50] },
          { name: "Y", values: [50, 50] },
          { name: "Size", values: [100, 4] },
        ],
      },
    });
    const { nodes } = layoutScatter(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const order = nodes
      .filter((n) => n.kind === "ellipse" && n.name?.startsWith("point-"))
      .map((n) => (n as EllipseNode).name);
    // Datasheet order would emit the big one first and paint the small one
    // under it — invisible in all three renderers.
    expect(order).toEqual(["point-0", "point-1"]);
    const rx = (i: number) =>
      (nodes.find((n) => n.name === `point-${i}`) as EllipseNode).rx;
    expect(rx(0)).toBeGreaterThan(rx(1)); // point-0 IS the big one, drawn first
  });

  it("labels points without overlaps", () => {
    const scene = buildChart(sampleConfig("bubble"));
    const labels = scene.nodes.filter(
      (n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("label-"),
    );
    expect(labels.length).toBeGreaterThan(3);
    for (let i = 0; i < labels.length; i++) {
      for (let j = i + 1; j < labels.length; j++) {
        const a = labels[i];
        const b = labels[j];
        const overlap = a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
        expect(overlap).toBe(false);
      }
    }
  });
});

describe("gantt", () => {
  it("draws activity bars spanning start to end", () => {
    const c = cfg({
      kind: "gantt",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Start", values: [0, 5] },
          { name: "End", values: [5, 10] },
          { name: "Milestone", values: [null, 10] },
        ],
      },
    });
    const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const bars = nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("bar-"));
    expect(bars).toHaveLength(2);
    // Equal durations → equal widths; B starts where A ends.
    expect(bars[0].w).toBeCloseTo(bars[1].w, 1);
    expect(bars[1].x).toBeCloseTo(bars[0].x + bars[0].w, 1);
    expect(nodes.find((n) => n.name === "milestone-1")).toBeTruthy();
  });

  describe("gutter columns (\"Column <label>\" rows)", () => {
    const withColumns = (extra: Record<string, unknown> = {}) =>
      cfg({
        kind: "gantt",
        width: 640,
        height: 260,
        data: {
          categories: ["Phase 1", "> Design", "> Build"],
          series: [
            { name: "Start", values: [null, 0, 4] },
            { name: "End", values: [null, 4, 10] },
            { name: "Column Cost", values: [250, 90.5, 159.5] },
            { name: "Column FTE", values: [3, 1, 2] },
          ],
        },
        ...extra,
      });

    it("renders each Column row as a headed, right-aligned gutter column", () => {
      const { nodes } = layoutGantt(withColumns(), DEFAULT_STYLE, DEFAULT_DECOR);
      const text = (name: string) => nodes.find((n) => n.name === name) as TextNode;
      expect(text("col-head-0").text).toBe("Cost");
      expect(text("col-head-1").text).toBe("FTE");
      expect(text("col-0-1").text).toBe("91"); // 90.5, at the Cost column's own precision
      expect(text("col-1-1").text).toBe("1.0"); // FTE resolves its own decimals
      for (const n of nodes.filter((x) => x.name?.startsWith("col-"))) {
        expect((n as TextNode).align).toBe("right");
      }
      // Columns sit side by side, left of the plot.
      expect(text("col-head-1").x).toBeGreaterThan(text("col-head-0").x);
    });

    it("does not turn a Column row into a bar, and shrinks the plot to make room", () => {
      const bars = (c: ChartConfig) =>
        layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR).nodes.filter(
          (n): n is RectNode => n.kind === "rect" && /^bar-\d+$/.test(n.name ?? ""),
        );
      const withCols = bars(withColumns());
      expect(withCols).toHaveLength(2); // Design + Build; Phase 1 has no span

      // Same chart without the Column rows: the bars must start further left.
      const bare = withColumns();
      bare.data.series = bare.data.series.filter((s) => !s.name.startsWith("Column"));
      const without = bars(bare);
      expect(withCols[0].x).toBeGreaterThan(without[0].x);
    });

    it("paints cells after the section band, which is full-width and would cover them", () => {
      // Phase 1 has no Start/End, so it is a section header: `section-0` is a
      // rect spanning the whole chart width, including the gutter.
      const { nodes } = layoutGantt(withColumns(), DEFAULT_STYLE, DEFAULT_DECOR);
      const band = nodes.findIndex((n) => n.name === "section-0");
      const cell = nodes.findIndex((n) => n.name === "col-0-0");
      expect(band).toBeGreaterThanOrEqual(0);
      expect(cell).toBeGreaterThan(band);
      // A header row still shows the value it carries — nothing is auto-summed.
      expect((nodes[cell] as TextNode).text).toBe("250");
    });
  });
});

describe("segment order & manual scale", () => {
  const data = {
    categories: ["A"],
    series: [
      { name: "S1", values: [10] },
      { name: "S2", values: [30] },
      { name: "S3", values: [20] },
    ],
  };
  it("descending puts the largest segment at the baseline", () => {
    const { nodes, anchors } = layoutColumns(
      cfg({ data, segmentOrder: "descending" }),
      DEFAULT_STYLE,
      DEFAULT_DECOR,
    );
    const segs = nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("seg-"));
    const bottom = segs.reduce((a, b) => (a.y + a.h > b.y + b.h ? a : b));
    expect(bottom.name).toBe("seg-1-0"); // S2 = 30, the largest
    expect(bottom.y + bottom.h).toBeCloseTo(anchors.baselineY, 5);
  });
  it("pins the axis max", () => {
    const { anchors } = layoutColumns(
      cfg({ data, scale: { max: 120 } }),
      DEFAULT_STYLE,
      DEFAULT_DECOR,
    );
    // Total 60 of max 120 → column fills half the plot.
    expect((anchors.baselineY - anchors.columnTop[0]) / anchors.plot.h).toBeCloseTo(0.5, 2);
  });
});
