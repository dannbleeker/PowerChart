import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, RectNode, TextNode, WedgeNode } from "../src/core/scene";

/** Backlog batch F: pie-of-pie breakout, small multiples. */

describe("pie breakout (bar-of-pie)", () => {
  const cfg: ChartConfig = {
    kind: "pie",
    ...DEFAULT_SIZE,
    data: {
      categories: ["EMEA", "Americas", "APAC", "Nordics", "Benelux", "DACH"],
      series: [{ name: "Revenue", values: [80, 100, 60, 20, 15, 25] }],
    },
    pie: { breakout: [3, 4, 5] },
  };
  const s = buildChart(cfg);

  it("collapses breakout categories into one muted Other slice facing the bar", () => {
    expect(s.nodes.some((n) => n.name === "slice-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "slice-3")).toBe(false);
    const other = s.nodes.find((n) => n.name === "slice-other") as WedgeNode;
    expect(other.fill).toBe("#898781");
    // Other (60 of 300 = 72°) is centered at 3 o'clock: spans 90° ± 36°.
    const mid = ((other.startAngle + other.endAngle) / 2) % 360;
    expect(mid).toBeCloseTo(90, 1);
  });

  it("details the breakout in a stacked bar with connectors and grand-total %", () => {
    const segs = [3, 4, 5].map((c) => s.nodes.find((n) => n.name === `breakout-seg-${c}`) as RectNode);
    expect(segs.every(Boolean)).toBe(true);
    // Stacked contiguously, heights ∝ values (20/15/25 of 60).
    expect(segs[0].y + segs[0].h).toBeCloseTo(segs[1].y, 5);
    expect(segs[1].y + segs[1].h).toBeCloseTo(segs[2].y, 5);
    expect(segs[2].h / segs[0].h).toBeCloseTo(25 / 20, 5);
    // Bar sits right of the pie.
    const other = s.nodes.find((n) => n.name === "slice-other") as WedgeNode;
    expect(segs[0].x).toBeGreaterThan(other.cx);
    // Labels carry the share of the GRAND total (20/300 ≈ 7%).
    expect((s.nodes.find((n) => n.name === "breakout-label-3") as TextNode).text).toContain("7%");
    const conns = s.nodes.filter((n): n is LineNode => !!n.name?.startsWith("breakout-conn"));
    expect(conns).toHaveLength(2);
    // Connectors join the bar's top and bottom.
    const ends = conns.map((c) => c.y2).sort((a, b) => a - b);
    expect(ends[0]).toBeCloseTo(segs[0].y, 5);
    expect(ends[1]).toBeCloseTo(segs[2].y + segs[2].h, 5);
  });

  it("plain pies and doughnuts are unaffected", () => {
    const plain = buildChart({ ...cfg, pie: {} });
    expect(plain.nodes.some((n) => n.name === "slice-3")).toBe(true);
    expect(plain.nodes.some((n) => n.name === "slice-other")).toBe(false);
    const dough = buildChart({ ...cfg, kind: "doughnut" });
    expect(dough.nodes.some((n) => n.name === "slice-other")).toBe(false);
  });
});

describe("small multiples", () => {
  const cfg: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    title: "Revenue by region",
    multiples: {},
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [100, 90] },
        { name: "South", values: [40, 100] },
        { name: "East", values: [25, 30] },
      ],
    },
    decorations: { segmentLabels: false },
  };
  const s = buildChart(cfg);

  it("renders one titled panel per series in a grid", () => {
    expect((s.nodes.find((n) => n.name === "p0-title") as TextNode).text).toBe("North");
    expect((s.nodes.find((n) => n.name === "p2-title") as TextNode).text).toBe("East");
    expect((s.nodes.find((n) => n.name === "title") as TextNode).text).toBe("Revenue by region");
    // Three panels side by side (n ≤ 3 → one row).
    const p0 = s.nodes.find((n) => n.name === "p0-seg-0-0") as RectNode;
    const p1 = s.nodes.find((n) => n.name === "p1-seg-0-0") as RectNode;
    expect(p1.x).toBeGreaterThan(p0.x);
  });

  it("panels share one value scale: equal values → equal bar heights", () => {
    // North Q1 = South Q2 = 100 → same height in different panels.
    const north = s.nodes.find((n) => n.name === "p0-seg-0-0") as RectNode;
    const south = s.nodes.find((n) => n.name === "p1-seg-0-1") as RectNode;
    expect(north.h).toBeCloseTo(south.h, 3);
    // East's 30 is visibly smaller than 100, same scale.
    const east = s.nodes.find((n) => n.name === "p2-seg-0-1") as RectNode;
    expect(east.h / north.h).toBeCloseTo(0.3, 1);
  });

  it("columns option and carried special rows work; single series stays whole", () => {
    const stackedCols = buildChart({ ...cfg, multiples: { columns: 1 } });
    const a = stackedCols.nodes.find((n) => n.name === "p0-seg-0-0") as RectNode;
    const b = stackedCols.nodes.find((n) => n.name === "p1-seg-0-0") as RectNode;
    expect(Math.abs(a.x - b.x)).toBeLessThan(30); // same column
    expect(b.y).toBeGreaterThan(a.y);
    const withTarget = buildChart({
      ...cfg,
      data: { ...cfg.data, series: [...cfg.data.series, { name: "Target", values: [110, 110] }] },
    });
    for (const pi of [0, 1, 2]) {
      expect(withTarget.nodes.some((n) => n.name === `p${pi}-target-0`)).toBe(true);
    }
    const single = buildChart({
      ...cfg,
      data: { ...cfg.data, series: [cfg.data.series[0]] },
    });
    expect(single.nodes.some((n) => n.name?.startsWith("p0-"))).toBe(false);
  });

  it("scatter and other row-semantics kinds ignore multiples", () => {
    const scatter = buildChart({
      kind: "scatter",
      ...DEFAULT_SIZE,
      multiples: {},
      data: {
        categories: ["A", "B"],
        series: [
          { name: "X", values: [1, 2] },
          { name: "Y", values: [3, 4] },
        ],
      },
    });
    expect(scatter.nodes.some((n) => n.name?.startsWith("p0-"))).toBe(false);
  });
});
