import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { EllipseNode, LineNode, RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Scatter / bubble — colour scales, quadrants, trajectories, size legends. */

describe("bubble size legend", () => {
  const bubble: ChartConfig = {
    kind: "bubble",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "X", values: [10, 40, 70] },
        { name: "Y", values: [20, 50, 30] },
        { name: "Size", values: [25, 100, 60] },
      ],
    },
  };

  it("draws two outline reference circles with area-true radii and labels", () => {
    const s = buildChart(bubble);
    const refs = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("size-legend-"));
    expect(refs).toHaveLength(2);
    expect(refs[0].fill).toBe("none");
    // Area ∝ value: full vs half → radius ratio √2.
    expect(refs[0].rx / refs[1].rx).toBeCloseTo(Math.SQRT2, 1);
    expect(s.nodes.filter((n) => n.name?.startsWith("size-legend-label"))).toHaveLength(2);
  });

  it("stays off plain scatter charts", () => {
    expect(buildChart({ ...bubble, kind: "scatter" }).nodes.some((n) => n.name?.startsWith("size-legend"))).toBe(false);
  });
});

describe("scatter quadrants", () => {
  const scatter: ChartConfig = {
    kind: "scatter",
    ...DEFAULT_SIZE,
    data: {
      categories: ["P1", "P2"],
      series: [
        { name: "X", values: [20, 80] },
        { name: "Y", values: [30, 70] },
      ],
    },
    decorations: { quadrants: { x: 50, y: 50, labels: ["Question marks", "Stars", "Dogs", "Cash cows"] } },
  };

  it("shades four zones meeting at the crossing, with labels and lines", () => {
    const s = buildChart(scatter);
    const zones = s.nodes.filter(
      (n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("quadrant-") && !n.name.includes("label"),
    );
    expect(zones).toHaveLength(4);
    // TL and TR share a top edge and meet exactly at the crossing x.
    expect(zones[0].x + zones[0].w).toBeCloseTo(zones[1].x, 5);
    expect(zones[0].y + zones[0].h).toBeCloseTo(zones[2].y, 5);
    const labels = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("quadrant-label"));
    expect(labels.map((l) => l.text)).toContain("Stars");
    expect(s.nodes.some((n) => n.name === "quadrant-x")).toBe(true);
    // Zones render before points (behind them).
    const zoneIdx = s.nodes.findIndex((n) => n.name === "quadrant-0");
    const pointIdx = s.nodes.findIndex((n) => n.name === "point-0");
    expect(zoneIdx).toBeLessThan(pointIdx);
  });
});

describe("scatter/bubble trajectory", () => {
  const base: ChartConfig = {
    kind: "scatter",
    ...DEFAULT_SIZE,
    data: {
      categories: ["t1", "t2", "t3", "t4"],
      series: [
        { name: "X", values: [10, 20, 30, 40] },
        { name: "Y", values: [10, 30, 20, 45] },
      ],
    },
  };

  it("adds no trail by default", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.name?.startsWith("trajectory"))).toBe(false);
  });

  it("connects points in row order with a direction arrowhead per segment", () => {
    const s = buildChart({ ...base, decorations: { trajectory: true } });
    const lines = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("trajectory-"));
    const heads = s.nodes.filter((n) => n.kind === "arrowhead" && n.name?.startsWith("trajectory-head-"));
    expect(lines).toHaveLength(3); // 4 points → 3 segments
    expect(heads).toHaveLength(3);
    // Segment 0 joins the first two points; x increases with X.
    expect(lines[0].x2).toBeGreaterThan(lines[0].x1);
    // Trail sits before the markers in draw order (behind them).
    const firstTraj = s.nodes.findIndex((n) => n.name === "trajectory-0");
    const firstPoint = s.nodes.findIndex((n) => n.name === "point-0");
    expect(firstTraj).toBeLessThan(firstPoint);
  });

  it("works for bubble charts too", () => {
    const s = buildChart({
      kind: "bubble",
      ...DEFAULT_SIZE,
      decorations: { trajectory: true },
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [1, 2, 3] },
          { name: "Y", values: [3, 1, 2] },
          { name: "Size", values: [10, 20, 30] },
        ],
      },
    });
    expect(s.nodes.filter((n) => n.name?.startsWith("trajectory-head-"))).toHaveLength(2);
  });
});

/**
 * Backlog batch I — more §2 within-kind gaps: scatter/bubble continuous
 * color scale, smoothed lines, waterfall grouping spacers.
 */
describe("scatter/bubble continuous color scale", () => {
  const base: ChartConfig = {
    kind: "scatter",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C", "D"],
      series: [
        { name: "X", values: [1, 2, 3, 4] },
        { name: "Y", values: [4, 3, 2, 1] },
        { name: "Color", values: [0, 10, 20, 30] },
      ],
    },
  };

  it("maps points onto a ramp and draws a gradient legend", () => {
    const s = buildChart(base);
    const points = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("point-"));
    const fills = new Set(points.map((p) => p.fill));
    expect(fills.size).toBe(4); // four distinct colors along the ramp
    expect(s.nodes.some((n) => n.name === "color-legend-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "color-legend-min")).toBe(true);
    expect(s.nodes.some((n) => n.name === "color-legend-max")).toBe(true);
  });

  it("supersedes group coloring (no group chips under a color scale)", () => {
    const s = buildChart({
      ...base,
      data: { ...base.data, series: [...base.data.series, { name: "Group", values: [1, 2, 1, 2] }] },
    });
    expect(s.nodes.some((n) => n.name?.startsWith("legend-chip-"))).toBe(false);
    expect(s.nodes.some((n) => n.name === "color-legend-0")).toBe(true);
  });

  it("no color row → no gradient legend (plain scatter)", () => {
    const s = buildChart({
      ...base,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "X", values: [1, 2] },
          { name: "Y", values: [3, 4] },
        ],
      },
    });
    expect(s.nodes.some((n) => n.name?.startsWith("color-legend"))).toBe(false);
  });
});

describe("scatter group colouring honours a short palette", () => {
  it("gives a group beyond the palette length a real colour, not 'undefined'", () => {
    const cfg: ChartConfig = {
      kind: "scatter",
      width: 480,
      height: 320,
      style: { palette: ["#111111", "#222222", "#333333"] } as ChartConfig["style"],
      data: {
        categories: ["a", "b"],
        series: [
          { name: "X", values: [1, 2] },
          { name: "Y", values: [1, 2] },
          { name: "Group", values: [1, 5] }, // group id 5, past the 3-colour palette
        ],
      },
    };
    const fills = buildChart(cfg)
      .nodes.filter((n) => n.name?.startsWith("point-"))
      .map((n) => (n as { fill?: string }).fill);
    expect(fills.length).toBeGreaterThan(0);
    for (const f of fills) {
      expect(f).toBeTruthy();
      expect(f).not.toBe("undefined");
      expect(f).toMatch(/^#[0-9a-fA-F]{6}$/);
    }
  });
});
