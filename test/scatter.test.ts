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

/**
 * `scatter.spread` nudges co-located markers apart along one axis. The labels
 * were anchored to the UNDISPLACED position — and the vacated spot is, by
 * construction, exactly where a neighbouring marker moved to, so a label named
 * one point while sitting on another.
 */
describe("point labels follow a spread marker", () => {
  const s = buildChart({
    kind: "scatter",
    ...DEFAULT_SIZE,
    scatter: { spread: "x", spreadLimit: 8 },
    data: {
      // Two points piled up on one coordinate: both markers must be nudged off
      // it. The third only widens the domain the cap is measured against.
      categories: ["P0", "P1", "P2"],
      series: [
        { name: "x", values: [50, 50, 10] },
        { name: "y", values: [20, 20, 60] },
      ],
    },
    decorations: { segmentLabels: true },
  } as unknown as ChartConfig);

  const marker = (i: number) => s.nodes.find((n) => n.name === `point-${i}`) as EllipseNode;
  const label = (text: string) => (s.nodes as TextNode[]).find((n) => n.kind === "text" && n.text === text)!;

  it("spreads the markers apart in the first place", () => {
    expect(marker(0).cx).not.toBeCloseTo(marker(1).cx, 1);
  });

  it("puts each label nearer its own marker than any other", () => {
    for (let i = 0; i < 3; i++) {
      const l = label(`P${i}`);
      const cx = l.x + l.w / 2;
      const cy = l.y + l.h / 2;
      const d = (j: number) => Math.hypot(cx - marker(j).cx, cy - marker(j).cy);
      const nearest = [0, 1, 2].reduce((best, j) => (d(j) < d(best) ? j : best), 0);
      expect(nearest, `label P${i} sits closest to point-${nearest}`).toBe(i);
    }
  });
});

describe("scatter edge cases", () => {
  const scatterData = (over: Partial<ChartConfig["data"]> = {}): ChartConfig["data"] => ({
    categories: ["P1", "P2"],
    series: [
      { name: "X", values: [1, 4] },
      { name: "Y", values: [2, 8] },
    ],
    ...over,
  });
  const c = (partial: Partial<ChartConfig>): ChartConfig =>
    ({ kind: "scatter", ...DEFAULT_SIZE, data: scatterData(), ...partial }) as ChartConfig;

  it("a single point renders without dividing by zero", () => {
    const s = buildChart(
      c({
        data: {
          categories: ["P"],
          series: [
            { name: "X", values: [3] },
            { name: "Y", values: [3] },
          ],
        },
      }),
    );
    expect(s.nodes.some((n) => n.name?.startsWith("point"))).toBe(true);
  });

  it("uses a custom palette for group colors", () => {
    const s = buildChart(
      c({
        style: { palette: ["#111111", "#222222"] },
        data: scatterData({ series: [...scatterData().series, { name: "Group", values: [1, 2] }] }),
      }),
    );
    const chips = s.nodes.filter((n) => n.name?.startsWith("legend-chip"));
    expect(chips.length).toBe(2);
    expect(chips.some((ch) => ch.kind === "rect" && ch.fill === "#111111")).toBe(true);
  });

  it("labelContent controls point labels", () => {
    const s = buildChart(c({ decorations: { segmentLabels: true, labelContent: ["category", "value"] } }));
    const labels = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("label-"));
    expect(labels.map((l) => l.text)).toContain("P1 (1.0, 2.0)");
  });
});

/**
 * Two guards that guarded one direction only.
 */
describe("a scatter stays inside its own plot", () => {
  const cfg = (extra: Partial<ChartConfig>, xs = [1, 2, 3, 4]): ChartConfig =>
    ({
      kind: "scatter",
      ...DEFAULT_SIZE,
      data: {
        categories: ["a", "b", "c", "d"],
        series: [
          { name: "X", values: xs },
          { name: "Y", values: [100, 200, 300, 400] },
        ],
      },
      ...extra,
    }) as ChartConfig;

  it("rides the RIGHT edge with an all-negative x domain, as it rides the left", () => {
    // `toX(0) >= plot.x ? toX(0) : plot.x` pinned the spine when zero fell left
    // of the domain and left it free when zero fell right of it, so a variance
    // or drawdown scatter put a full-height axis 109pt past the canvas.
    const spine = (xs: number[]) => buildChart(cfg({}, xs)).nodes.find((n): n is LineNode => n.name === "y-axis-line")!;
    const neg = spine([-100, -60, -20]);
    const pos = spine([1000, 1025, 1050]);
    expect(neg.x1).toBeLessThanOrEqual(DEFAULT_SIZE.width!);
    expect(pos.x1).toBeGreaterThanOrEqual(0);
    // Non-vacuous: zero inside the domain is still drawn where zero is.
    const inside = spine([-5, 0, 5]);
    expect(inside.x1).toBeGreaterThan(pos.x1);
    expect(inside.x1).toBeLessThan(neg.x1);
  });

  it("clips the OLS trend line to the plot rather than extrapolating off-canvas", () => {
    // The line spans the padded NICE-TICK ends, and `scatterDomain` folds in
    // `X line` rows, x-axis bands and a quadrants crossing — so any of those
    // three stretched x0 far past the data and the fitted y exploded with it
    // (y = 93043 on a 300pt canvas). The polynomial branch already clipped.
    const withTrend = (extra: Partial<ChartConfig>): LineNode | undefined =>
      buildChart({
        ...cfg(extra),
        data: {
          categories: ["a", "b", "c", "d"],
          series: [
            { name: "X", values: [1, 2, 3, 4] },
            { name: "Y", values: [100, 200, 300, 400] },
            { name: "Trend", values: [1, 1, 1, 1] },
          ],
          ...(extra.data ?? {}),
        },
      } as ChartConfig).nodes.find((n): n is LineNode => n.name === "trend");

    const stretchers: [string, Partial<ChartConfig>][] = [
      ["plain", {}],
      ["a quadrants crossing far left", { decorations: { quadrants: { x: -500, y: 200 } } }],
      ["an x band far left", { decorations: { bands: [{ axis: "x", from: -100, to: -90 }] } }],
    ];
    for (const [label, extra] of stretchers) {
      const t = withTrend(extra);
      expect(t, `${label}: no trend line drawn`).toBeTruthy();
      for (const y of [t!.y1, t!.y2]) {
        expect(y, `${label}: trend above the canvas`).toBeGreaterThanOrEqual(0);
        expect(y, `${label}: trend below the canvas`).toBeLessThanOrEqual(DEFAULT_SIZE.height!);
      }
      // Still a real fit, not a flattened one: it rises across the plot.
      expect(Math.abs(t!.y2 - t!.y1), `${label}: trend collapsed`).toBeGreaterThan(10);
    }
  });
});

/**
 * Two ways a fit or a key said nothing, or said something untrue.
 */
describe("a trend the user asked for is drawn, or is not claimed", () => {
  const cloud = (n: number, degree: number): ChartConfig =>
    ({
      kind: "scatter",
      ...DEFAULT_SIZE,
      scatter: { trendDegree: degree },
      data: {
        categories: Array.from({ length: n }, (_, i) => `p${i}`),
        series: [
          { name: "X", values: Array.from({ length: n }, (_, i) => i + 1) },
          { name: "Y", values: Array.from({ length: n }, (_, i) => (i + 1) * (i + 1)) },
          { name: "Trend", values: Array(n).fill(1) },
        ],
      },
    }) as unknown as ChartConfig;
  const trendNodes = (c: ChartConfig) => buildChart(c).nodes.filter((n) => n.name?.startsWith("trend")).length;

  it("still fits three points when a polynomial degree was asked for", () => {
    // `polyTrend` clamps its degree to n − 2 to keep a residual degree of
    // freedom, so three points come back degree 1 — and the caller discarded
    // anything under degree 2. Two points drew a straight fit; three drew
    // nothing at all, with no diagnostic.
    for (const degree of [2, 3, 4]) {
      expect(trendNodes(cloud(3, degree)), `degree ${degree}`).toBeGreaterThan(0);
      expect(trendNodes(cloud(2, degree)), `degree ${degree}, two points`).toBeGreaterThan(0);
    }
  });

  it("still draws the CURVE once there are enough points for one", () => {
    // The negative control: widening the fallback too far would turn every
    // polynomial request into a straight line.
    expect(trendNodes(cloud(6, 2))).toBeGreaterThan(10);
  });
});

describe("a bubble size legend keys something", () => {
  const bubbles = (sizes: number[]): ChartConfig =>
    ({
      kind: "bubble",
      ...DEFAULT_SIZE,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [1, 2, 3] },
          { name: "Y", values: [1, 2, 3] },
          { name: "Size", values: sizes },
        ],
      },
    }) as unknown as ChartConfig;
  const legend = (c: ChartConfig) => buildChart(c).nodes.filter((n) => n.name?.startsWith("size-legend"));

  it("draws nothing when every Size is zero", () => {
    // `maxSize` is floored to an epsilon so the ratios stay finite, so the two
    // reference circles came out DIFFERENT sizes and both were labelled "0.00",
    // over a plot where every bubble sits at the 2.5pt floor.
    expect(legend(bubbles([0, 0, 0]))).toHaveLength(0);
  });

  it("still draws one for a Size row that means something", () => {
    expect(legend(bubbles([5, 10, 20])).length).toBeGreaterThan(0);
  });
});
