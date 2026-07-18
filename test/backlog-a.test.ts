import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, LineNode, RectNode, TextNode } from "../src/core/scene";

/** Backlog batch A: bullet targets, combo modes, size legend, forecast, quadrants. */

const base: ChartConfig = {
  kind: "clustered",
  ...DEFAULT_SIZE,
  data: {
    categories: ["Q1", "Q2", "Q3"],
    series: [{ name: "Actual", values: [42, 55, 48] }],
  },
};

describe("bullet chart (Target row)", () => {
  const cfg: ChartConfig = {
    ...base,
    data: {
      ...base.data,
      series: [...base.data.series, { name: "Target", values: [50, 50, 60] }],
    },
    decorations: { bands: [{ axis: "y", from: 0, to: 40 }], segmentLabels: true },
  };

  it("draws a bold tick per category at the target, never a bar", () => {
    const s = buildChart(cfg);
    const ticks = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("target-"));
    expect(ticks).toHaveLength(3);
    expect(ticks[0].strokeWidth).toBeGreaterThan(2);
    // Target row does not render as a second series of bars.
    expect(s.nodes.some((n) => n.name === "seg-1-0")).toBe(false);
    // Tick sits above the Q1 bar top (target 50 > actual 42).
    const bar = s.nodes.find((n) => n.name === "seg-0-0") as RectNode;
    expect(ticks[0].y1).toBeLessThan(bar.y);
  });

  it("widens the auto scale to cover targets above the data", () => {
    const withT = buildChart({
      ...cfg,
      data: { ...cfg.data, series: [cfg.data.series[0], { name: "Target", values: [null, null, 90] }] },
    });
    const tick = withT.nodes.find((n) => n.name === "target-2") as LineNode;
    expect(tick.y1).toBeGreaterThan(0); // inside the plot, not clipped
  });
});

describe("combo column modes", () => {
  const combo: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    data: {
      categories: ["2023", "2024"],
      series: [
        { name: "A", values: [30, 40] },
        { name: "B", values: [20, 25] },
        { name: "Margin", values: [31, 38], type: "line" },
      ],
    },
  };

  it("clustered mode puts column series side by side", () => {
    const stacked = buildChart(combo);
    const clustered = buildChart({ ...combo, combo: { columns: "clustered" } });
    const seg = (s: ReturnType<typeof buildChart>, name: string) => s.nodes.find((n) => n.name === name) as RectNode;
    // Stacked: same x for both series. Clustered: different x.
    expect(seg(stacked, "seg-0-0").x).toBeCloseTo(seg(stacked, "seg-1-0").x);
    expect(seg(clustered, "seg-0-0").x).not.toBeCloseTo(seg(clustered, "seg-1-0").x);
    // Line overlay still present in both.
    expect(clustered.nodes.some((n) => n.name?.startsWith("combo-line"))).toBe(true);
  });

  it("stacked100 mode keeps the line on its own right-hand axis", () => {
    const pct = buildChart({ ...combo, combo: { columns: "stacked100" }, decorations: { segmentLabels: true } });
    expect(pct.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
    expect(pct.nodes.some((n) => n.name?.startsWith("combo-line"))).toBe(true);
  });
});

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

describe("forecast styling on lines", () => {
  const line: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["2023", "2024", "2025", "2026"], series: [{ name: "Rev", values: [40, 46, 52, 60] }] },
    decorations: { forecastFrom: 2, segmentLabels: false },
  };

  it("dashes segments and hollows markers from the boundary on", () => {
    const s = buildChart(line);
    const segs = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    expect(segs.find((l) => l.name === "line-0-1")!.dash).toBeUndefined(); // actuals
    expect(segs.find((l) => l.name === "line-0-2")!.dash).toEqual([4, 3]); // into forecast
    expect(segs.find((l) => l.name === "line-0-3")!.dash).toEqual([4, 3]);
    const solid = s.nodes.find((n) => n.name === "marker-0-1") as RectNode;
    const hollow = s.nodes.find((n) => n.name === "marker-0-2") as RectNode;
    expect(solid.fill).not.toBe("#ffffff");
    expect(hollow.fill).toBe("#ffffff");
    expect(s.nodes.some((n) => n.name === "forecast-divider")).toBe(true);
  });

  it("is inert without the option", () => {
    const s = buildChart({ ...line, decorations: { segmentLabels: false } });
    expect(s.nodes.some((n) => n.name === "forecast-divider")).toBe(false);
    expect(
      s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-")).every((l) => !l.dash),
    ).toBe(true);
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
