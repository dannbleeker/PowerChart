import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { LineNode, RectNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Bullet chart (target row). */

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
