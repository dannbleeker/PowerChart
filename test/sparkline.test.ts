import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { EllipseNode, LineNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Sparklines. */

const base = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  ...DEFAULT_SIZE,
  width: W,
  height: H,
  data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 20, 30] }] },
  ...partial,
});

/**
 * Geometry that must stay inside the canvas. These were found by the review as
 * off-frame / wrong-size layout bugs the loose fuzz bound (|c| < 5000) hid.
 */
const W = 480;

const H = 300;

describe("sparklines", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: {
      categories: ["1", "2", "3", "4", "5"],
      series: [{ name: "Visits", values: [10, 40, 20, 5, 25] }],
    },
    decorations: { sparkline: true },
  };
  const s = buildChart(cfg);

  it("draws a thin line with min/max/last dots and no axis chrome", () => {
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && n.strokeWidth === 1.25)).toBe(true);
    const min = s.nodes.find((n): n is EllipseNode => n.name === "spark-min-0")!;
    const max = s.nodes.find((n): n is EllipseNode => n.name === "spark-max-0")!;
    const last = s.nodes.find((n): n is EllipseNode => n.name === "spark-last-0")!;
    expect(min.fill).toBe("#e34948"); // red on the minimum (value 5)
    expect(max.fill).toBe("#1a9e6e"); // green on the maximum (value 40)
    // Max (40) sits above min (5) on screen.
    expect(max.cy).toBeLessThan(min.cy);
    // Last dot is at the final category, to the right of the (earlier) max.
    expect(last.cx).toBeGreaterThan(max.cx);
    // No gridlines/axis lines emitted.
    expect(s.nodes.some((n) => n.name === "gridline")).toBe(false);
  });

  it("area sparklines add a light fill beneath the line", () => {
    const areaSpark = buildChart({ ...cfg, kind: "area" });
    expect(areaSpark.nodes.some((n) => n.name?.startsWith("spark-fill-"))).toBe(true);
  });
});

describe("sparkline area fill stays within its shape budget", () => {
  it("a wide sparkline segment emits no more slabs than the pre-#128 cap of 8", () => {
    const scene = buildChart(
      base({
        kind: "area",
        width: 2000, // very wide: slabSteps' 24-cap would triple the count
        height: 40, // sparkline territory
        decorations: { sparkline: true },
        data: { categories: ["a", "b"], series: [{ name: "S", values: [1, 5] }] },
      }),
    );
    const slabs = scene.nodes.filter((n) => n.name?.startsWith("spark-fill-0-0-"));
    expect(slabs.length).toBeGreaterThan(0);
    expect(slabs.length).toBeLessThanOrEqual(8);
  });
});
