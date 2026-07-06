import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/** Regression tests for the deferred combo / stacked100 / small-multiples fixes. */

const hasNaN = (nodes: { [k: string]: unknown }[]) =>
  nodes.some((n) => Object.values(n).some((v) => typeof v === "number" && Number.isNaN(v)));

describe("combo with a single unmarked series", () => {
  it("renders a plain column, not a column AND a line", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "X", values: [10, 20] }] },
    };
    const scene = buildChart(cfg);
    const lineNodes = scene.nodes.filter((n) => (n.name ?? "").startsWith("combo-line"));
    const markerNodes = scene.nodes.filter((n) => (n.name ?? "").startsWith("combo-marker"));
    const colSegs = scene.nodes.filter((n) => n.kind === "rect" && (n.name ?? "").startsWith("seg"));
    expect(lineNodes.length).toBe(0); // the bug: it was drawn as a line too
    expect(markerNodes.length).toBe(0);
    expect(colSegs.length).toBeGreaterThan(0); // still a column
  });

  it("still draws a line for the last series when there are two", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Col", values: [10, 20] },
          { name: "Ln", values: [5, 8] },
        ],
      },
    };
    const scene = buildChart(cfg);
    expect(scene.nodes.some((n) => (n.name ?? "").startsWith("combo-line"))).toBe(true);
  });
});

describe("stacked100 with an all-negative category", () => {
  it("fills the segments downward instead of collapsing to zero", () => {
    const cfg: ChartConfig = {
      kind: "stacked100",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A"],
        series: [
          { name: "P", values: [-30] },
          { name: "Q", values: [-20] },
        ],
      },
    };
    const scene = buildChart(cfg);
    const segs = scene.nodes.filter(
      (n) => n.kind === "rect" && (n.name ?? "").startsWith("seg") && (n as { h: number }).h > 0.5,
    );
    expect(segs.length).toBe(2); // both shares visible (were 0-height before)
    expect(hasNaN(scene.nodes as never)).toBe(false);
  });
});

describe("combo-waterfall with an overflowing line", () => {
  it("stretches the shared axis so the tall line stays inside the plot", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      ...DEFAULT_SIZE,
      combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "W", values: [10, 10, 10] }, // cumulative peak = 30
          { name: "L", type: "line", values: [5, 60, 5] }, // spikes to 60
        ],
      },
    };
    const scene = buildChart(cfg);
    const markers = scene.nodes.filter((n) => (n.name ?? "").startsWith("combo-marker"));
    expect(markers.length).toBe(3);
    // None of the markers escapes above the top of the canvas.
    expect(Math.min(...markers.map((m) => (m as { y: number }).y))).toBeGreaterThanOrEqual(0);
  });
});

describe("small multiples with a high Target row", () => {
  it("includes the carried Target in the shared panel scale (no clipping)", () => {
    const cfg: ChartConfig = {
      kind: "line",
      ...DEFAULT_SIZE,
      multiples: {},
      decorations: { seriesLabels: true },
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [10, 20] },
          { name: "S2", values: [15, 25] },
          { name: "Target", values: [100, 100] }, // far above the data
        ],
      },
    };
    const scene = buildChart(cfg);
    // Nothing renders above the top of the canvas (the target used to overflow).
    const ys = scene.nodes.map((n) => (n as { y?: number }).y).filter((y): y is number => typeof y === "number");
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(0);
  });
});
