import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * Geometry that must stay inside the canvas. These were found by the review as
 * off-frame / wrong-size layout bugs the loose fuzz bound (|c| < 5000) hid.
 */
const W = 480;
const H = 300;
const base = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  ...DEFAULT_SIZE,
  width: W,
  height: H,
  data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 20, 30] }] },
  ...partial,
});

/** Every rect stays within the canvas (a small overhang margin for edge labels). */
const inFrame = (nodes: { kind: string }[], margin = 40) => {
  for (const n of nodes as any[]) {
    if (n.kind !== "rect") continue;
    expect(n.y, `rect top ${n.name}`).toBeGreaterThanOrEqual(-margin);
    expect(n.y + n.h, `rect bottom ${n.name}`).toBeLessThanOrEqual(H + margin);
    expect(n.x, `rect left ${n.name}`).toBeGreaterThanOrEqual(-margin);
    expect(n.x + n.w, `rect right ${n.name}`).toBeLessThanOrEqual(W + margin);
  }
};

describe("decoration bands stay clipped to the plot", () => {
  it("a y-band whose range exceeds the value domain does not render off-frame", () => {
    // valueToY extrapolates past the axis; an unclamped band ran off the canvas.
    const scene = buildChart(base({ decorations: { bands: [{ axis: "y", from: -10_000, to: 10_000 }] } }));
    const band = scene.nodes.find((n) => n.name === "band-0") as any;
    expect(band).toBeTruthy();
    inFrame(scene.nodes);
  });

  it("a band entirely outside the domain is dropped, not drawn with 0/neg height", () => {
    const scene = buildChart(base({ decorations: { bands: [{ axis: "y", from: 1e6, to: 2e6 }] } }));
    // Off-plot band collapses to h<=0 and is filtered out.
    expect(scene.nodes.find((n) => n.name === "band-0")).toBeUndefined();
  });
});

describe("funnel bands fit a short frame", () => {
  it("many stages on a short frame stay inside the canvas", () => {
    const scene = buildChart(
      base({
        kind: "funnel",
        height: 60, // deliberately short: a fixed gap used to overshoot the bottom
        data: { categories: ["a", "b", "c", "d", "e", "f"], series: [{ name: "S", values: [60, 50, 40, 30, 20, 10] }] },
      }),
    );
    const stages = scene.nodes.filter((n) => /^stage-\d+$/.test(n.name ?? "")); // the bands, not stage-value-*
    expect(stages.length).toBe(6);
    for (const s of stages as any[]) expect(s.y + s.h).toBeLessThanOrEqual(60 + 1);
  });
});

describe("variance tier reserves height only where it is drawn", () => {
  it("a line chart does not lose plot height to a variance strip it never draws", () => {
    const off = buildChart(base({ kind: "line", decorations: {} }));
    const on = buildChart(base({ kind: "line", decorations: { variance: { actual: 0, reference: 0 } } }));
    // No variance nodes on a line chart either way…
    expect(on.nodes.some((n) => n.name?.startsWith("variance"))).toBe(false);
    // …and the markers occupy the same vertical extent (no phantom reservation).
    const yOf = (s: any) => s.nodes.filter((n: any) => n.name?.startsWith("marker-")).map((n: any) => n.y);
    expect(Math.max(...yOf(on))).toBeCloseTo(Math.max(...yOf(off)), 5);
  });

  it("a stacked column chart still reserves and draws the variance tier", () => {
    const cfg = base({ kind: "stacked", decorations: { variance: { actual: 0, reference: 0 } } });
    const scene = buildChart(cfg);
    expect(scene.nodes.some((n) => n.name?.startsWith("variance"))).toBe(true);
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
