import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { WedgeNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Radial bar (coxcomb) and variable-radius pie. */

/** Batch S — polar family: radial bars, stacked radar, variable-radius pie. */
const dist = (w: WedgeNode) => w.r - w.innerR;

describe("radial bar chart (coxcomb)", () => {
  const cfg: ChartConfig = {
    kind: "radar",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "V", values: [10, 40, 20] }] },
    radar: { bars: true },
  };
  const s = buildChart(cfg);
  const bar = (c: number) => s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === `bar-${c}`)!;

  it("draws one radius-encoded wedge per category (no connecting polygon)", () => {
    expect(bar(0)).toBeTruthy();
    expect(bar(1)).toBeTruthy();
    expect(bar(2)).toBeTruthy();
    // Bar length encodes value: B (40) > C (20) > A (10).
    expect(dist(bar(1))).toBeGreaterThan(dist(bar(2)));
    expect(dist(bar(2))).toBeGreaterThan(dist(bar(0)));
    // It is not the polygon radar — no series polygon emitted.
    expect(s.nodes.some((n) => n.kind === "polygon" && n.name === "series-0")).toBe(false);
  });

  it("stacks multiple series outward within each sector", () => {
    const multi = buildChart({
      ...cfg,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "X", values: [10, 5] },
          { name: "Y", values: [8, 12] },
        ],
      },
    });
    const x0 = multi.nodes.find((n): n is WedgeNode => n.name === "bar-0-0")!;
    const y0 = multi.nodes.find((n): n is WedgeNode => n.name === "bar-0-1")!;
    expect(x0).toBeTruthy();
    expect(y0).toBeTruthy();
    // Second series sits on top of the first: its inner radius ≈ first's outer.
    expect(y0.innerR).toBeGreaterThanOrEqual(x0.r - 0.01);
  });
});
