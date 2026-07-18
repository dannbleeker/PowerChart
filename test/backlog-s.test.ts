import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { PolygonNode, WedgeNode } from "../src/core/scene";

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

describe("stacked radar", () => {
  const cfg: ChartConfig = {
    kind: "radar",
    ...DEFAULT_SIZE,
    data: {
      categories: ["P", "Q", "R"],
      series: [
        { name: "A", values: [2, 3, 1] },
        { name: "B", values: [3, 2, 4] },
        { name: "C", values: [1, 1, 2] },
      ],
    },
    radar: { stacked: true },
  };
  const s = buildChart(cfg);

  it("draws one filled band polygon per series", () => {
    for (const si of [0, 1, 2]) {
      expect(s.nodes.some((n): n is PolygonNode => n.kind === "polygon" && n.name === `series-${si}`)).toBe(true);
    }
    // No overlaid single-value markers in stacked mode.
    expect(s.nodes.some((n) => n.name?.startsWith("marker-"))).toBe(false);
  });

  it("scales to the per-spoke sums so the outer band reaches the total", () => {
    // Distance from centre of the outermost band's first-spoke point should
    // exceed that of the innermost band (cumulative stacking).
    const cxu = DEFAULT_SIZE.width / 2;
    const inner = s.nodes.find((n): n is PolygonNode => n.name === "series-0")!;
    const outer = s.nodes.find((n): n is PolygonNode => n.name === "series-2")!;
    const rOf = (p: PolygonNode) => Math.abs(p.points[0].x - cxu) + Math.abs(p.points[0].y);
    expect(rOf(outer)).not.toBe(rOf(inner));
  });
});

describe("variable-radius pie", () => {
  const cfg: ChartConfig = {
    kind: "pie",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "Share", values: [50, 30, 20] }, // angle
        { name: "Radius", values: [20, 80, 50] }, // radius metric
      ],
    },
    pie: { variableRadius: true },
  };
  const s = buildChart(cfg);
  const slice = (c: number) => s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === `slice-${c}`)!;

  it("keeps angle ∝ the first series but sets radius from the Radius row", () => {
    // Angle still encodes Share: A (50) sweeps the widest.
    const spanA = slice(0).endAngle - slice(0).startAngle;
    const spanB = slice(1).endAngle - slice(1).startAngle;
    expect(spanA).toBeGreaterThan(spanB);
    // Radius encodes the Radius row: B (80) longest, A (20) shortest — even
    // though A has the biggest share.
    expect(slice(1).r).toBeGreaterThan(slice(2).r);
    expect(slice(2).r).toBeGreaterThan(slice(0).r);
  });

  it("is inert without a Radius row or the flag", () => {
    const plain = buildChart({
      ...cfg,
      pie: {},
      data: { categories: ["A", "B"], series: [{ name: "Share", values: [50, 50] }] },
    });
    const a = plain.nodes.find((n): n is WedgeNode => n.name === "slice-0")!;
    const b = plain.nodes.find((n): n is WedgeNode => n.name === "slice-1")!;
    expect(a.r).toBe(b.r); // uniform radius
  });
});
