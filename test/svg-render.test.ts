import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";
import type { SymbolNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * SVG renderer node emission — the paths, polygons and options only this
 * renderer produces. Its accessibility and injection guards live in
 * `a11y-svg.test.ts` and `security-svg.test.ts`.
 */

const cfg = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  data: {
    categories: ["A", "B", "C"],
    series: [
      { name: "S1", values: [10, 20, 30] },
      { name: "S2", values: [5, 5, 5] },
    ],
  },
  ...DEFAULT_SIZE,
  ...partial,
});

describe("SVG renderer options", () => {
  it("paints an explicit background", () => {
    const svg = sceneToSvg(buildChart(cfg({})), { background: "#ffffff" });
    expect(svg).toContain('fill="#ffffff"');
  });

  it("emits annular paths for doughnut wedges", () => {
    const c = cfg({
      kind: "doughnut",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
    });
    const svg = sceneToSvg(buildChart(c));
    // Two arcs per annular wedge: outer radius sweep + inner return.
    expect(svg.match(/<path/g)!.length).toBeGreaterThanOrEqual(2);
  });
});

describe("SVG annular wedge path", () => {
  it("emits inner and outer arcs for wedges with a hole", () => {
    const svg = sceneToSvg({
      width: 100,
      height: 100,
      nodes: [
        {
          kind: "wedge",
          cx: 50,
          cy: 50,
          r: 40,
          innerR: 20,
          startAngle: 0,
          endAngle: 120,
          fill: "#123456",
          stroke: "#000000",
          strokeWidth: 1,
        },
      ],
    });
    expect(svg.match(/A /g)!.length).toBe(2);
    expect(svg).toContain('stroke="#000000"');
  });
});

describe("SVG marker symbols", () => {
  const svgOf = (n: Partial<SymbolNode> = {}) =>
    sceneToSvg({
      width: 100,
      height: 100,
      nodes: [{ kind: "symbol", shape: "diamond", cx: 50, cy: 50, size: 10, fill: "#123456", ...n } as SymbolNode],
    });

  it("draws a filled polygon on the symbol's own points", () => {
    const svg = svgOf();
    expect(svg).toContain('fill="#123456"');
    // The four diamond vertices, at the box edge midpoints.
    expect(svg).toContain('points="50,40 60,50 50,60 40,50"');
  });

  it("carries an optional stroke, and omits it when absent", () => {
    expect(svgOf({ stroke: "#ffffff", strokeWidth: 2 })).toContain('stroke="#ffffff" stroke-width="2"');
    expect(svgOf()).not.toContain("stroke=");
  });

  it("emits the data-name so a symbol is addressable like any other node", () => {
    expect(svgOf({ name: "point-3" })).toContain('data-name="point-3"');
  });

  it("draws every shape, and rounds coordinates like the rest of the renderer", () => {
    for (const shape of ["diamond", "triangle", "plus"] as const) {
      const svg = svgOf({ shape, cx: 33.333333, cy: 12.126, size: 7.77 });
      expect(svg).toContain("<polygon");
      // r() quantises to 2dp; a raw float here would be snapshot noise.
      const pts = svg.match(/points="([^"]+)"/)![1];
      for (const n of pts.split(/[ ,]/)) expect(n).toMatch(/^-?\d+(\.\d{1,2})?$/);
    }
  });
});
