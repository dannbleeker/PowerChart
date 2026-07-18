import { describe, expect, it } from "vitest";
import { zoneFill } from "../src/core/color";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode } from "../src/core/scene";

/**
 * Subtle background zones (scatter quadrants, gantt weekend/holiday shading,
 * heatmap total strips) hardcode light-canvas tints. zoneFill keeps them exact
 * on a light background (so default charts are byte-identical) but adapts them on
 * a dark background, where a near-white box would glare.
 */
describe("zoneFill", () => {
  it("returns the light tint unchanged on a light background (byte-identical)", () => {
    for (const c of ["#efe7e7", "#f4f3f0", "#f0efec", "#f2f1ec", "#faf9f6"]) {
      expect(zoneFill("#ffffff", c)).toBe(c);
    }
  });

  it("adapts a light tint to a subtle panel on a dark background", () => {
    const dark = "#1a1a1a";
    const out = zoneFill(dark, "#f2f1ec");
    expect(out).not.toBe("#f2f1ec"); // not the glaring light box
    // A faint lift off the dark canvas: darker than the light literal, lighter
    // than the background itself.
    const lum = (h: string) => parseInt(h.slice(1, 3), 16) + parseInt(h.slice(3, 5), 16) + parseInt(h.slice(5, 7), 16);
    expect(lum(out)).toBeLessThan(lum("#f2f1ec"));
    expect(lum(out)).toBeGreaterThan(lum(dark));
  });

  it("preserves the two-tone quadrant checkerboard on both themes", () => {
    const cfg = (bg?: string): ChartConfig => ({
      kind: "scatter",
      width: 480,
      height: 360,
      style: bg ? ({ background: bg } as ChartConfig["style"]) : undefined,
      decorations: { quadrants: { x: 5, y: 5 } },
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [2, 8, 5] },
          { name: "Y", values: [3, 7, 6] },
        ],
      },
    });
    const quads = (c: ChartConfig) =>
      buildChart(c)
        .nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("quadrant-"))
        .map((n) => n.fill);
    const light = quads(cfg());
    const dark = quads(cfg("#1a1a1a"));
    // Light: the tuned literals, unchanged.
    expect(light).toEqual(["#f2f1ec", "#faf9f6", "#faf9f6", "#f2f1ec"]);
    // Dark: adapted, and still alternating (0 and 3 share a tone, 1 and 2 share the other).
    expect(dark[0]).not.toBe("#f2f1ec");
    expect(dark[0]).toBe(dark[3]);
    expect(dark[1]).toBe(dark[2]);
    expect(dark[0]).not.toBe(dark[1]);
  });
});
