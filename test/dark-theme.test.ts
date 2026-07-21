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

/**
 * Label ink and series tints must follow the SURFACE, not assume a light canvas.
 * Two survivors of the #138 sweep: a pie's inside label was hardcoded white (so a
 * pale slice — including the default palette's #eda100 — printed white on light,
 * i.e. invisible), and treemap/violin tinted toward a literal "#ffffff" (which
 * washes out on a dark canvas instead of lifting off it).
 */
describe("label ink and tints follow the surface", () => {
  const lum = (h: string) => parseInt(h.slice(1, 3), 16) + parseInt(h.slice(3, 5), 16) + parseInt(h.slice(5, 7), 16);

  it("inks a pie's inside label against its own slice, not a presumed dark fill", () => {
    const cfg = {
      kind: "pie",
      width: 420,
      height: 320,
      decorations: { segmentLabels: true },
      data: {
        categories: ["Pale", "Dark"],
        series: [{ name: "S", values: [50, 50], colors: ["#f5f2c8", "#123456"] }],
      },
    } as unknown as ChartConfig;
    const nodes = buildChart(cfg).nodes;
    const ink = (i: number) => nodes.find((n) => n.name === `label-${i}`) as { color?: string } | undefined;
    expect(ink(0)?.color).toBe("#0b0b0b"); // dark ink on the pale slice — was #ffffff (invisible)
    expect(ink(1)?.color).toBe("#ffffff"); // white still correct on the dark slice
  });

  it("keeps the default palette's pale slice readable", () => {
    // #eda100 is PALETTE[2], so a 3-slice pie hits this with NO custom colours.
    const cfg = {
      kind: "pie",
      width: 420,
      height: 320,
      decorations: { segmentLabels: true },
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [40, 35, 25] }] },
    } as unknown as ChartConfig;
    const nodes = buildChart(cfg).nodes;
    const pale = nodes.find((n) => n.name === "label-2") as { color?: string } | undefined;
    if (pale) expect(pale.color).toBe("#0b0b0b");
  });

  it("tints treemap groups and violin bodies toward the canvas on a dark theme", () => {
    const base = (kind: string, extra: Record<string, unknown> = {}) =>
      ({
        kind,
        width: 480,
        height: 320,
        data: {
          categories: ["G | one", "G | two", "H | three"],
          series: [{ name: "S", values: [10, 20, 30] }],
        },
        ...extra,
      }) as unknown as ChartConfig;

    for (const kind of ["treemap", "violin"]) {
      const dark = "#1b1b1b";
      const fills = buildChart({ ...base(kind), style: { background: dark } } as ChartConfig)
        .nodes.map((n) => (n as { fill?: string }).fill)
        .filter((f): f is string => typeof f === "string" && /^#[0-9a-f]{6}$/i.test(f));
      // Nothing may end up brighter than a mid grey on a near-black canvas: a tint
      // anchored on a literal white does exactly that.
      const brightest = Math.max(...fills.map(lum));
      expect(brightest, `${kind} on a dark canvas`).toBeLessThan(lum("#c8c8c8"));
    }
  });
});
