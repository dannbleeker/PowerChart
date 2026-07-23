import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { WedgeNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Sunburst rings. */

/**
 * Batch Q — three new kinds (owner-requested, previously in §2/§4): sunburst,
 * violin, candlestick.
 */
describe("sunburst", () => {
  const cfg: ChartConfig = {
    kind: "sunburst",
    ...DEFAULT_SIZE,
    data: {
      categories: ["G1 | A", "G1 | B", "G2 | C"],
      series: [{ name: "V", values: [30, 10, 40] }],
    },
    decorations: { segmentLabels: false },
  };
  const s = buildChart(cfg);

  it("nests items on an outer ring inside group wedges on the inner ring", () => {
    const g0 = s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === "group-0")!;
    const item = s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === "slice-0")!;
    expect(g0).toBeTruthy();
    expect(item).toBeTruthy();
    // Inner ring (groups) sits inside the outer ring (items).
    expect(g0.r).toBeLessThan(item.r);
    expect(item.innerR).toBeGreaterThanOrEqual(g0.r - 0.01);
    // Group spans are proportional: G1 (40) vs G2 (40) → equal; item A (30) > B (10).
    const a = s.nodes.find((n): n is WedgeNode => n.name === "slice-0")!;
    const b = s.nodes.find((n): n is WedgeNode => n.name === "slice-1")!;
    expect(a.endAngle - a.startAngle).toBeGreaterThan(b.endAngle - b.startAngle);
  });

  it("flat data (no groups) makes a single-ring doughnut", () => {
    const flat = buildChart({ ...cfg, data: { categories: ["A", "B"], series: [{ name: "V", values: [1, 1] }] } });
    expect(flat.nodes.some((n) => n.name?.startsWith("group-"))).toBe(false);
    const slices = flat.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    expect(slices).toHaveLength(2);
    expect(slices.every((w) => w.innerR > 0)).toBe(true);
  });

  it("mixes a loose (ungrouped) label into the grouped ring without throwing", () => {
    // One "|"-less label makes the chart grouped (some label has a "|") yet sends
    // that item through groupOf/labelOf's no-"|" branch — an unnamed "" group.
    const mixed = buildChart({
      ...cfg,
      data: { categories: ["G1 | A", "loose", "G2 | B"], series: [{ name: "V", values: [20, 15, 25] }] },
    });
    // Still a two-ring sunburst: group wedges on the inner ring, items on the outer.
    expect(mixed.nodes.some((n) => n.name?.startsWith("group-"))).toBe(true);
    const slices = mixed.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    expect(slices).toHaveLength(3);
    // Every wedge angle is finite (the loose item did not produce NaN spans).
    for (const w of mixed.nodes.filter((n): n is WedgeNode => n.kind === "wedge")) {
      expect(Number.isFinite(w.startAngle)).toBe(true);
      expect(Number.isFinite(w.endAngle)).toBe(true);
    }
  });
});
