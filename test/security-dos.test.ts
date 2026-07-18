import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * buildChart is fed arbitrary authored JSON (the skill) and shape-tag configs.
 * A config claiming a huge grid must not be able to drive unbounded allocation
 * or super-linear CPU — the size caps bound both without touching real charts.
 */

describe("grid-size caps bound abusive configs", () => {
  it("truncates an over-cap category count instead of allocating it", () => {
    const cfg: ChartConfig = {
      kind: "stacked",
      width: 480,
      height: 300,
      data: {
        categories: Array.from({ length: 100_000 }, (_, i) => `C${i}`),
        series: [{ name: "S", values: [] }],
      },
    };
    const scene = buildChart(cfg);
    // Renders (does not OOM) and the per-category node count is bounded.
    const cols = scene.nodes.filter((n) => n.name?.startsWith("seg-")).length;
    expect(cols).toBeLessThanOrEqual(4096);
  });

  it("truncates an over-cap series count", () => {
    const cfg: ChartConfig = {
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B"],
        series: Array.from({ length: 5000 }, (_, i) => ({ name: `S${i}`, values: [1, 2] })),
      },
    };
    // Completes quickly and doesn't throw.
    expect(() => buildChart(cfg)).not.toThrow();
  });

  it("skips heatmap clustering above the row cap (no CPU blow-up, no dendrogram)", () => {
    const cfg: ChartConfig = {
      kind: "heatmap",
      width: 600,
      height: 800,
      heatmap: { cluster: true },
      data: {
        categories: ["a", "b", "c", "d"],
        series: Array.from({ length: 200 }, (_, i) => ({
          name: `R${i}`,
          values: [i % 5, (i * 3) % 7, i % 2, (i * 2) % 4],
        })),
      },
    };
    const start = Date.now();
    const nodes = buildChart(cfg).nodes;
    // O(rows³) clustering on 200 rows would take seconds; the cap keeps it instant.
    expect(Date.now() - start).toBeLessThan(1000);
    // Above the cap no dendrogram is drawn.
    expect(nodes.some((n) => n.name === "dendro-v")).toBe(false);
  });

  it("still clusters a small heatmap (cap is a ceiling, not a switch-off)", () => {
    const cfg: ChartConfig = {
      kind: "heatmap",
      width: 600,
      height: 400,
      heatmap: { cluster: true },
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "lo1", values: [1, 2, 1] },
          { name: "hi1", values: [90, 91, 89] },
          { name: "lo2", values: [2, 1, 2] },
          { name: "hi2", values: [88, 90, 91] },
        ],
      },
    };
    expect(buildChart(cfg).nodes.some((n) => n.name === "dendro-v")).toBe(true);
  });
});
