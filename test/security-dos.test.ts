import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * buildChart is fed arbitrary authored JSON (the skill) and shape-tag configs.
 * A config claiming a huge grid must not be able to drive unbounded allocation
 * or super-linear CPU — the size caps bound both without touching real charts.
 */

describe("grid-size caps bound abusive configs", () => {
  it("truncates an over-cap category count to exactly MAX_CATEGORIES", () => {
    const cfg: ChartConfig = {
      kind: "stacked",
      width: 480,
      height: 300,
      data: {
        // Non-null values, or every padded cell is null and NO seg- node is emitted —
        // which made the old `<= 4096` assertion trivially 0 <= 4096, passing even
        // with the cap deleted.
        categories: Array.from({ length: 100_000 }, (_, i) => `C${i}`),
        series: [{ name: "S", values: Array.from({ length: 100_000 }, () => 1) }],
      },
    };
    const scene = buildChart(cfg);
    // One segment per surviving category: the cap is OBSERVED, not just an upper bound.
    const cols = scene.nodes.filter((n) => n.name?.startsWith("seg-")).length;
    expect(cols).toBe(4096); // === MAX_CATEGORIES
  });

  it("truncates an over-cap series count to exactly MAX_SERIES", () => {
    const cfg: ChartConfig = {
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B"],
        series: Array.from({ length: 5000 }, (_, i) => ({ name: `S${i}`, values: [1, 2] })),
      },
    };
    const scene = buildChart(cfg);
    // Highest surviving series index proves the cap fired (256 series, 0-based → 255),
    // instead of the old "does not throw" which held whether or not the cap existed.
    const maxSeries = Math.max(
      -1,
      ...scene.nodes.map((n) => Number(/^seg-(\d+)-/.exec(n.name ?? "")?.[1] ?? -1)).filter((i) => i >= 0),
    );
    expect(maxSeries).toBe(255); // MAX_SERIES - 1
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
