import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * Two combo line-overlay scale bugs found by a layout-modes hunt: both put a line
 * series off the canvas (finite but wrong coordinates). Byte-identical for the
 * valid cases (positive line / plain-chain waterfall) — snapshots unchanged — and
 * only rescue the broken configs.
 */

const HEIGHT = 300;
const markerYs = (cfg: ChartConfig) =>
  buildChart(cfg)
    .nodes.filter((n) => n.name?.startsWith("combo-marker") || n.name?.startsWith("combo-line"))
    .flatMap((n) => (n.kind === "rect" ? [n.y, n.y + n.h] : n.kind === "line" ? [n.y1, n.y2] : []));

describe("combo waterfall wfMax ignores off-chain columns", () => {
  it("keeps a tall shared-axis line on-plot over a detail-group waterfall", () => {
    // detailGroups[2,3] decompose column 1 in place — they do NOT advance the
    // chain (peak = 150). The old wfMax summed them (200), matched the line's 200,
    // so the overflow stretch never fired and the line clipped off the top.
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: HEIGHT,
      combo: { columns: "waterfall" },
      waterfall: { detailGroups: [{ of: 1, indices: [2, 3] }] },
      data: {
        categories: ["a", "b", "c", "d", "e"],
        series: [
          { name: "delta", values: [100, 50, 30, 20, null] },
          { name: "line", type: "line", values: [200, 200, 200, 200, 200] },
        ],
      },
    };
    const ys = markerYs(cfg);
    expect(ys.length).toBeGreaterThan(0);
    for (const y of ys) expect(y).toBeGreaterThanOrEqual(-1);
  });
});

describe("combo secondary axis spans negative line values", () => {
  it("keeps a negative overlay line on-plot", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: HEIGHT,
      secondaryAxis: true,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "bars", values: [10, 20, 30] },
          { name: "line", type: "line", values: [-5, 8, -3] },
        ],
      },
    };
    const ys = markerYs(cfg);
    expect(ys.length).toBeGreaterThan(0);
    for (const y of ys) expect(y).toBeLessThanOrEqual(HEIGHT + 1);
  });

  it("still renders an all-positive secondary line on-plot (unchanged path)", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: HEIGHT,
      secondaryAxis: true,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "bars", values: [10, 20, 30] },
          { name: "line", type: "line", values: [5, 8, 3] },
        ],
      },
    };
    const ys = markerYs(cfg);
    expect(ys.length).toBeGreaterThan(0);
    for (const y of ys) {
      expect(y).toBeGreaterThanOrEqual(-1);
      expect(y).toBeLessThanOrEqual(HEIGHT + 1);
    }
  });
});
