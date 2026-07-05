import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig, ChartKind } from "../src/core/types";

/** Deterministic PRNG so failures reproduce (mulberry32). */
function rng(seed: number) {
  return () => {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const KINDS: ChartKind[] = [
  "stacked", "clustered", "stacked100", "waterfall", "mekko",
  "line", "area", "butterfly", "combo", "pie", "doughnut",
  "boxplot", "radar", "heatmap", "cascade", "funnel",
];

/**
 * Property-based smoke test: random datasets (including negatives, nulls,
 * zeros, single categories) must never produce NaN geometry or nodes wildly
 * outside the frame.
 */
describe("layout engine fuzz", () => {
  const rand = rng(20260705);
  for (let iter = 0; iter < 60; iter++) {
    const kind = KINDS[Math.floor(rand() * KINDS.length)];
    const nCats = 1 + Math.floor(rand() * 7);
    const nSeries = 1 + Math.floor(rand() * 4);
    const cfg: ChartConfig = {
      kind,
      width: 480,
      height: 300,
      data: {
        categories: Array.from({ length: nCats }, (_, i) => `C${i}`),
        series: Array.from({ length: nSeries }, (_, s) => ({
          name: `S${s}`,
          values: Array.from({ length: nCats }, () => {
            const roll = rand();
            if (roll < 0.08) return null;
            if (roll < 0.14) return 0;
            const v = (rand() - 0.25) * 200;
            return Math.round(v * 10) / 10;
          }),
        })),
      },
      decorations: {
        totals: rand() < 0.4,
        cagr: rand() < 0.25 ? { from: 0, to: nCats - 1 } : undefined,
        difference: rand() < 0.25 ? { from: 0, to: nCats - 1 } : undefined,
        valueLines: rand() < 0.2 ? [{ mode: "mean" }] : undefined,
      },
      horizontal: kind !== "waterfall" && rand() < 0.2 ? true : undefined,
      segmentOrder: rand() < 0.25 ? "descending" : undefined,
    };

    it(`#${iter} ${kind} ${nSeries}x${nCats}`, () => {
      const scene = buildChart(cfg);
      expect(scene.nodes.length).toBeGreaterThan(0);
      for (const n of scene.nodes) {
        const coords =
          n.kind === "rect" || n.kind === "text"
            ? [n.x, n.y, n.w, n.h]
            : n.kind === "line"
              ? [n.x1, n.y1, n.x2, n.y2]
              : n.kind === "ellipse"
                ? [n.cx, n.cy, n.rx, n.ry]
                : n.kind === "wedge"
                  ? [n.cx, n.cy, n.r, n.startAngle, n.endAngle]
                  : n.kind === "chevron"
                    ? [n.x, n.y, n.w, n.h]
                    : n.kind === "polygon"
                      ? n.points.flatMap((p) => [p.x, p.y])
                      : [n.x, n.y, n.angle, n.size];
        for (const c of coords) {
          expect(Number.isFinite(c)).toBe(true);
          expect(Math.abs(c)).toBeLessThan(5000);
        }
      }
    });
  }
});
