import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { buildChart } from "../src/core/chart";
import type { ChartConfig, ChartKind } from "../src/core/types";

const KINDS: ChartKind[] = [
  "stacked",
  "clustered",
  "stacked100",
  "waterfall",
  "mekko",
  "line",
  "area",
  "butterfly",
  "combo",
  "pie",
  "doughnut",
  "boxplot",
  "radar",
  "heatmap",
  "cascade",
  "funnel",
  "waffle",
  "violin",
  "candlestick",
  "gantt",
];

/** Every numeric coordinate a node carries, flattened for invariant checks. */
function coordsOf(n: import("../src/core/scene").SceneNode): number[] {
  switch (n.kind) {
    case "rect":
    case "text":
    case "chevron":
      return [n.x, n.y, n.w, n.h];
    case "line":
      return [n.x1, n.y1, n.x2, n.y2];
    case "ellipse":
      return [n.cx, n.cy, n.rx, n.ry];
    case "wedge":
      return [n.cx, n.cy, n.r, n.startAngle, n.endAngle];
    case "polygon":
      return n.points.flatMap((p) => [p.x, p.y]);
    case "symbol":
      return [n.cx, n.cy, n.size];
    case "arrowhead":
      return [n.x, n.y, n.angle, n.size];
  }
}

/**
 * Property-based layout invariants (fast-check, so counterexamples shrink to a
 * minimal reproducing config instead of a raw seed). No random dataset — any mix
 * of negatives, nulls, zeros, single categories, decoration flags, orientation —
 * may ever make `buildChart` throw or emit NaN / runaway geometry.
 */
describe("layout engine properties", () => {
  // A cell: realistic datasheet values (rounded to 1 decimal, as real sheets are —
  // this also keeps subnormals like 5e-324 out, which aren't data, they're float
  // pathology), with nulls and zeros mixed in.
  const cell = fc.oneof(
    {
      weight: 8,
      arbitrary: fc
        .double({ min: -1e4, max: 1e4, noNaN: true, noDefaultInfinity: true })
        .map((v) => Math.round(v * 10) / 10),
    },
    { weight: 1, arbitrary: fc.constant(0) },
    { weight: 1, arbitrary: fc.constant(null) },
  );

  const configArb: fc.Arbitrary<ChartConfig> = fc
    .record({
      kind: fc.constantFrom(...KINDS),
      nCats: fc.integer({ min: 1, max: 8 }),
      nSeries: fc.integer({ min: 1, max: 5 }),
      totals: fc.boolean(),
      wantCagr: fc.boolean(),
      wantDiff: fc.boolean(),
      wantValueLine: fc.boolean(),
      horizontal: fc.boolean(),
      descending: fc.boolean(),
    })
    .chain((base) =>
      fc
        .array(fc.array(cell, { minLength: base.nCats, maxLength: base.nCats }), {
          minLength: base.nSeries,
          maxLength: base.nSeries,
        })
        .map((matrix): ChartConfig => {
          const span = { from: 0, to: base.nCats - 1 };
          return {
            kind: base.kind,
            width: 480,
            height: 300,
            data: {
              categories: Array.from({ length: base.nCats }, (_, i) => `C${i}`),
              series: matrix.map((values, s) => ({ name: `S${s}`, values })),
            },
            decorations: {
              totals: base.totals,
              cagr: base.wantCagr ? span : undefined,
              difference: base.wantDiff ? span : undefined,
              valueLines: base.wantValueLine ? [{ mode: "mean" }] : undefined,
            },
            horizontal: base.kind !== "waterfall" && base.horizontal ? true : undefined,
            segmentOrder: base.descending ? "descending" : undefined,
          };
        }),
    );

  it("never emits NaN or runaway geometry for any dataset", () => {
    fc.assert(
      fc.property(configArb, (cfg) => {
        const scene = buildChart(cfg);
        // Degenerate data may legitimately render nothing: an axis-less pie with a
        // single 0 has no wedge, and the part-to-whole kinds (pie, funnel, waffle)
        // ignore negatives AND read only the first series. A positive value in the
        // FIRST series is the cross-kind floor that must produce output — those kinds
        // draw it, and the cartesian kinds draw axis chrome regardless.
        const firstSeriesPositive = cfg.data.series[0]?.values.some((v) => v != null && v > 0);
        if (firstSeriesPositive) expect(scene.nodes.length).toBeGreaterThan(0);
        for (const node of scene.nodes) {
          for (const c of coordsOf(node)) {
            expect(Number.isFinite(c)).toBe(true);
            expect(Math.abs(c)).toBeLessThan(5000);
          }
        }
      }),
      { numRuns: 400, seed: 20260718 },
    );
  });

  it("is deterministic — the same config renders byte-identically twice", () => {
    fc.assert(
      fc.property(configArb, (cfg) => {
        expect(JSON.stringify(buildChart(cfg).nodes)).toBe(JSON.stringify(buildChart(cfg).nodes));
      }),
      { numRuns: 100, seed: 20260718 },
    );
  });
});
