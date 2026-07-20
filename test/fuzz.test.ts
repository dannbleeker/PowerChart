import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { buildChart } from "../src/core/chart";
import type { ChartConfig, ChartKind } from "../src/core/types";

// Keyed by EVERY ChartKind so the fuzzer can never silently skip a kind: the
// Record makes TS error if a kind is missing here or a stale one lingers. The
// hand-maintained array used to omit scatter/bubble/tilemap/treemap/sunburst —
// exactly the kinds recent perf/colour work touched.
const ALL_KINDS: Record<ChartKind, true> = {
  stacked: true,
  clustered: true,
  stacked100: true,
  waterfall: true,
  mekko: true,
  line: true,
  area: true,
  butterfly: true,
  scatter: true,
  bubble: true,
  gantt: true,
  combo: true,
  pie: true,
  doughnut: true,
  boxplot: true,
  radar: true,
  heatmap: true,
  tilemap: true,
  cascade: true,
  funnel: true,
  waffle: true,
  treemap: true,
  sunburst: true,
  violin: true,
  candlestick: true,
};
const KINDS = Object.keys(ALL_KINDS) as ChartKind[];

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

  // Indices are drawn from a range that DELIBERATELY overshoots the data on both
  // ends (negative and past-the-end), so every decoration that anchors to a
  // category/series index is stressed against its bounds guard — the exact class
  // of the callout/difference NaN bugs. clampPair and the range guards must absorb
  // them without producing NaN or off-frame geometry.
  const idx = fc.integer({ min: -2, max: 10 });

  const configArb: fc.Arbitrary<ChartConfig> = fc
    .record({
      kind: fc.constantFrom(...KINDS),
      nCats: fc.integer({ min: 1, max: 8 }),
      nSeries: fc.integer({ min: 1, max: 5 }),
      totals: fc.boolean(),
      wantCagr: fc.boolean(),
      wantDiff: fc.boolean(),
      diffSeries: fc.option(idx, { nil: undefined }),
      wantValueLine: fc.boolean(),
      wantCallout: fc.boolean(),
      calloutCat: idx,
      calloutSeries: fc.option(idx, { nil: undefined }),
      wantBand: fc.boolean(),
      bandAxis: fc.constantFrom("x" as const, "y" as const),
      bandFrom: fc.integer({ min: -2, max: 10 }),
      bandTo: fc.integer({ min: -2, max: 10 }),
      wantQuadrants: fc.boolean(),
      from: idx,
      to: idx,
      perPointColor: fc.boolean(),
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
          const span = { from: base.from, to: base.to };
          return {
            kind: base.kind,
            width: 480,
            height: 300,
            data: {
              categories: Array.from({ length: base.nCats }, (_, i) => `C${i}`),
              series: matrix.map((values, s) => ({
                name: `S${s}`,
                values,
                // A per-point highlight on some categories (exercises series.colors).
                ...(base.perPointColor ? { colors: values.map((_, c) => (c % 2 ? "#ff0000" : null)) } : {}),
              })),
            },
            decorations: {
              totals: base.totals,
              cagr: base.wantCagr ? span : undefined,
              difference: base.wantDiff ? { ...span, series: base.diffSeries } : undefined,
              valueLines: base.wantValueLine ? [{ mode: "mean" }] : undefined,
              callouts: base.wantCallout
                ? [{ text: "note", category: base.calloutCat, series: base.calloutSeries }]
                : undefined,
              bands: base.wantBand ? [{ axis: base.bandAxis, from: base.bandFrom, to: base.bandTo }] : undefined,
              quadrants: base.wantQuadrants ? { x: 0, y: 0 } : undefined,
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
            // Tightened from 5000 to ~2× the 480×300 canvas: enough slack for a
            // label or decoration that legitimately overhangs the edge, but tight
            // enough to actually CATCH off-frame geometry — the old bound could not
            // see a whole series translated 1000px off the canvas.
            expect(Math.abs(c)).toBeLessThan(1000);
          }
        }
      }),
      // Fixed seed for DETERMINISTIC CI (a red build always reproduces locally);
      // fast-check still shrinks a failure to a minimal config. The generous
      // per-test timeout absorbs the 600 runs under coverage instrumentation,
      // which intermittently exceeded vitest's 5s default and flaked the build.
      { numRuns: 600, seed: 20260718 },
    );
  }, 30_000);

  it("is deterministic — the same config renders byte-identically twice", () => {
    fc.assert(
      fc.property(configArb, (cfg) => {
        expect(JSON.stringify(buildChart(cfg).nodes)).toBe(JSON.stringify(buildChart(cfg).nodes));
      }),
      { numRuns: 100, seed: 20260718 },
    );
  });
});
