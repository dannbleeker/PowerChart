import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig, ChartKind } from "../src/core/types";
import type { PolygonNode } from "../src/core/scene";

/**
 * Degenerate-input guards. Every layout has early-returns for data that carries
 * no drawable shape — an empty column, a single observation, a zero-width
 * distribution. Sample and showcase data never trips them, so they were the
 * layouts' uncovered corner; a hand-authored config reaches them, and the
 * contract is the same everywhere: degrade gracefully, never throw, never emit
 * NaN geometry.
 */

const base = (kind: ChartKind, data: ChartConfig["data"]): ChartConfig => ({
  kind,
  ...DEFAULT_SIZE,
  data,
});

/** Assert a scene carries only finite coordinates — no NaN leaked from a guard. */
function expectFiniteGeometry(cfg: ChartConfig) {
  const scene = buildChart(cfg);
  for (const n of scene.nodes) {
    for (const [k, v] of Object.entries(n)) {
      if (typeof v === "number") expect(Number.isFinite(v), `${n.kind}.${k}`).toBe(true);
    }
  }
  return scene;
}

describe("violin — degenerate columns", () => {
  const cats = ["Empty", "One", "Spread"];

  it("survives a chart with no observations at all (empty value axis)", () => {
    // Every value missing → the flattened allSamples is empty, so the value axis
    // falls back to [0, 1] instead of computing min/max over nothing.
    const cfg = base("violin", {
      categories: ["A", "B"],
      series: [
        { name: "x", values: [null as unknown as number, null as unknown as number] },
        { name: "y", values: [null as unknown as number, null as unknown as number] },
      ],
    });
    expect(() => buildChart(cfg)).not.toThrow();
    const scene = expectFiniteGeometry(cfg);
    // No column has two observations, so no density body is drawn.
    expect(scene.nodes.filter((n) => n.kind === "polygon")).toHaveLength(0);
  });

  it("skips a category whose observations are all missing (empty samples)", () => {
    // Column 0 is entirely null → samplesOf(0) empty; column 2 has a real spread.
    // Hits the allSamples-empty axis fallback and the per-column length<2 skip.
    const cfg = base("violin", {
      categories: cats,
      series: [
        { name: "a", values: [null as unknown as number, 10, 20] },
        { name: "b", values: [null as unknown as number, null as unknown as number, 60] },
        { name: "c", values: [null as unknown as number, null as unknown as number, 40] },
      ],
    });
    const scene = expectFiniteGeometry(cfg);
    // The empty and single-point columns draw no violin body; the spread one may.
    const bodies = scene.nodes.filter((n): n is PolygonNode => n.kind === "polygon");
    expect(bodies.every((b) => b.points.every((p) => Number.isFinite(p.x) && Number.isFinite(p.y)))).toBe(true);
  });

  it("draws nothing and does not throw when every observation is identical", () => {
    // A zero-width distribution collapses the KDE support window (gHi <= gLo) —
    // the guard that must return instead of dividing a flat density.
    const cfg = base("violin", {
      categories: ["Flat"],
      series: Array.from({ length: 6 }, (_, i) => ({ name: `s${i}`, values: [42] })),
    });
    expect(() => buildChart(cfg)).not.toThrow();
    expectFiniteGeometry(cfg);
  });

  it("handles a single lonely observation per category", () => {
    const cfg = base("violin", {
      categories: ["Lonely"],
      series: [{ name: "only", values: [7] }],
    });
    const scene = expectFiniteGeometry(cfg);
    expect(scene.nodes.filter((n) => n.kind === "polygon")).toHaveLength(0);
  });
});

describe("layouts — empty data never throws", () => {
  // A cheap breadth check: an empty column-set is a real config a user can build
  // by clearing the datasheet, and each layout must survive it.
  const kinds: ChartKind[] = ["stacked", "clustered", "line", "area", "waterfall", "mekko", "pie", "scatter"];
  const empty: ChartConfig["data"] = { categories: [], series: [] };

  for (const kind of kinds) {
    it(`${kind}: empty data yields a finite, shape-safe scene`, () => {
      expect(() => buildChart(base(kind, empty))).not.toThrow();
      expectFiniteGeometry(base(kind, empty));
    });
  }
});
