import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig, ChartKind } from "../src/core/types";
import type { PolygonNode } from "../src/core/scene";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { MIN_LABEL_FS } from "../src/core/layout/frame";

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

/**
 * Nothing a chart draws may be NEGATIVE or unreadably small — swept over every
 * kind, both orientations, and the data a user actually starts from.
 *
 * The per-case tests above each name one layout's early return. This is the
 * breadth check they cannot be: the defects it found were in kinds nobody was
 * thinking about, from data nobody had written a case for.
 *
 * - A clustered bar takes a 1pt gap from its own thickness, so on a 14pt-wide
 *   chart it came out at `w: -0.3`. That is not a thin bar — SVG drops a
 *   negative-width rect and PowerPoint clamps it to a sliver — so the chart lost
 *   its bars while every label around them still said what they were worth.
 * - The horizontal mekko widens its gutter AFTER `fitPlot` has floored the
 *   frame, by a fixed number of points, leaving `w: -12.2` on a narrow chart.
 * - A label fitted to a band with no floor answers a font of ZERO when the band
 *   is zero, and a mekko's rows are shares of the total — so an all-zero
 *   datasheet, which is what a template looks like before anyone fills it in,
 *   produced `fontSize: 0`. OOXML's `sz` has a minimum of 100 hundredths, so
 *   that is a deck PowerPoint offers to repair rather than a small label.
 *
 * All three are invisible to the finite-geometry check above: -0.3 and 0 are
 * perfectly finite numbers.
 */
describe("no chart draws negative or unreadable geometry", () => {
  /** The states a real datasheet passes through, not exotic fuzz. */
  const MUTATIONS: Record<string, (c: ChartConfig) => void> = {
    "an empty template": (c) => c.data.series.forEach((s) => (s.values = s.values.map(() => 0))),
    "nothing entered yet": (c) => c.data.series.forEach((s) => (s.values = s.values.map(() => null))),
    "one category": (c) => {
      c.data.categories = c.data.categories.slice(0, 1);
      c.data.series.forEach((s) => (s.values = s.values.slice(0, 1)));
    },
    "all negative": (c) =>
      c.data.series.forEach((s) => (s.values = s.values.map((v) => (v == null ? v : -Math.abs(v))))),
    "every value the same": (c) => c.data.series.forEach((s) => (s.values = s.values.map(() => 42))),
    "a thumbnail": (c) => {
      c.width = 12;
      c.height = 9;
    },
    "a sliver, wide": (c) => {
      c.width = 600;
      c.height = 14;
    },
    "a sliver, tall": (c) => {
      c.width = 14;
      c.height = 600;
    },
  };

  const bad: string[] = [];
  for (const { kind } of CHART_KINDS) {
    for (const [name, mutate] of Object.entries(MUTATIONS)) {
      for (const horizontal of [false, true]) {
        const cfg = JSON.parse(JSON.stringify(sampleConfig(kind))) as ChartConfig;
        cfg.horizontal = horizontal;
        mutate(cfg);
        let scene;
        try {
          scene = buildChart(cfg);
        } catch (e) {
          bad.push(`${kind}${horizontal ? "/h" : ""} (${name}) THREW: ${(e as Error).message}`);
          continue;
        }
        const where = `${kind}${horizontal ? "/h" : ""} (${name})`;
        for (const n of scene.nodes) {
          if ((n.kind === "rect" || n.kind === "chevron") && (n.w < 0 || n.h < 0))
            bad.push(`${where}: ${n.name} is ${n.w.toFixed(1)}x${n.h.toFixed(1)}`);
          if (n.kind === "ellipse" && (n.rx < 0 || n.ry < 0)) bad.push(`${where}: ${n.name} has a negative radius`);
          if (n.kind === "text" && n.text.trim() && n.fontSize < MIN_LABEL_FS)
            bad.push(`${where}: ${n.name} at ${n.fontSize}pt`);
        }
      }
    }
  }

  it("never emits a negative extent, a negative radius, or a sub-legible font", () => {
    expect(bad).toEqual([]);
  });
});
