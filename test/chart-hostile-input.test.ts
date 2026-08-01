import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import { niceTicks } from "../src/core/format";
import { BoxHash } from "../src/core/grid";
import { sampleConfig, CHART_KINDS } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";

/**
 * Values a datasheet cell can hold, against every chart kind.
 *
 * Both bugs this file pins are the same shape: a loop whose bound came from
 * the data, with nothing between the number someone typed and the number of
 * iterations. Neither was slow — both were unbounded, in time and in memory,
 * and the tab dies before the chart does.
 */

const HOSTILE: [string, (c: ChartConfig) => ChartConfig][] = [
  ["huge values", (c) => vals(c, () => 1e308)],
  ["tiny values", (c) => vals(c, () => 5e-324)],
  ["NaN value", (c) => vals(c, () => NaN)],
  ["Infinity value", (c) => vals(c, () => Infinity)],
  ["-Infinity value", (c) => vals(c, () => -Infinity)],
  ["all zero", (c) => vals(c, () => 0)],
  ["negative values", (c) => vals(c, (v) => -Math.abs(Number(v) || 1))],
  ["no series", (c) => ({ ...c, data: { ...c.data, series: [] } })],
  ["empty values", (c) => ({ ...c, data: { ...c.data, series: c.data.series.map((s) => ({ ...s, values: [] })) } })],
  ["one category", (c) => ({ ...c, data: { ...c.data, categories: c.data.categories.slice(0, 1) } })],
  ["no categories", (c) => ({ ...c, data: { ...c.data, categories: [] } })],
  ["zero size", (c) => ({ ...c, width: 0, height: 0 })],
  ["negative size", (c) => ({ ...c, width: -100, height: -50 })],
  ["nan size", (c) => ({ ...c, width: NaN, height: NaN })],
];

function vals(c: ChartConfig, f: (v: unknown) => number): ChartConfig {
  return { ...c, data: { ...c.data, series: c.data.series.map((s) => ({ ...s, values: s.values.map(f) })) } };
}

describe("a hostile number in a cell cannot hang the chart engine", () => {
  it("builds every kind against every hostile input, and returns", () => {
    // The whole grid, timed as one. Before the two fixes below this did not
    // fail — it never finished, and vitest killed the run at its timeout.
    const t0 = Date.now();
    for (const { kind } of CHART_KINDS) {
      for (const [name, mutate] of HOSTILE) {
        expect(() => buildChart(mutate(sampleConfig(kind))), `${kind} / ${name}`).not.toThrow();
      }
    }
    expect(Date.now() - t0).toBeLessThan(30_000);
  }, 60_000);

  it("does not fill memory generating ticks for an axis near MAX_VALUE", () => {
    // `ceil(1.7e308 / 5e307)` is 4, and 4 x 5e307 is Infinity — so the loop
    // bound was Infinity while both its inputs were finite. The entry guard
    // checks the inputs; these are outputs.
    for (const max of [1e300, 1e307, 1.7e308, Number.MAX_VALUE]) {
      const ticks = niceTicks(0, max);
      expect(ticks.length, `max=${max}`).toBeGreaterThan(0);
      expect(ticks.length, `max=${max}`).toBeLessThan(1100);
      for (const t of ticks) expect(Number.isFinite(t), `max=${max} tick ${t}`).toBe(true);
    }
  });

  it("does not enumerate a grid cell per unit of a runaway box", () => {
    // The label index sizes its cell to the largest label, so a real label
    // covers a handful of cells. A label placed at 1e300 — which a value of
    // 1e308 produces — spans about 1e298 of them, and each one allocated a Map
    // key. One box, unbounded memory.
    const hash = new BoxHash<number>(8);
    const t0 = Date.now();
    hash.insert({ x: 0, y: 0, w: 1e300, h: 1e300 }, 1);
    hash.insert({ x: NaN, y: NaN, w: NaN, h: NaN }, 2);
    hash.insert({ x: 0, y: 0, w: Infinity, h: Infinity }, 3);
    expect(Date.now() - t0).toBeLessThan(1000);
    // The finite-but-huge box is still findable — degenerate, not discarded.
    expect(hash.some({ x: 0, y: 0, w: 8, h: 8 }, (v) => v === 1)).toBe(true);
  });

  it("still indexes an ordinary label across the cells it really covers", () => {
    // The negative control: the cap must not have turned the index into a
    // single bucket, which would make every label collide with every other.
    const hash = new BoxHash<string>(8);
    hash.insert({ x: 0, y: 0, w: 20, h: 20 }, "near");
    hash.insert({ x: 400, y: 400, w: 20, h: 20 }, "far");
    expect(hash.some({ x: 0, y: 0, w: 4, h: 4 }, (v) => v === "near")).toBe(true);
    expect(hash.some({ x: 0, y: 0, w: 4, h: 4 }, (v) => v === "far")).toBe(false);
  });
});
