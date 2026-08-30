import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * HOW MANY SHAPES A CHART IS ALLOWED TO BE, and nothing has ever capped it.
 *
 * Shapes are the unit of cost on PowerPoint on the web, and the archive is
 * unambiguous about the price. An in-place update runs about **767ms per shape**
 * — measured over 178 updates — which makes `updated only the shapes that
 * changed` 11.8 calls and 108 seconds a round, a quarter of a round, with a
 * worst single chart at 26.4s. It is not batchable: within one update the chunks
 * cost 4514, 4366 and 4657ms, flat, so the cost is linear in shapes written and
 * server-side.
 *
 * And past a point it stops being slow and starts being fatal. Round 150
 * recorded **400-500 shapes in one context crashing PowerPoint on all seven
 * attempts**, and before the slab rule took the rise into account a chart in the
 * shipped deck sat at 300.
 *
 * So this is the third ratchet in the house style — `overlap-budget` for the
 * variant sweep, `showcase-overlap` for the shipped deck, and this for density.
 * Same rules: a kind over budget is a regression, an unbudgeted kind is a
 * regression by definition, and a budget left ABOVE the real figure fails, so
 * making a chart lighter forces the number down instead of banking slack.
 *
 * WHAT THIS IS NOT. It is not a claim that any of these figures is too high; it
 * is a claim that they should not RISE without somebody saying so. The one that
 * would repay attention is `line` at 226, which is the smoothed Catmull-Rom
 * chart — it overtook the stacked area as the densest thing we ship only because
 * the area got lighter.
 */

/** Worst chart of each kind in `examples/showcase.json`, as of 2026-08-30. */
const BUDGET: Record<string, number> = {
  line: 226,
  combo: 164,
  tilemap: 146,
  area: 123,
  waffle: 107,
  heatmap: 100,
  scatter: 75,
  gantt: 68,
  boxplot: 62,
  butterfly: 53,
  stacked: 47,
  bubble: 46,
  clustered: 40,
  radar: 36,
  candlestick: 35,
  stacked100: 33,
  mekko: 32,
  waterfall: 31,
  treemap: 31,
  cascade: 30,
  sunburst: 21,
  funnel: 20,
  pie: 17,
  violin: 16,
  doughnut: 15,
};

/** The whole deck, so a rise spread thinly across many charts still shows. */
const TOTAL_BUDGET = 5560;

function measure(): { byKind: Map<string, number>; total: number; worst: Map<string, string> } {
  const items = JSON.parse(readFileSync("examples/showcase.json", "utf8")) as ChartConfig[];
  const byKind = new Map<string, number>();
  const worst = new Map<string, string>();
  let total = 0;
  for (const cfg of items) {
    let n: number;
    try {
      n = buildChart(cfg).nodes.length;
    } catch {
      // A config the engine refuses is not a density problem, and
      // `showcase.test.ts` is what holds the deck to building at all.
      continue;
    }
    total += n;
    const kind = String(cfg.kind ?? "?");
    if (n > (byKind.get(kind) ?? 0)) {
      byKind.set(kind, n);
      worst.set(kind, String(cfg.title ?? kind));
    }
  }
  return { byKind, total, worst };
}

describe("no chart quietly becomes too heavy to draw", () => {
  const { byKind, total, worst } = measure();

  it("draws no kind it has never drawn before", () => {
    const unbudgeted = [...byKind.keys()].filter((k) => BUDGET[k] === undefined);
    expect(unbudgeted, "a new chart kind reached the shipped deck — budget it here and say what it is").toEqual([]);
  });

  it("keeps every kind inside its budget", () => {
    const over = [...byKind.entries()]
      .filter(([k, n]) => n > (BUDGET[k] ?? 0))
      .map(([k, n]) => `${k}: ${n} nodes (budget ${BUDGET[k] ?? 0}) — ${worst.get(k)}`);
    expect(over, "a chart kind got denser, and shapes are what this host charges for").toEqual([]);
  });

  it("holds no budget ABOVE the real figure", () => {
    // The half that makes it a ratchet rather than a ceiling. Making a chart
    // lighter is supposed to fail this file so the number gets edited down —
    // which is exactly what the slab rule did to `area`, 300 -> 123.
    const slack = Object.entries(BUDGET)
      .filter(([k, n]) => (byKind.get(k) ?? 0) < n)
      .map(([k, n]) => `${k}: budget ${n}, actual ${byKind.get(k) ?? 0}`);
    expect(slack, "a budget here is above the real figure — edit it down").toEqual([]);
  });

  it("has not grown across the deck as a whole", () => {
    // Per-kind ceilings cannot see a rise spread thinly over many charts, which
    // is the shape a data change makes.
    expect(total, `${total} nodes across the shipped deck, budget ${TOTAL_BUDGET}`).toBeLessThanOrEqual(TOTAL_BUDGET);
  });

  it("keeps every shipped chart clear of the count that crashed the host", () => {
    /**
     * The hard one, and the reason this file exists rather than a comment.
     * Round 150: 400-500 shapes in ONE context, PowerPoint down on all seven
     * attempts. A single chart is not a whole context — a slide can hold
     * several — so the per-chart line sits well below it.
     */
    const heavy = [...byKind.entries()].filter(([, n]) => n >= 300).map(([k, n]) => `${k}: ${n} — ${worst.get(k)}`);
    expect(heavy, "a shipped chart is within range of the count that crashed PowerPoint").toEqual([]);
  });
});
