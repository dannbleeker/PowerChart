import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { buildChart } from "../src/core/chart";
import { textWidth, type SceneNode, type TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * TEXT OVER TEXT IN THE DECK WE SHIP, measured by the frame gate's own ink rule.
 *
 * A THIRD POPULATION, and it took two blind spots to notice it was missing.
 * `frame-fit.test.ts` sweeps every kind at synthetic frames and fonts.
 * `overlap-budget.test.ts` sweeps option and data-shape variants OF THE
 * SAMPLES. Neither of them builds these: 123 curated, realistic configs at the
 * sizes they ship at, carrying options the samples do not — a boxplot with both
 * mean and median marks, a scatter with fit statistics, a bump chart.
 *
 * It is also the population that matters most. `examples/showcase.pptx` is
 * linked from the README and is what somebody looks at to decide whether to
 * install this. When it first ran, on 2026-08-29, six of its charts drew text
 * through other text at 480x300 — and FIVE of the six shapes involved were
 * shapes the ratchet had never seen, because its sweep starts from
 * `sampleConfig(kind)` and these configs do not.
 *
 * Ratcheted the same way as the budget: a per-shape ceiling, an unknown shape
 * is a regression by definition, and a ceiling left above the real figure fails
 * so improving the engine forces the number down.
 */

/** EXACTLY `inkBox` from the frame gate, so all three gates measure one thing. */
const ink = (t: TextNode) => {
  const w = textWidth(t.text, t.fontSize, t.bold);
  const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
  const base =
    t.valign === "top"
      ? t.y + t.fontSize
      : t.valign === "bottom"
        ? t.y + t.h - t.fontSize * 0.25
        : t.y + t.h / 2 + t.fontSize * 0.36;
  return { x0: x, y0: base - t.fontSize * 0.8, x1: x + w, y1: base + t.fontSize * 0.21 };
};

const sharedArea = (a: ReturnType<typeof ink>, b: ReturnType<typeof ink>) => {
  const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
  const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
  return w > 0 && h > 0 ? w * h : 0;
};

/**
 * What each shape is allowed in the shipped deck, as of 2026-08-29. Total 4.
 *
 * It opened at 9 across six charts, and five went the same morning:
 *
 *     mean# / median-label#   3 -> 0   the boxplot's median number sat under
 *                                      the mean's `x`; the two are pushed by
 *                                      one loop and neither knew the other
 *     trend-stats / label#    2 -> 0   the scatter's R² was fitted against
 *                                      nothing, and could not use the global
 *                                      de-collision pass because scatter is a
 *                                      `decorlessKind` and never runs it
 *
 * What is left is one shape of the `valueAxisTitle` family — the open decision
 * in docs/BACKLOG.md — and two that are their own small defects, named here so
 * they are not mistaken for it.
 */
const BUDGET: Record<string, number> = {
  // The dumbbell's two end labels at one category.
  "label## / label##": 1,
  // The bump chart's period heading over a rank label at the same end.
  "period# / bump-label-l#": 1,
  "period# / bump-label-r#": 1,
  // The open decision: where a unit belongs when the band above the plot
  // cannot hold the chart title, the unit and the topmost tick all three.
  "value-axis / value-axis-title": 1,
};

function measure(): { total: number; byShape: Map<string, number>; charts: Map<string, number> } {
  const items = JSON.parse(readFileSync("examples/showcase.json", "utf8")) as ChartConfig[];
  const byShape = new Map<string, number>();
  const charts = new Map<string, number>();
  let total = 0;
  for (const cfg of items) {
    let ts: TextNode[];
    try {
      ts = buildChart(cfg).nodes.filter((n: SceneNode): n is TextNode => n.kind === "text" && !!String(n.text).trim());
    } catch {
      // A config the engine refuses is not an overlap, and `showcase.test.ts`
      // is what holds the deck to building at all.
      continue;
    }
    const boxes = ts.map(ink);
    for (let i = 0; i < boxes.length; i++)
      for (let j = i + 1; j < boxes.length; j++)
        if (sharedArea(boxes[i], boxes[j]) > 1) {
          total++;
          const a = (ts[i].name || "?").replace(/-?\d+/g, "#");
          const b = (ts[j].name || "?").replace(/-?\d+/g, "#");
          byShape.set(`${a} / ${b}`, (byShape.get(`${a} / ${b}`) ?? 0) + 1);
          const title = String(cfg.title ?? cfg.kind);
          charts.set(title, (charts.get(title) ?? 0) + 1);
        }
  }
  return { total, byShape, charts };
}

describe("the deck we ship draws no text through other text", () => {
  const { total, byShape, charts } = measure();

  it("draws no shape it has never drawn before", () => {
    const unbudgeted = [...byShape.entries()].filter(([k]) => BUDGET[k] === undefined);
    expect(
      unbudgeted,
      "a new overlap appeared in examples/showcase.json — fix it, or budget it here and say what it is",
    ).toEqual([]);
  });

  it("stays inside every shape's budget", () => {
    const over = [...byShape.entries()]
      .filter(([k, v]) => v > (BUDGET[k] ?? 0))
      .map(([k, v]) => `${k}: ${v} (budget ${BUDGET[k] ?? 0})`);
    expect(over, "an overlap in the shipped deck got worse").toEqual([]);
  });

  it("holds no budget ABOVE the real figure", () => {
    // The half that makes this a ratchet rather than a ceiling: improving the
    // engine is supposed to fail this file so the numbers get edited down.
    const slack = Object.entries(BUDGET)
      .filter(([k, v]) => (byShape.get(k) ?? 0) < v)
      .map(([k, v]) => `${k}: budget ${v}, actual ${byShape.get(k) ?? 0}`);
    expect(slack, "a budget here is above the real figure — edit it down").toEqual([]);
  });

  it("has not grown in total", () => {
    const allowed = Object.values(BUDGET).reduce((a, b) => a + b, 0);
    expect(total, `${total} overlapping pairs in the shipped deck, budget ${allowed}`).toBeLessThanOrEqual(allowed);
  });

  it("names the charts, so a failure says which slide to look at", () => {
    // Not an assertion about the count — an assertion that the diagnosis is
    // available at all. A gate over 123 charts that only says "3" sends the
    // reader to open all of them.
    if (total > 0) expect(charts.size, "overlaps were counted but no chart was named").toBeGreaterThan(0);
    for (const [name, n] of charts) expect(typeof name === "string" && n > 0).toBe(true);
  });
});
