import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { textWidth, type TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * THE UNIT LABEL IS A SHORT UNIT, and clipping is what makes that true of the
 * layout as well as of the documentation.
 *
 * `ChartConfig.valueAxisTitle` is documented as "units label shown at the top
 * of the value axis (e.g. `€m`)", and the two uses in the shipped showcase are
 * `€m` and `$m (log)`. It was drawn at a fixed size with no width bound at all,
 * so a long string ran clear across the chart — and the variant sweep, which
 * tests it with a twenty-seven-character sentence, made that the single largest
 * overlap family in the engine.
 *
 * Clipping is the remedy that keeps AUTHOR TEXT. The three tried on 2026-08-28
 * all dropped the unit outright; this one truncates what cannot fit and leaves
 * `€m` untouched. Across the sweep: 1,327 overlapping pairs to 1,014.
 */
const unitNode = (cfg: ChartConfig): TextNode | undefined =>
  buildChart(cfg).nodes.find((n) => n.name === "value-axis-title") as TextNode | undefined;

const base = (over: Partial<ChartConfig> = {}): ChartConfig =>
  ({ ...sampleConfig("clustered"), ...DEFAULT_SIZE, ...over }) as ChartConfig;

describe("the value-axis unit label", () => {
  it("leaves a short unit exactly as the author typed it", () => {
    // The case the option exists for, and the one the showcase uses. Clipping
    // must be invisible here or it is not a clip, it is a truncation.
    for (const unit of ["€m", "$m (log)", "%", "EUR millions"]) {
      const n = unitNode(base({ valueAxisTitle: unit }));
      expect(n, `no unit drawn for ${unit}`).toBeTruthy();
      expect(n!.text, `${unit} was altered`).toBe(unit);
    }
  });

  it("clips a unit that would otherwise span the chart", () => {
    const long = "Revenue in millions of euro, restated on the 2024 perimeter";
    const n = unitNode(base({ valueAxisTitle: long }));
    expect(n, "a long unit was dropped rather than clipped").toBeTruthy();
    // Shortened, and still the beginning of what the author wrote — which is
    // the half that carries the meaning.
    expect(n!.text.length, "a long unit was not clipped at all").toBeLessThan(long.length);
    expect(long.startsWith(n!.text.replace(/…$/, "")), "the clip did not keep the start").toBe(true);
  });

  it("never draws the unit wider than its share of the chart", () => {
    // The bound itself, over every kind and a range of frames. Its absence is
    // what let a unit run across the totals row and the legend.
    const long = "Revenue in millions of euro, restated on the 2024 perimeter";
    for (const kind of ["clustered", "stacked", "line", "area", "waterfall"] as const)
      for (const [w, h] of [
        [200, 150],
        [480, 300],
        [960, 540],
      ] as [number, number][]) {
        const n = unitNode(base({ kind, width: w, height: h, valueAxisTitle: long }));
        if (!n) continue;
        const inkW = textWidth(n.text, n.fontSize, n.bold);
        expect(inkW, `${kind} at ${w}x${h} drew a ${inkW.toFixed(0)}pt unit on a ${w}pt chart`).toBeLessThanOrEqual(
          w * 0.4 + 0.5,
        );
      }
  });

  it("drops the unit rather than drawing an empty box when nothing fits", () => {
    // A 60pt-wide chart at a large font has room for nothing. An empty text
    // node would be litter every readback and the de-collision pass then carry.
    const n = unitNode(base({ width: 60, height: 300, style: { fontSize: 18 }, valueAxisTitle: "Revenue in euro" }));
    if (n) expect(String(n.text).trim().length, "an empty unit box was pushed into the scene").toBeGreaterThan(0);
  });

  /**
   * THE `y` IS DELIBERATELY NOT FLOORED AT THE TITLE, and this pins the
   * decision so it is not "fixed" again by reflex.
   *
   * `title / value-axis-title` is 205 of the family and is length-independent —
   * the same for two characters as for twenty-seven — because the `y` is
   * `Math.max(0, …)`, which parks the unit in the title's band on a chart whose
   * plot starts high. That looks exactly like the clamp bug this engine has
   * recorded five times, and flooring it was written and measured:
   *
   *     clip alone                1,014 pairs, no new shapes
   *     clip + floor at the ink   1,156 pairs, and `value-axis-title /
   *                               category#` at 310, a family that did not exist
   *
   * The floor buys the title collisions by moving the unit into the category
   * names on short charts. Same lesson the CAGR caption carries about the same
   * move: a clamp moves a label whether or not the destination is free.
   */
  it("tracks the plot's top edge, and is allowed into the title's band", () => {
    const withTitle = { valueAxisTitle: "€m", title: "A chart with a title" };
    const roomy = unitNode(base({ width: 300, height: 100, ...withTitle }));
    const squeezed = unitNode(base({ width: 300, height: 50, ...withTitle }));
    const title = buildChart(base({ width: 300, height: 50, ...withTitle })).nodes.find((n) => n.name === "title");
    expect(roomy && squeezed && title, "no unit or no title drawn").toBeTruthy();

    // It follows the plot up as the chart shortens — that is what makes it read
    // as a label ON the axis rather than a second heading.
    expect(squeezed!.y, "the unit stopped tracking the plot").toBeLessThan(roomy!.y);

    // And it is permitted to end up inside the title's band. THIS IS THE
    // REFUSED FIX: flooring it here is what costs 310 `value-axis-title /
    // category#`. If a change makes this fail, re-measure the whole sweep
    // (`DUMP_BUDGET=1 npx vitest run test/overlap-budget.test.ts`) before
    // deciding it is an improvement — the last time it was measured it was not.
    expect(squeezed!.y, "the unit was floored below the title — see the note above").toBeLessThan(
      (title as { y: number; h: number }).y + (title as { y: number; h: number }).h,
    );
  });
});
