import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig, CHART_KINDS } from "../src/core/samples";
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
 *
 * AND THEN 1,014 TO 785, by making the unit yield to the title — the second half
 * of this file. Width was never going to reach that part: `title /
 * value-axis-title` was identical at two characters and at twenty-seven.
 *
 * ONE NUMBER TO DISTRUST, and it is the one every figure above is quoted in. The
 * sweep drives this option with a twenty-seven-character sentence; at the
 * documented `€m` the same sweep totals 206 rather than 785, and the family is
 * 118 rather than 687. `test/overlap-budget.test.ts` carries the full warning.
 * Nothing here is wrong, but a stress figure is not a product figure.
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
   * THE UNIT YIELDS TO THE TITLE, and it is the LAST thing in that band to do so.
   *
   * `title / value-axis-title` was 205 of the family and length-independent —
   * identical for two characters and for twenty-seven. Both nodes are
   * `align: "left"` at `x: 0`, so their ink always shares the x range and only
   * the `y` decides; `Math.max(0, …)` then parks the unit inside the title's
   * band on a chart whose plot cannot start below both. **No width remedy can
   * move that number**, which is why the clip, the 2026-08-19 gutter-fit and the
   * shrink-to-fit all left it untouched.
   *
   * MOVING it was tried and refused: flooring the `y` at the title's ink costs
   * 310 `value-axis-title / category#` pairs that did not exist, because a clamp
   * moves a label whether or not the destination is free. So the unit is not
   * moved — it is DROPPED, on exactly the charts that cannot hold it.
   *
   * Measured over the 176 charts that draw a title and a unit: 22 overlap, and
   * the split is clean — every overlapping pair has an ink gap of -2.56pt or
   * worse, every clear one +3.31pt or better. All 22 are `80x60` and `300x60` at
   * 18pt, and on those charts the engine ALREADY drops the category names, the
   * axis strip and the legend ("Chrome yields to the title" in docs/MANUAL.md,
   * whose worked example is a 300x60 banner). The unit was the exception.
   *
   * And the old behaviour did not KEEP the author's text: it printed it through
   * the title, so the reader lost both.
   */
  it("yields to the title rather than printing through it", () => {
    // 300x60 at 18pt: the manual's own banner example, at a font where the band
    // genuinely cannot hold both. Nothing else in this band survives that chart
    // either — the category names and the axis strip are already gone.
    const banner = base({
      width: 300,
      height: 60,
      style: { fontSize: 18 },
      valueAxisTitle: "€m",
      title: "A chart with a title",
    });
    expect(unitNode(banner), "the unit still prints through the title on a banner").toBeUndefined();
    // The title is what stays — that is the rule being obeyed, and a version of
    // this that dropped the title instead would pass a bare "no unit" assertion.
    expect(
      buildChart(banner).nodes.find((n) => n.name === "title"),
      "dropped the title, which is the one label that must survive",
    ).toBeTruthy();
  });

  it("keeps the unit at every size anybody presents at", () => {
    /**
     * THE OBJECTION THAT SANK THE LAST TWO REMEDIES, pinned as a test.
     *
     * Drop-on-ink was refused on 2026-08-28 because "a clustered chart at
     * 480x300 in 18pt loses its unit, and that is a size people present at". The
     * first version of THIS rule did the same thing, for a reason worth keeping:
     * it compared `bandTop` against `titleHeight`, a BOX against a reserved
     * height, when the question is about the ink. Measuring the box where the
     * ink was meant is the mistake this repo has now made in five places.
     */
    for (const [w, h] of [
      [480, 300],
      [960, 540],
      [200, 150],
    ] as [number, number][])
      for (const fs of [10, 14, 18, 24]) {
        const n = unitNode(
          base({ width: w, height: h, style: { fontSize: fs }, valueAxisTitle: "€m", title: "Quarterly revenue" }),
        );
        expect(n, `lost the unit on a ${w}x${h} chart at ${fs}pt — a size people present at`).toBeTruthy();
      }
  });

  it("does not yield to a title that is not there", () => {
    // The clearance is measured against the TITLE's ink, so a chart with no
    // title has nothing to yield to and must keep its unit however short it is.
    // Without this, the rule reads `titleFontSize` — which answers whatever the
    // style says regardless of whether a title is drawn — and quietly drops the
    // unit from untitled charts, which is a strictly worse version of the bug it
    // was written to fix.
    for (const [w, h] of [
      [300, 60],
      [80, 60],
      [200, 150],
    ] as [number, number][]) {
      // EXPLICITLY untitled. `sampleConfig` carries a title of its own, so a
      // fixture that merely omits one still has "Sales by quarter" — which is
      // how the first version of this test failed against correct code.
      const n = unitNode(
        base({ width: w, height: h, style: { fontSize: 18 }, valueAxisTitle: "€m", title: undefined }),
      );
      expect(n, `an untitled ${w}x${h} chart lost its unit to a title it does not have`).toBeTruthy();
    }
  });

  it("drops the unit exactly where it would have overprinted, and nowhere else", () => {
    /**
     * THE CONTRACT, as a count — because the endpoints alone do not pin it.
     *
     * Tests naming one frame that keeps and one that drops leave the THRESHOLD
     * free: a rule that yields twice as eagerly passes both, quietly costing
     * units on charts that had room. This is the measurement the change was
     * argued from, so it is the thing to assert.
     *
     * Over the 176 charts here that draw both a title and a unit: 22 overlapped
     * before, all at 80x60 and 300x60 at 18pt, and the split was clean — every
     * overlapping pair at -2.56pt of ink gap or worse, every clear one at
     * +3.31pt or better. So the right rule drops exactly those 22.
     */
    const drawn: string[] = [];
    const dropped: string[] = [];
    // ONLY THE KINDS THAT DRAW A UNIT AT ALL. Fourteen of the twenty-five never
    // reach this block — a pie has no value axis — and counting them as losses
    // buries the real number under a constant. Both earlier attempts at this
    // measurement got it wrong that way.
    const drawsAUnit = CHART_KINDS.filter(({ kind }) => {
      try {
        return !!buildChart(base({ kind, width: 960, height: 540, valueAxisTitle: "€m" })).nodes.find(
          (n) => n.name === "value-axis-title",
        );
      } catch {
        return false;
      }
    });
    for (const { kind } of drawsAUnit)
      for (const [w, h] of [
        [80, 60],
        [120, 90],
        [160, 120],
        [200, 150],
        [300, 60],
        [60, 300],
        [480, 300],
        [960, 540],
      ] as [number, number][])
        for (const fs of [10, 18]) {
          const cfg = base({ kind, width: w, height: h, style: { fontSize: fs }, valueAxisTitle: "€m" });
          let scene;
          try {
            scene = buildChart(cfg);
          } catch {
            continue;
          }
          const title = scene.nodes.find((n) => n.name === "title");
          if (!title) continue;
          const unit = scene.nodes.find((n) => n.name === "value-axis-title") as TextNode | undefined;
          const where = `${kind} ${w}x${h}@${fs}`;
          if (unit) {
            drawn.push(where);
            // Nothing that IS drawn may overprint the title — the whole point.
            const t = title as TextNode;
            const gap = unit.y + unit.h - unit.fontSize * 0.25 - unit.fontSize * 0.8 - (t.y + t.fontSize * 1.21);
            expect(gap, `${where} draws its unit through the title`).toBeGreaterThanOrEqual(0);
          } else dropped.push(where);
        }
    // Every loss is one of the two frames the manual already strips to a title
    // and its bars. A threshold that reaches any further shows up here.
    const frames = [...new Set(dropped.map((d) => d.split(" ")[1]))].sort();
    expect(frames, "the unit is yielding on charts that had room for it").toEqual(["300x60@18", "80x60@18"].sort());
    expect(dropped.length, "the number of charts losing their unit moved").toBe(22);
    expect(drawn.length, "charts keeping their unit moved").toBe(154);
  });

  it("still tracks the plot's top edge where it survives", () => {
    // It follows the plot up as the chart shortens — that is what makes it read
    // as a label ON the axis rather than a second heading. Unchanged by the
    // yield rule, which only decides whether it is drawn at all.
    const roomy = unitNode(base({ width: 300, height: 200, valueAxisTitle: "€m" }));
    const squeezed = unitNode(base({ width: 300, height: 120, valueAxisTitle: "€m" }));
    expect(roomy && squeezed, "no unit drawn").toBeTruthy();
    expect(squeezed!.y, "the unit stopped tracking the plot").toBeLessThanOrEqual(roomy!.y);
  });
});
