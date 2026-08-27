import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { textWidth, type SceneNode, type TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * The semi-circle gauge: its slice labels, and the big total in the middle.
 *
 * `layoutGauge` leaves `layoutPie` on its first line, so none of the pie's own
 * label work ever reached it. The variant sweep counted 46 pairs of `label` on
 * `label` — the pie's outside-label crowding, unfixed here — and 58 of `label`
 * on `gauge-total`, which its own note said could not happen.
 *
 * That note is worth reading, because it was half right: "the centre is empty by
 * construction so there is nothing for it to land on". True of the CENTRE. Not
 * true of the box, which is clamped to stay inside the chart and on a short
 * gauge is carried down into the band the slice labels use — and not true of the
 * TEXT either, which was drawn at `fs * 1.7` inside a box `r * 2` wide and so
 * spilled off both sides of a small arc.
 *
 * Measured across the frame sweep at both fonts and both orientations:
 *
 *     before   128 labels drawn   42 label/label   58 label/total
 *     after     88 labels drawn    0 label/label    0 label/total
 */
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
const collides = (a: TextNode, b: TextNode) => {
  const p = ink(a);
  const q = ink(b);
  const w = Math.min(p.x1, q.x1) - Math.max(p.x0, q.x0);
  const h = Math.min(p.y1, q.y1) - Math.max(p.y0, q.y0);
  return w > 0 && h > 0 && w * h > 1;
};
const FRAMES: [number, number][] = [
  [80, 60],
  [120, 90],
  [160, 120],
  [200, 150],
  [300, 60],
  [60, 300],
  [480, 300],
  [960, 540],
];
const gauge = (w: number, h: number, fontSize: number): ChartConfig =>
  ({ ...sampleConfig("doughnut"), pie: { semi: true }, width: w, height: h, style: { fontSize } }) as ChartConfig;
const textsOf = (cfg: ChartConfig, re: RegExp): TextNode[] =>
  buildChart(cfg).nodes.filter(
    (n: SceneNode): n is TextNode => n.kind === "text" && re.test(n.name ?? "") && !!String(n.text).trim(),
  );

describe("a gauge's labels and its total", () => {
  it("draws no slice label over another", () => {
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18]) {
        const labels = textsOf(gauge(w, h, fontSize), /^label-/);
        let pairs = 0;
        for (let i = 0; i < labels.length; i++)
          for (let j = i + 1; j < labels.length; j++) if (collides(labels[i], labels[j])) pairs++;
        expect(pairs, `${w}x${h} fs=${fontSize}: gauge labels collide`).toBe(0);
      }
  });

  it("never prints a slice label through the total", () => {
    // The total wins, which is the verdict its own note already reached for a
    // different reason: it is the number the gauge exists to show. The leader
    // line is still drawn and still points at the wedge, so what is lost is a
    // name, not the identification.
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18]) {
        const cfg = gauge(w, h, fontSize);
        const labels = textsOf(cfg, /^label-/);
        const total = textsOf(cfg, /^gauge-total$/)[0];
        if (!total) continue;
        const bad = labels.filter((l) => collides(l, total));
        expect(
          bad.map((l) => l.text),
          `${w}x${h} fs=${fontSize}: a label is printed through the total`,
        ).toEqual([]);
      }
  });

  it("fits the total to the arc instead of spilling out of its own box", () => {
    // THE DEFECT A BOX-TO-BOX CHECK COULD NOT SEE. The box is `r * 2` wide and
    // the text was always `fs * 1.7`, so at 160x120 and 18pt a 30.6pt number sat
    // in a 40pt box — centred, it reached 22pt to the LEFT of where the box
    // starts, onto the slice label there.
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18]) {
        const total = textsOf(gauge(w, h, fontSize), /^gauge-total$/)[0];
        if (!total) continue;
        // UNLESS THE FLOOR IS HOLDING IT. Below 5pt there is nothing sensible left
        // to shrink to, and a gauge with no total is not a gauge — so at the floor
        // the number may still be a shade wider than the arc. At 80x60 and 18pt
        // that is 40.8pt in a 40pt box, which is the trade, not a miss.
        if (total.fontSize > 5)
          expect(
            textWidth(total.text, total.fontSize, total.bold),
            `${w}x${h} fs=${fontSize}: the total is wider than the box it is centred in`,
          ).toBeLessThanOrEqual(total.w);
      }
  });

  it("keeps the total at full size where the arc can hold it", () => {
    // THE REGRESSION THAT WOULD MATTER: the total is the headline, and shrinking
    // it on a chart with room would be paying for the small case out of the
    // ordinary one.
    for (const [w, h] of [
      [480, 300],
      [960, 540],
    ]) {
      const total = textsOf(gauge(w, h, 10), /^gauge-total$/)[0];
      expect(total, `${w}x${h}: the gauge lost its total`).toBeTruthy();
      expect(total.fontSize, `${w}x${h}: a roomy gauge's total was shrunk`).toBeCloseTo(17, 5);
    }
  });

  it("never leaves the gauge without a total", () => {
    // A gauge with no total is not a gauge. The fit shrinks and the floor wins;
    // it does not drop, which is where it differs from every label rule beside
    // it — those name a wedge, this IS the chart.
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18]) {
        const total = textsOf(gauge(w, h, fontSize), /^gauge-total$/)[0];
        expect(total, `${w}x${h} fs=${fontSize}: the gauge has no total at all`).toBeTruthy();
        expect(total.fontSize, `${w}x${h} fs=${fontSize}: the total is below the floor`).toBeGreaterThanOrEqual(5);
      }
  });

  it("holds the total at the floor rather than shrinking to ink", () => {
    // THE FLOOR ONLY BINDS ON A BIG NUMBER, which is why a mutant that removed
    // it survived the sweep above: the sample gauge totals a few thousand, and
    // `(2r - 2) / em` never falls near 5 for a five-character number. A total in
    // the billions with a currency and two decimals is eighteen characters, and
    // on a 20pt-radius arc the fit asks for less than two points.
    const big: ChartConfig = {
      kind: "doughnut",
      pie: { semi: true },
      width: 80,
      height: 60,
      style: { fontSize: 18 },
      numberFormat: { locale: "de-DE", currency: "EUR", decimals: 2 },
      data: {
        categories: ["A", "B", "C"],
        series: [{ name: "S", values: [1234567890, 2234567890, 3234567890] }],
      },
    } as ChartConfig;
    const total = textsOf(big, /^gauge-total$/)[0];
    expect(total, "a gauge with a huge total lost it entirely").toBeTruthy();
    expect(total.fontSize, "the total shrank past the point of being text").toBeGreaterThanOrEqual(5);
  });

  it("still labels the slices it has room for", () => {
    // The drop must not be total — a rule that answered "no labels" everywhere
    // would pass every check above and leave a gauge nobody can read.
    let drawn = 0;
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18]) drawn += textsOf(gauge(w, h, fontSize), /^label-/).length;
    expect(drawn, "the gauge lost every slice label").toBeGreaterThan(30);
  });

  it("keeps every wedge and every leader", () => {
    // Dropping a NAME loses a name. The wedge is the data and the leader is what
    // says which wedge a surviving name belongs to.
    const nodes = buildChart(gauge(480, 300, 10)).nodes;
    expect(nodes.filter((n) => /^slice-/.test(n.name ?? "")).length).toBeGreaterThan(2);
    expect(nodes.filter((n) => /^leader-/.test(n.name ?? "")).length).toBeGreaterThan(2);
  });
});
