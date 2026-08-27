import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { textWidth, type SceneNode, type TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * The last of the "24 categories" overlaps, and they were two problems wearing
 * one name.
 *
 * The BACKLOG line that recorded this family said "pie/doughnut adjacent outside
 * labels, radar category names". Re-measuring after both of those were fixed
 * showed the heading had outlived its description: what was left under it was a
 * combo's point labels running into each other, and a scatter's or bubble's
 * point labels printed across the axis tick numbers.
 *
 * Measured across the frame sweep's eight sizes at both sweep fonts:
 *
 *     combo    225 overlapping pairs -> 0     301 point labels drawn -> 156
 *     scatter   72 overlapping pairs -> 0     point labels unchanged
 *     bubble    68 overlapping pairs -> 0     point labels unchanged
 *
 * Both fixes follow rules this engine had already written down. The combo's is
 * the pie's and the radar's: shrink to the room the neighbour leaves, drop below
 * the legibility floor, move nothing. The scatter's is its own recorded verdict
 * — "a point's label is DATA and a tick number is chrome, so the chrome yields"
 * — which was decided, written in a comment, and never actually carried out.
 */

/** EXACTLY the frame gate's `inkBox` for text, so this measures what it does. */
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

/** Shared area with more than a square point, the gate's own threshold. */
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

const textsOf = (cfg: ChartConfig, re: RegExp): TextNode[] =>
  buildChart(cfg).nodes.filter(
    (n: SceneNode): n is TextNode => n.kind === "text" && re.test(n.name ?? "") && !!String(n.text).trim(),
  );

/** The sample for a kind, stretched to 24 categories — the sweep's own shape. */
const wide = (kind: Parameters<typeof sampleConfig>[0], w: number, h: number, fontSize: number): ChartConfig => {
  const base = sampleConfig(kind) as ChartConfig;
  return {
    ...base,
    width: w,
    height: h,
    style: { ...base.style, fontSize },
    data: {
      ...base.data,
      categories: Array.from({ length: 24 }, (_, i) => `C${i + 1}`),
      series: base.data.series.map((s) => ({
        ...s,
        values: Array.from({ length: 24 }, (_, i) => (s.values[i % s.values.length] ?? 0) + i),
      })),
    },
  };
};

const plain = (kind: Parameters<typeof sampleConfig>[0], w: number, h: number, fontSize: number): ChartConfig => {
  const base = sampleConfig(kind) as ChartConfig;
  return { ...base, width: w, height: h, style: { ...base.style, fontSize } };
};

function pairsAmong(labels: TextNode[]): number {
  let n = 0;
  for (let i = 0; i < labels.length; i++)
    for (let j = i + 1; j < labels.length; j++) if (collides(labels[i], labels[j])) n++;
  return n;
}

describe("a combo line with a label on every point", () => {
  it("draws no point label over another, at any frame the sweep uses", () => {
    // Frame by frame, not in total: one number hides WHICH size regressed.
    for (const [w, h] of FRAMES) {
      for (const fontSize of [10, 18]) {
        const labels = textsOf(wide("combo", w, h, fontSize), /^combo-label-/);
        expect(pairsAmong(labels), `${w}x${h} fs=${fontSize}: point labels are colliding`).toBe(0);
      }
    }
  });

  it("shrinks to the category pitch before it drops anything", () => {
    // SHRINK-then-drop, not drop alone. 480x300 at 18pt is the case that shows
    // both halves: twenty-four labels asking for 18pt on a plot with 17.5pt of
    // pitch, and all twenty-four still drawn — at 14.8. A rule that only dropped
    // would have taken labels away here; a rule that only shrank would print
    // unreadable ink on the smaller frames the sweep also covers.
    const all = textsOf(wide("combo", 480, 300, 18), /^combo-label-/);
    expect(all.length, "labels were dropped on a frame that had room to shrink into").toBe(24);
    const sizes = all.map((l) => l.fontSize);
    expect(Math.min(...sizes), "nothing shrank, so the fit is not engaging").toBeLessThan(18);
    expect(Math.min(...sizes), "shrank past the legibility floor").toBeGreaterThanOrEqual(5);
    // And a frame with genuine room is not charged for it.
    const roomy = textsOf(wide("combo", 960, 540, 18), /^combo-label-/);
    expect(roomy.length, "a 960x540 combo lost labels").toBe(24);
    expect(Math.min(...roomy.map((l) => l.fontSize)), "a 960x540 combo was shrunk for nothing").toBe(18);
  });

  it("leaves the everyday combo alone", () => {
    // THE REGRESSION THAT WOULD MATTER: paying for the 24-category case out of
    // the six-category one. The sample keeps every label it drew before, and
    // four MORE across the sweep — smaller labels collide with the column totals
    // less often, so the drop pass that runs after this takes fewer.
    let drawn = 0;
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18]) drawn += textsOf(plain("combo", w, h, fontSize), /^combo-label-/).length;
    expect(drawn, "the everyday combo lost labels to the crowding rule").toBeGreaterThanOrEqual(37);
  });

  it("measures against the nearest DRAWN neighbour, not the next category", () => {
    // A line skips its gaps: a null value draws no mark and no label, so "the
    // label beside this one" is not `c - 1`. Taking the category index instead
    // shrinks a label to clear a number that is not on the chart.
    //
    // THE GAP HAS TO BE MORE THAN TWO WIDE for this to discriminate, which cost
    // two surviving mutants to learn. Skip exactly one category and the wrong
    // version computes the SAME answer by arithmetic accident: it halves the gap
    // and drops the absent neighbour"s width, and the two errors cancel exactly.
    // With two categories skipped they no longer cancel — the true gap is three
    // pitches and the mistaken one is still a single pitch.
    const gappy = (nulls: boolean): ChartConfig => ({
      kind: "combo",
      width: 480,
      height: 300,
      // 32pt so the chart font is not itself the binding cap: with a 10pt ceiling
      // both the right answer and the wrong one clamp to 10 and the comparison
      // says nothing.
      style: { fontSize: 32 },
      data: {
        categories: Array.from({ length: 12 }, (_, i) => `C${i + 1}`),
        series: [
          { name: "Bars", values: Array.from({ length: 12 }, () => 10) },
          {
            name: "Line",
            type: "line",
            // Every third index is drawn; the two between are seven digits wide
            // in the dense case and ABSENT in the gapped one, so the nearest drawn
            // neighbour is three categories away rather than one.
            values: Array.from({ length: 12 }, (_, i) => (i % 3 !== 0 && nulls ? null : 1_000_000 + i)),
          },
        ],
      },
      decorations: { segmentLabels: true },
    });
    const dense = textsOf(gappy(false), /^combo-label-/);
    const sparse = textsOf(gappy(true), /^combo-label-/);
    const seven = (ls: TextNode[]) => ls.filter((l) => String(l.text).replace(/\D/g, "").length >= 7);
    expect(seven(dense).length, "no seven-digit label survived the dense case").toBeGreaterThan(0);
    expect(seven(sparse).length, "no seven-digit label survived the gapped case").toBeGreaterThan(0);
    // THE RATIO, not merely "larger", and that is the whole discrimination. The
    // drawn points are three pitches apart where the dense line's are one, and
    // the size is linear in the gap, so the right answer is three times the
    // dense one — 20.2 against 6.7. Reading `c - 1` instead gives a single
    // pitch and, because the absent neighbour contributes no width, TWO times.
    // "Larger" passes for both. 2.5x separates them, and sits clear of either.
    const biggest = (ls: TextNode[]) => Math.max(...seven(ls).map((l) => l.fontSize));
    expect(biggest(sparse) / biggest(dense), "the gap between drawn points was only half-counted").toBeGreaterThan(2.5);
  });
});

describe("a scatter's point labels over its axis numbers", () => {
  for (const kind of ["scatter", "bubble"] as const) {
    it(`${kind}: no tick number is left under a point label`, () => {
      for (const [w, h] of FRAMES) {
        for (const fontSize of [10, 18]) {
          for (const cfg of [plain(kind, w, h, fontSize), wide(kind, w, h, fontSize)]) {
            const ticks = textsOf(cfg, /^[xy]-axis$/);
            const points = textsOf(cfg, /^label-\d+$/);
            const bad = ticks.filter((t) => points.some((p) => collides(t, p)));
            expect(
              bad.map((t) => t.text),
              `${kind} ${w}x${h} fs=${fontSize}: a tick number is buried`,
            ).toEqual([]);
          }
        }
      }
    });
  }

  it("keeps the point labels — the chrome yields, not the data", () => {
    // The whole basis of the trade. If dropping tick numbers cost point labels
    // as well, both halves of the decision would have been paid for and only one
    // delivered. Measured across seven fonts: 266 point labels become 265 on a
    // scatter and 265 become 267 on a bubble, because dodging a tick can take a
    // slot a later label wanted and can equally free one. Two either way.
    for (const [kind, floor] of [
      ["scatter", 260],
      ["bubble", 260],
    ] as const) {
      let drawn = 0;
      for (const [w, h] of FRAMES)
        for (const fs of [6, 8, 10, 14, 18, 24, 32]) drawn += textsOf(plain(kind, w, h, fs), /^label-\d+$/).length;
      expect(drawn, `${kind} lost point labels to the tick pass`).toBeGreaterThan(floor);
    }
  });

  it("dodges a tick before overwriting it", () => {
    // Without the dodging pass the drop is far more expensive: these two frames
    // kept 2 of 10 and 5 of 11 tick numbers. Placing what can avoid a tick
    // first, and only then placing the rest, is what buys them back.
    expect(
      textsOf(plain("scatter", 200, 150, 18), /^[xy]-axis$/).length,
      "200x150 lost most of its axis",
    ).toBeGreaterThanOrEqual(5);
    expect(
      textsOf(plain("scatter", 120, 90, 10), /^[xy]-axis$/).length,
      "120x90 lost most of its axis",
    ).toBeGreaterThanOrEqual(10);
  });

  it("leaves a comfortable chart's axis untouched", () => {
    // THE REGRESSION THAT WOULD MATTER. At the sizes anyone actually presents,
    // no point label sits on a tick number, so nothing should be dropped at all.
    for (const kind of ["scatter", "bubble"] as const) {
      for (const [w, h] of [
        [480, 300],
        [960, 540],
      ]) {
        expect(textsOf(plain(kind, w, h, 10), /^[xy]-axis$/).length, `${kind} ${w}x${h} lost a tick number`).toBe(11);
      }
    }
  });

  it("keeps the gridlines whatever happens to the numbers", () => {
    // A dropped number loses a reading; a dropped gridline would lose the axis.
    // The structure stays even where every number on it is covered.
    const nodes = buildChart(plain("scatter", 160, 120, 32)).nodes;
    expect(nodes.some((n) => n.kind === "line" && /^gridline-/.test(n.name ?? ""))).toBe(true);
  });
});
