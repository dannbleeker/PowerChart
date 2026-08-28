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

/**
 * The other axis of the same problem: several SERIES labelling one category.
 *
 * `neighbourFs` above bounds a line's labels against the ones beside it in the
 * SAME series. A combo with ten line series draws ten numbers per category, and
 * nothing compared them to each other — sideways they run along one row,
 * upright they stack in one column above the same mark. It was the whole of the
 * `combo-label` residue left in the variant sweep after the category fix.
 *
 * Measured across the frame sweep at both fonts, both orientations:
 *
 *     before   209 labels drawn   137 overlapping pairs
 *     after     95 labels drawn     0 overlapping pairs
 *
 * The everyday combo is untouched at 79 labels, before and after. Ten series
 * labelling one category at nearly the same value cannot all be labelled at any
 * size — their MARKS are on top of each other too, so a number beside them names
 * neither — and that is what the drop is saying.
 */
const tenSeries = (base: ChartConfig): ChartConfig => ({
  ...base,
  data: {
    ...base.data,
    series: Array.from({ length: 10 }, (_, i) => ({
      ...base.data.series[i % base.data.series.length],
      name: `Series ${i + 1}`,
      values: base.data.series[i % base.data.series.length].values.map((v) => (v == null ? v : v + i)),
    })),
  },
});

describe("a combo with more series than any category can label", () => {
  it("draws no point label over another, in either orientation", () => {
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18])
        for (const horizontal of [false, true]) {
          const cfg = tenSeries({ ...plain("combo", w, h, fontSize), horizontal } as ChartConfig);
          const labels = textsOf(cfg, /^combo-label-/);
          expect(pairsAmong(labels), `${w}x${h} fs=${fontSize} H=${horizontal}: ten series are colliding`).toBe(0);
        }
  });

  it("still labels what it can", () => {
    // The drop must not be total — a rule that answered "no labels" everywhere
    // would pass the overlap check and be useless. Nearly half survive.
    let drawn = 0;
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18])
        for (const horizontal of [false, true])
          drawn += textsOf(
            tenSeries({ ...plain("combo", w, h, fontSize), horizontal } as ChartConfig),
            /^combo-label-/,
          ).length;
    expect(drawn, "ten series lost every label").toBeGreaterThan(60);
    expect(drawn, "ten series lost nothing, so the rule is not engaging").toBeLessThan(150);
  });

  it("charges the everyday combo nothing at all", () => {
    // THE REGRESSION THAT WOULD MATTER, and it is exact rather than a bound: the
    // peer rule must charge the everyday combo NOTHING.
    //
    // It was 79 when this was written and is 77 now, and the two lost labels are
    // not this rule's doing: the combo's legend has since been corrected to name
    // its LINE as well as its columns, so it takes a row it did not take before
    // and the plot beneath it is slightly shorter. Re-based rather than loosened
    // to a bound — an exact number is what makes this test worth having, and the
    // commit that moves it should say why, as this one does.
    let drawn = 0;
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 18])
        for (const horizontal of [false, true])
          drawn += textsOf({ ...plain("combo", w, h, fontSize), horizontal } as ChartConfig, /^combo-label-/).length;
    expect(drawn, "the everyday combo paid for the ten-series case").toBe(77);
  });

  it("keeps every mark, whatever happens to the numbers", () => {
    // Dropping a LABEL loses a reading; the series themselves are the chart.
    const cfg = tenSeries(plain("combo", 480, 300, 10));
    const marks = buildChart(cfg).nodes.filter((n) => /^combo-(marker|line)-/.test(n.name ?? ""));
    expect(marks.length, "the overlay lost marks along with its labels").toBeGreaterThan(20);
  });
});

describe("what the peer fit measures, per orientation", () => {
  /**
   * Two series a hair apart, labelled with single digits, upright.
   *
   * THE MUTANT THIS EXISTS FOR. Upright the peer labels stack in one column, so
   * what has to fit between two marks is a LINE HEIGHT. Measuring their WIDTH
   * instead — which is right sideways, where they run along a row — passed every
   * other test in this file, because a wide number makes the width rule the
   * TIGHTER of the two and nothing was left overlapping. It is single digits
   * that separate them: "1" is about 0.6 em wide and 1.4 em tall, so a width
   * rule grants more than twice the size a line height allows, and the two
   * labels are drawn through each other.
   */
  const hairsBreadth = (fontSize: number, gap: number): ChartConfig =>
    ({
      kind: "combo",
      width: 480,
      height: 300,
      style: { fontSize },
      // SINGLE-DIGIT labels, which is what makes the two rules disagree: "9" is
      // about 0.6 em wide and 1.4 em tall. With the formatter left alone these
      // come out "8.0" and "9.0", three characters and 1.65 em, at which point
      // the width rule is the TIGHTER of the two and the mutant passes.
      numberFormat: { decimals: 0 },
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          { name: "Bars", values: [10, 20, 30, 40] },
          { name: "Up", type: "line", values: [8, 8, 8, 8] },
          { name: "Down", type: "line", values: [8 + gap, 8 + gap, 8 + gap, 8 + gap] },
        ],
      },
      decorations: { segmentLabels: true },
    }) as ChartConfig;

  it("upright, a line height — not the width of the digits", () => {
    // SWEPT over the value gap, not pinned to one. A single gap is hostage to
    // where the two constraints happen to cross: at 3% apart both rules drop
    // the label and the mutant survives, at 6% only the line-height rule does.
    // The property is "no gap between two series produces a collision".
    for (const fontSize of [10, 14, 18])
      for (const gap of [0.1, 0.25, 0.5, 1, 2, 4]) {
        const labels = textsOf(hairsBreadth(fontSize, gap), /^combo-label-/);
        expect(
          pairsAmong(labels),
          `fs=${fontSize} gap=${gap}: two near-identical series are drawn through each other`,
        ).toBe(0);
      }
  });

  it("EQUIVALENT MUTANTS, recorded rather than chased", () => {
    // TWO of them, and both were hunted before being written down.
    //
    // 1. UPRIGHT, USING THE LABEL WIDTH INSTEAD OF A LINE HEIGHT. `1.4` is the
    //    quantity that is true — the peer labels stack in one column, so what
    //    must fit between two marks is a line of text — but a width rule cannot
    //    be caught out. For every label wider than 1.4 em, which is everything
    //    from three characters up, the width rule is the STRICTER of the two.
    //    For single digits it is looser, and looser can only draw MORE; across
    //    eight frames, three fonts, six value gaps and both orientations it
    //    never draws one that collides, because the 5pt floor removes the
    //    marginal cases first. Kept as 1.4 because it is the right number, not
    //    because a test can tell.
    //
    // 2. THE SINGLE-SERIES GUARD, below.

    // `lines.length < 2` skips the peer pass entirely. Removing it changes no
    // answer and no test caught it — correctly, because with one series the
    // sorted list holds one entry, there is no next, and every label keeps
    // `pointFs`. It is a fast path for the common chart, not a rule, and it is
    // asserted as one: one line series must be untouched by the peer fit.
    const one: ChartConfig = {
      kind: "combo",
      width: 480,
      height: 300,
      style: { fontSize: 10 },
      data: {
        categories: ["A", "B", "C", "D"],
        series: [
          { name: "Bars", values: [10, 20, 30, 40] },
          { name: "Line", type: "line", values: [12, 22, 32, 42] },
        ],
      },
      decorations: { segmentLabels: true },
    } as ChartConfig;
    const labels = textsOf(one, /^combo-label-/);
    expect(labels.length, "a single line series lost labels to a rule about peers").toBe(4);
    for (const l of labels) expect(l.fontSize, "a single line series was shrunk by the peer fit").toBe(10);
  });
});

/**
 * The scatter's own x tick numbers, against each other.
 *
 * This strip HAD a fit — `gapScale` guarantees the tick spacing can carry the
 * widest label — and still drew numbers through each other, because the fit was
 * spent after it was granted. The label on the axis's origin is centred on its
 * tick, so half of it hangs into the y axis's gutter, and it is nudged right by
 * exactly the overhang. Right is toward its neighbour.
 *
 * The same defect as the secondary value axis, and it takes the other remedy.
 * There, one shift for the whole strip keeps every gap. Here only the FIRST
 * label moves, so shifting all of them by its overhang would push the last one
 * off the canvas — this strip pays in SIZE instead, shrinking until the nudged
 * layout clears and dropping the numbers if it cannot.
 *
 * The cost is three tick numbers in 731 across the sweep.
 */
describe("a scatter's x tick numbers pay for their own nudge", () => {
  const DATA: [string, (c: ChartConfig) => ChartConfig][] = [
    ["plain", (c) => c],
    [
      "billions",
      (c) => ({
        ...c,
        data: {
          ...c.data,
          series: c.data.series.map((s) => ({ ...s, values: s.values.map((v) => (v == null ? v : v * 1234567)) })),
        },
      }),
    ],
    [
      "tiny fractions",
      (c) => ({
        ...c,
        data: {
          ...c.data,
          series: c.data.series.map((s) => ({ ...s, values: s.values.map((v) => (v == null ? v : v / 100000)) })),
        },
      }),
    ],
  ];

  it("draws no tick number over another, at any frame or magnitude", () => {
    for (const kind of ["scatter", "bubble"] as const)
      for (const [dname, shape] of DATA)
        for (const [w, h] of FRAMES)
          for (const fontSize of [10, 18]) {
            const ticks = textsOf(shape(plain(kind, w, h, fontSize)), /^x-axis$/);
            expect(pairsAmong(ticks), `${kind} ${dname} ${w}x${h} fs=${fontSize}: x tick numbers collide`).toBe(0);
          }
  });

  it("pays in size, not by abandoning the strip", () => {
    // The shrink must engage without emptying the axis — a rule that answered
    // "no numbers" would pass the check above and tell the reader nothing.
    let drawn = 0;
    for (const kind of ["scatter", "bubble"] as const)
      for (const [, shape] of DATA)
        for (const [w, h] of FRAMES)
          for (const fontSize of [10, 18]) drawn += textsOf(shape(plain(kind, w, h, fontSize)), /^x-axis$/).length;
    expect(drawn, "the x strip was abandoned rather than shrunk").toBeGreaterThan(200);
  });

  it("drops the strip rather than shrinking it past reading", () => {
    // The shrink loop has to STOP. Without a floor it converges on a legible-
    // looking layout made of two-point numbers: the overlap check passes and the
    // reader gets ink. The same answer the pie ring, the radar names and the
    // secondary axis give — below the floor there is no strip.
    const bad: string[] = [];
    for (const kind of ["scatter", "bubble"] as const)
      for (const [dname, shape] of DATA)
        for (const [w, h] of FRAMES)
          for (const fontSize of [10, 18])
            for (const t of textsOf(shape(plain(kind, w, h, fontSize)), /^x-axis$/))
              if (t.fontSize < 5) bad.push(`${kind} ${dname} ${w}x${h} fs=${fontSize}: ${t.fontSize.toFixed(2)}pt`);
    expect(bad, "an x tick was drawn below the legibility floor").toEqual([]);
  });

  it("leaves a comfortable scatter's x strip at full size", () => {
    // THE REGRESSION THAT WOULD MATTER: the nudge only bites where the origin
    // label is wide relative to its gap, so an ordinary chart must not shrink.
    for (const [w, h] of [
      [480, 300],
      [960, 540],
    ]) {
      const ticks = textsOf(plain("scatter", w, h, 10), /^x-axis$/);
      expect(ticks.length, `${w}x${h}: the x strip vanished`).toBeGreaterThan(2);
      for (const t of ticks) expect(t.fontSize, `${w}x${h}: a roomy chart's x tick was shrunk`).toBeCloseTo(9, 5);
    }
  });
});
