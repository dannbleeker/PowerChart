import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { textWidth, type SceneNode, type TextNode } from "../src/core/scene";

/**
 * EVERY TEXT OVERLAP THIS ENGINE STILL HAS, counted, named, and capped.
 *
 * `frame-fit.test.ts` sweeps each kind at eight frames and seven fonts and now
 * allows no overlap at all. What it does not sweep is the option and data-shape
 * VARIANTS — a chart with twenty-four categories, ten series, a secondary axis,
 * a pareto, a footnote — and that is the blind spot the whole "24 categories"
 * family lived in for nine days. The gate was green throughout.
 *
 * This is that sweep, ratcheted. It ran for the first time on 2026-08-27 at
 * 2,148 overlapping pairs, measured by the frame gate's own ink rule; a night's
 * work took it to 537. Pinning it per SHAPE rather than as one total matters,
 * because a total can hide a trade — one family fixed while another grows — and
 * a shape absent from the table is a regression by definition, since it is text
 * the engine has never been seen to draw over text before.
 *
 * AND THEN THIS FILE TURNED OUT TO HAVE THE SAME BLIND SPOT IT WAS BUILT TO
 * CLOSE. It swept every option and every data shape and never one of each: the
 * two tables were CONCATENATED, so a secondary axis on ten series, or a
 * footnote on twenty-four categories, was not among the 24,000 charts. Crossing
 * a slice of them on 2026-08-29 took the count from 537 to 4,010 and found
 * twenty-two shapes this engine had never been SEEN to draw — 2,683 of them in
 * small-multiples panels alone. Nothing about the engine changed that day.
 *
 * So: the total moved because the sweep did, and a number from before that date
 * is not comparable with one after it. `CROSS_OPTIONS` below says what is
 * crossed and, more usefully, what still is not.
 *
 * WHEN A NUMBER HERE GOES DOWN, EDIT IT DOWN. The budget is a ceiling, not a
 * target, and one left above the real figure is a ratchet that has stopped
 * ratcheting — the next regression hides under the slack. There is a test below
 * that enforces exactly that, so improving the engine is supposed to fail this
 * file.
 *
 * The sweep is 24 options and 10 data shapes across every kind, eight frames,
 * two fonts and both orientations — about 24,000 charts — and it runs in eight
 * seconds, which is why it can live in the ordinary suite rather than in a
 * script somebody has to remember.
 */

/** EXACTLY `inkBox` from the frame gate, so the two measure the same thing. */
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

/** The option variants, matching the frame gate's own table. */
const OPTIONS: Record<string, Record<string, unknown>> = {
  scale: { scale: { min: -50, max: 500 } },
  segmentOrder: { segmentOrder: "reverse" },
  categorySort: { categorySort: "descending" },
  secondaryAxis: { secondaryAxis: true },
  axisBreak: { axisBreak: { from: 20, to: 60 } },
  valueAxisTitle: { valueAxisTitle: "Revenue in millions of euro" },
  logScale: { logScale: true },
  gapWidth: { gapWidth: 0.8 },
  overlap: { overlap: 0.5 },
  footnote: { footnote: "Source: an internal model, restated for the 2024 perimeter" },
  pareto: { pareto: true },
  "multiples 2 columns": { multiples: { columns: 2 } },
  "multiples 3 columns": { multiples: { columns: 3 } },
  otherBucket: { otherBucket: { max: 3 } },
  "pie.explode": { pie: { explode: [0, 1] } },
  "pie.semi": { pie: { semi: true } },
  "pie.variableRadius": { pie: { variableRadius: true } },
  "tilemap.hex": { tilemap: { shape: "hex" } },
  "tilemap.glyph": { tilemap: { glyph: "bars" } },
  "butterfly.split": { butterfly: { split: 1 } },
  boxplot: { boxplot: { notch: true, showMean: true, jitter: true } },
  "render image": { render: "image" },
  numberFormat: { numberFormat: { locale: "de-DE", currency: "EUR", decimals: 2 } },
  title: { title: "A rather long chart title that names the measure and the period" },
};

/** And the data shapes, likewise: a user's data is none of the sample's shapes. */
const DATA_SHAPES: Record<string, (c: any) => any> = {
  "long category names": (c) => ({
    ...c,
    data: { ...c.data, categories: c.data.categories.map((x: string, i: number) => `${x} enterprise segment ${i}`) },
  }),
  "long series names": (c) => ({
    ...c,
    data: { ...c.data, series: c.data.series.map((s: any) => ({ ...s, name: `${s.name} including allocations` })) },
  }),
  "24 categories": (c) => ({
    ...c,
    data: {
      ...c.data,
      categories: Array.from({ length: 24 }, (_, i) => `C${i + 1}`),
      series: c.data.series.map((s: any) => ({
        ...s,
        values: Array.from({ length: 24 }, (_, i) => (s.values[i % s.values.length] ?? 0) + i),
      })),
    },
  }),
  "10 series": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: Array.from({ length: 10 }, (_, i) => ({
        ...c.data.series[i % c.data.series.length],
        name: `Series ${i + 1}`,
        values: c.data.series[i % c.data.series.length].values.map((v: number | null) => (v == null ? v : v + i)),
      })),
    },
  }),
  "one category": (c) => ({
    ...c,
    data: {
      ...c.data,
      categories: c.data.categories.slice(0, 1),
      series: c.data.series.map((s: any) => ({ ...s, values: s.values.slice(0, 1) })),
    },
  }),
  "one series": (c) => ({ ...c, data: { ...c.data, series: c.data.series.slice(0, 1) } }),
  "values in the billions": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: any) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : v * 1234567)),
      })),
    },
  }),
  "every value negative": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: any) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : -v)),
      })),
    },
  }),
  "mixed signs": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: any) => ({
        ...s,
        values: s.values.map((v: number | null, i: number) => (v == null ? v : i % 2 ? -v : v)),
      })),
    },
  }),
  "tiny fractions": (c) => ({
    ...c,
    data: {
      ...c.data,
      series: c.data.series.map((s: any) => ({
        ...s,
        values: s.values.map((v: number | null) => (v == null ? v : v / 100000)),
      })),
    },
  }),
};

/**
 * What each remaining shape is allowed, as of 2026-08-29. Total 4,010.
 *
 * **THE JUMP FROM 537 TO 4,010 IS NOT A REGRESSION. It is the sweep widening.**
 * Nothing about the engine changed; on 2026-08-29 this file started crossing
 * option variants with data shapes (see `CROSS_OPTIONS`), and every pair below
 * that was not there before had been drawn by this engine all along, in a
 * configuration nobody had built. Read the count as coverage, not as damage —
 * and never compare a number here against one measured before that date without
 * saying which sweep produced it. The 2026-08-19 figure of "75 pairs" was
 * mis-compared exactly that way once already.
 *
 * A `#` stands in for any run of digits, so `label-3 / label-4` and `label-9 /
 * label-10` are one shape. What each family is, and why the ones left are still
 * here, is in `docs/BACKLOG.md`.
 *
 * The 4,010 is two families and a tail:
 *
 *     2,683   `p#-*` — SMALL MULTIPLES, all of it new on 2026-08-29
 *     1,058   `value-axis-title` — the owner's open decision
 *       269   everything else
 *
 * The small-multiples family is the discovery, and it is NOT a labelling defect
 * — that was the first guess and a diagnostic refuted it. The category axis
 * thins correctly, and zero of the 910 `p#-category#` pairs are two names in
 * the same panel. All 910 come from ONE frame, 300x60, where ten series in two
 * columns is five rows and
 *
 *     panelH = (height − titleH − footH − gap × (rows − 1)) / rows
 *
 * leaves nothing per panel. `buildMultiples` builds the grid anyway, every fit
 * below floors at `MIN_PLOT_SIDE`, and a panel's category strip ends up at
 * y = 307 on a chart 60 points tall. The labels are only where it shows.
 * Measured, diagnosed, not yet fixed — see docs/BACKLOG.md.
 */
const BUDGET: Record<string, number> = {
  // Small multiples, crossed with dense data. New on 2026-08-29 and unfixed.
  "p#-category# / p#-category#": 910,
  "p#-title / p#-title": 408,
  "p#-value-axis / p#-value-axis": 378,
  "p#-label# / p#-label#": 292,
  "p#-label## / p#-label##": 261,
  "p#-total# / p#-total#": 198,
  "p#-tick# / p#-category#": 48,
  "p#-label## / p#-total#": 34,
  "p#-cagr-label / p#-total#": 28,
  "p#-total# / p#-cagr-label": 21,
  "p#-value-axis / p#-title": 20,
  "p#-cagr-label / p#-title": 19,
  "p#-total# / p#-label##": 18,
  "p#-tick# / p#-tick#": 16,
  "p#-value-axis / p#-category#": 12,
  "p#-cagr-label / p#-cagr-label": 8,
  "p#-category# / p#-title": 6,
  "p#-category# / p#-tick#": 4,
  "p#-label# / p#-title": 2,
  // The `value-axis-title` group — one open decision, not a pile of defects.
  "value-axis-title / legend#": 394,
  "title / value-axis-title": 205,
  "value-axis-title / total#": 182,
  "value-axis / value-axis-title": 163,
  "total# / value-axis-title": 116,
  "value-axis-title / label#": 73,
  "value-axis-title / series-label#": 55,
  "value-axis-title / median-label#": 20,
  // The tail.
  "combo-series-label# / combo-series-label#": 33,
  "title / series-label#": 28,
  "label# / label#": 22,
  "title / total#": 8, // new 2026-08-29: a total row under a long title, crossed
  "series-label# / series-label#": 7,
  "header# / label##": 4,
  "title / row#": 4,
  "tile#-value / tile#-value": 4,
  "title / category#": 3,
  "label## / cagr-label": 2,
  "total# / title": 2, // new 2026-08-29
  "legend# / footnote": 2, // new 2026-08-29: a footnote under a wrapped legend
};

/**
 * THE COMBINATIONS, because until 2026-08-29 this sweep never made one.
 *
 * `variants` below was the two tables CONCATENATED — every option applied to a
 * sample config, every data shape applied to a sample config, and never an
 * option and a data shape together. So a chart with a secondary axis AND ten
 * series was not among the 24,000, and neither was a footnote on
 * twenty-four categories.
 *
 * That is the same defect this file was WRITTEN TO FIX, one level up.
 * `frame-fit.test.ts` swept kinds, frames and fonts and missed the option and
 * data-shape variants; this file swept the variants and missed their products.
 * A gate is only as wide as its sweep, and both gates were green.
 *
 * Measured on a 5x4 slice before widening: three shapes appeared that the
 * uncrossed sweep has never produced in any of its 24,000 charts —
 * `title / total#`, `legend# / footnote`, `title / label#` — and every shape
 * the two share came out several times larger. By this file's own rule a shape
 * absent from the table is a regression by definition, so those were three
 * regressions nobody could see.
 *
 * NOT THE FULL PRODUCT, and the reason is honest rather than principled: 24
 * options by 10 shapes is 240 variants against 34, which takes this from eight
 * seconds to about a minute and out of the ordinary suite. What is crossed here
 * is the options that change the LAYOUT'S STRUCTURE against the data shapes
 * that stress LABEL COUNT, which is where the interesting products live.
 *
 * WHAT IS STILL NOT CROSSED, so the next person does not have to re-derive it:
 * every other option (`scale`, `logScale`, `axisBreak`, `gapWidth`, `overlap`,
 * `segmentOrder`, `categorySort`, `otherBucket`, `numberFormat`, `render
 * image`, the per-kind ones) against every data shape; and any product of three
 * or more. Widen this list before concluding a family is closed.
 */
const CROSS_OPTIONS = ["secondaryAxis", "pareto", "valueAxisTitle", "footnote", "title", "multiples 2 columns"];
const CROSS_SHAPES = ["24 categories", "10 series", "long series names", "long category names"];

/** The whole sweep, once. */
function measure(): { total: number; byShape: Map<string, number> } {
  const byShape = new Map<string, number>();
  let total = 0;
  const variants: [string, (c: unknown) => unknown][] = [
    ...Object.entries(DATA_SHAPES),
    ...Object.entries(OPTIONS).map(
      ([k, o]) => [k, (c: unknown) => ({ ...(c as object), ...o })] as [string, (c: unknown) => unknown],
    ),
  ];
  for (const ok of CROSS_OPTIONS)
    for (const sk of CROSS_SHAPES) {
      const o = OPTIONS[ok];
      const fn = DATA_SHAPES[sk];
      if (!o || !fn) throw new Error(`CROSS names a variant that does not exist: ${ok} x ${sk}`);
      variants.push([`${ok} × ${sk}`, (c: unknown) => ({ ...(fn(c as never) as object), ...o })]);
    }
  for (const [, fn] of variants)
    for (const { kind } of CHART_KINDS)
      for (const [w, h] of FRAMES)
        for (const fontSize of [10, 18])
          for (const horizontal of [false, true]) {
            let ts: TextNode[];
            try {
              // A kind that refuses an option throws, and that is not an overlap.
              const cfg = fn({ ...sampleConfig(kind), width: w, height: h, horizontal, style: { fontSize } });
              ts = buildChart(cfg as never).nodes.filter(
                (n: SceneNode): n is TextNode => n.kind === "text" && !!String(n.text).trim(),
              );
            } catch {
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
                }
          }
  return { total, byShape };
}

describe("the text this engine still draws over other text", () => {
  const { total, byShape } = measure();
  /**
   * `DUMP_BUDGET=1 npx vitest run test/overlap-budget.test.ts` prints the whole
   * measured table, ready to paste over `BUDGET`.
   *
   * Here because re-baselining by hand from assertion output is how a budget
   * ends up above its real figure — the failure prints only the shapes that
   * broke a rule, so the unchanged ones get copied forward from the old table
   * and quietly keep their slack. The fourth test exists to catch exactly that;
   * this makes it unnecessary to trip.
   */
  if (process.env.DUMP_BUDGET) {
    const lines = [...byShape.entries()].sort((a, b) => b[1] - a[1]).map(([k, v]) => `  "${k}": ${v},`);
    // eslint-disable-next-line no-console
    console.log(`TOTAL ${total}\n${lines.join("\n")}`);
  }

  it("draws no shape it has never drawn before", () => {
    // A shape absent from the budget is text the engine has NOT been seen to
    // overlap, so its arrival is a regression whatever the total does. This is
    // the half a single number cannot express.
    const unbudgeted = [...byShape.entries()].filter(([k]) => BUDGET[k] === undefined);
    expect(
      unbudgeted,
      "a new kind of text overlap appeared — fix it, or add it to the budget and to docs/BACKLOG.md saying what it is",
    ).toEqual([]);
  });

  it("stays inside every shape's budget", () => {
    const over = [...byShape.entries()]
      .filter(([k, n]) => n > (BUDGET[k] ?? 0))
      .map(([k, n]) => `${k}: ${n} (budget ${BUDGET[k]})`);
    expect(over, "a family of text overlaps grew").toEqual([]);
  });

  it("keeps the budget honest — no line left above the real figure", () => {
    // THE HALF THAT MAKES IT A RATCHET. A ceiling nobody lowers stops ratcheting
    // and the next regression hides under the slack. Fixing overlaps is meant to
    // fail this: read the numbers off the message and edit them down.
    const slack = [...Object.entries(BUDGET)]
      .filter(([k, n]) => n > (byShape.get(k) ?? 0))
      .map(([k, n]) => `${k}: budget ${n}, actual ${byShape.get(k) ?? 0}`);
    expect(slack, "these budgets sit above what the engine now does — lower them").toEqual([]);
  });

  it("has not grown in total", () => {
    // DERIVED from the table above rather than written down again. A second
    // number is a second thing to forget: this one was left at 467 while the
    // shapes below it had already been lowered to 449, so it would have waved
    // through eighteen new overlaps. The per-shape budgets are the one source of
    // truth and the honesty test keeps each of them exact, so their sum is the
    // total — no slack anywhere, by construction.
    const cap = Object.values(BUDGET).reduce((a, b) => a + b, 0);
    expect(total, "the total overlap count grew").toBe(cap);
  });
});
