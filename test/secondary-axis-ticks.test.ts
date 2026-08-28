import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { readFileSync } from "node:fs";
import { textWidth, type SceneNode, type TextNode } from "../src/core/scene";
import { tickGapScale } from "../src/core/layout/frame";
import type { ChartConfig } from "../src/core/types";

/**
 * The secondary value axis had no fit of any kind, and it was the largest
 * remaining overlap in this project.
 *
 * Every combo, every pareto and every dual-axis column draws this strip: the
 * cumulative or right-hand scale, five nice ticks, each label placed at its own
 * tick position and none of them measured against the next. On a short plot they
 * were simply drawn through each other. A sweep of the option and data-shape
 * variants counted 617 pairs of `secondary-axis` on `secondary-axis` and 320
 * more of it printed across the category names — 44% of every text overlap the
 * engine had left.
 *
 * The fix is not new code. `tickGapScale` already existed, a few hundred lines
 * away in `layoutScatter`, doing exactly this for the scatter's own two axes; it
 * moved to `frame.ts` and the secondary axis calls it. Nothing about the
 * scatter's output changed, which the frame gate checks.
 *
 * THE TWO ORIENTATIONS NEED DIFFERENT DIMENSIONS, and that is the whole reason
 * the helper takes a `want`. Sideways the labels sit at one y and spread along
 * x, so what one needs is its WIDTH; upright they stack down the right edge, so
 * it is a line height. Measuring the wrong one is how a scatter's x axis came to
 * be drawn through itself, and the same mistake was available here.
 */

/** EXACTLY the frame gate's `inkBox` for text. */
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

const textsOf = (cfg: ChartConfig, re: RegExp): TextNode[] =>
  buildChart(cfg).nodes.filter(
    (n: SceneNode): n is TextNode => n.kind === "text" && re.test(n.name ?? "") && !!String(n.text).trim(),
  );

/** Every kind, under the two options that raise a secondary strip. */
const withOption = (
  kind: string,
  opt: Record<string, unknown>,
  w: number,
  h: number,
  fs: number,
  horizontal: boolean,
) =>
  ({
    ...(sampleConfig(kind as Parameters<typeof sampleConfig>[0]) as ChartConfig),
    ...opt,
    width: w,
    height: h,
    horizontal,
    style: { fontSize: fs },
  }) as ChartConfig;

const RAISERS: [string, Record<string, unknown>][] = [
  ["secondaryAxis", { secondaryAxis: true }],
  ["pareto", { pareto: true }],
];

describe("the secondary value axis fits its own numbers", () => {
  it("draws no tick number over another, anywhere in the sweep", () => {
    // The sweep this was measured in: every kind, both options that raise the
    // strip, eight frames, two fonts, both orientations. Reported per case
    // rather than as a total — a single number hides which combination broke.
    const bad: string[] = [];
    for (const [oname, opt] of RAISERS)
      for (const { kind } of CHART_KINDS)
        for (const [w, h] of FRAMES)
          for (const fs of [10, 18])
            for (const horizontal of [false, true]) {
              let ticks: TextNode[];
              try {
                ticks = textsOf(withOption(kind, opt, w, h, fs, horizontal), /^secondary-axis$/);
              } catch {
                continue;
              }
              for (let i = 0; i < ticks.length; i++)
                for (let j = i + 1; j < ticks.length; j++)
                  if (collides(ticks[i], ticks[j])) bad.push(`${oname} ${kind} ${w}x${h} fs=${fs} H=${horizontal}`);
            }
    expect([...new Set(bad)], "secondary-axis numbers are drawn through each other").toEqual([]);
  });

  it("drops the numbers rather than drawing them below the floor", () => {
    // `tickGapScale` answers 0 where the strip cannot be read, and 0 means the
    // numbers are not drawn at all — the same answer the pie's ring, the radar's
    // names and the scatter's own axes give. A clamped-to-5pt strip would pass
    // the overlap check above and still tell the reader nothing.
    const bad: string[] = [];
    for (const [oname, opt] of RAISERS)
      for (const { kind } of CHART_KINDS)
        for (const [w, h] of FRAMES)
          for (const fs of [10, 18])
            for (const t of textsOf(withOption(kind, opt, w, h, fs, false), /^secondary-axis$/))
              if (t.fontSize < 5) bad.push(`${oname} ${kind} ${w}x${h} fs=${fs}: ${t.fontSize.toFixed(2)}pt`);
    expect(bad, "a secondary tick was drawn below the legibility floor").toEqual([]);
  });

  it("leaves a comfortable dual-axis chart at full size", () => {
    // THE REGRESSION THAT WOULD MATTER. At the sizes anyone presents at, the
    // ticks clear each other easily and the strip must not be shrunk one point
    // for the crowded case's sake.
    for (const [oname, opt] of RAISERS) {
      for (const [w, h] of [
        [480, 300],
        [960, 540],
      ]) {
        const ticks = textsOf(withOption("combo", opt, w, h, 10, false), /^secondary-axis$/);
        expect(ticks.length, `${oname} ${w}x${h}: the strip vanished on a roomy chart`).toBeGreaterThan(2);
        for (const t of ticks) {
          expect(t.fontSize, `${oname} ${w}x${h}: a roomy chart's tick was shrunk`).toBeCloseTo(10 * 0.9, 5);
        }
      }
    }
  });

  it("keeps the gridlines and the mapping when the numbers go", () => {
    // A dropped number loses a reading; the plot must still be drawn to the same
    // scale. The line's own points are positioned by `lineToY`, which the fit
    // does not touch — only the labels are sized by it.
    const tiny = withOption("combo", { secondaryAxis: true }, 120, 90, 18, false);
    const big = withOption("combo", { secondaryAxis: true }, 960, 540, 18, false);
    expect(textsOf(tiny, /^secondary-axis$/).length, "the tiny chart kept a strip it has no room for").toBe(0);
    // The overlay is still drawn on both — losing the numbers is not losing the
    // series they describe.
    for (const cfg of [tiny, big]) {
      const marks = buildChart(cfg).nodes.filter((n) => /^combo-(marker|line)-/.test(n.name ?? ""));
      expect(marks.length, "the combo overlay disappeared with its axis numbers").toBeGreaterThan(0);
    }
  });

  it("fits WIDTH sideways and HEIGHT upright — the rule itself", () => {
    // FIRST WRITTEN AS A WHOLE CHART, AND IT ASSERTED NOTHING. The idea was to
    // give a horizontal chart ten-digit numbers and check the strip shrank —
    // but numbers that wide make the strip wider than the canvas, so it is
    // dropped outright and the assertion ran against an empty list. It passed
    // against the correct code and against a mutant that measured the wrong
    // dimension, which is the worst kind of test: it reads as evidence.
    //
    // So the property is checked where it lives. Same ticks, same gap, same
    // labels; only `want` differs. A width rule sees that the numbers are wider
    // than the space between ticks and gives up 40% of the size. A line-height
    // rule sees a 14pt need against a 30pt gap and reports everything fine —
    // which is exactly how a scatter's x axis came to be drawn through itself.
    const ticks = [0, 1, 2, 3, 4];
    const to = (v: number) => v * 30; // 30pt between ticks
    const wide = ["1,234,567", "2,234,567", "3,234,567", "4,234,567", "5,234,567"];
    const byWidth = tickGapScale(9, ticks, to, 120, (t) => textWidth(wide[t], 9) + 2);
    const byHeight = tickGapScale(9, ticks, to, 120, () => 9 * 1.4);
    expect(byHeight, "a line height fits inside a 30pt gap, so nothing shrinks").toBe(1);
    expect(byWidth, "the width rule did not notice numbers wider than their gap").toBeLessThan(1);
    expect(byWidth, "the width rule shrank past the point of dropping").toBeGreaterThan(0);
    // And the caller has to pass the right one for each orientation.
    const src = readFileSync("src/core/layout/column.ts", "utf8");
    expect(src, "the sideways strip stopped measuring its labels' width").toMatch(
      /H \? \(t\) => textWidth\(formatNumber\(t, fmt2\), fs \* 0\.9\) \+ 2 : \(\) => fs \* 1\.4/,
    );
  });

  it("shifts the whole sideways strip rather than clamping label by label", () => {
    // The clamp that keeps a label on the canvas is not a free move sideways:
    // the labels are spaced along x, so pulling the last one left closes the gap
    // the fit had just opened. A horizontal pareto at 80x60 was fitted to a
    // 19.8pt gap, clamped into 14.8, and its numbers touched anyway.
    const src = readFileSync("src/core/layout/column.ts", "utf8");
    expect(src, "the per-label clamp is back on the sideways strip").toMatch(
      /const tx = H\s*\? q - fs \* 1\.7 \* tickScale2 \+ stripShift/,
    );
    // And a strip too wide to bring onto the canvas whole is dropped, not
    // crushed — the same answer the fit gives when the ticks will not fit.
    expect(src).toMatch(/if \(under < 0 && over > 0\) return NaN/);
    expect(src).toMatch(/tickScale2 > 0 && !Number\.isNaN\(stripShift\) \? ticks2 : \[\]/);
  });

  it("keeps every gap the fit measured, on the chart that exposed this", () => {
    // The integration half, on the case that actually reaches the branch: a
    // horizontal pareto whose three ticks sit near the canvas edge. Equal gaps
    // are the whole point — a clamped strip has unequal ones.
    const cfg = withOption("clustered", { pareto: true }, 80, 60, 18, true);
    const ticks = textsOf(cfg, /^secondary-axis$/);
    expect(ticks.length, "the strip this case exists for was not drawn").toBeGreaterThan(2);
    const centres = ticks.map((t) => t.x + t.w / 2).sort((a, b) => a - b);
    const gaps = centres.slice(1).map((c, i) => c - centres[i]);
    for (const g of gaps)
      expect(g, `gaps are uneven (${gaps.map((x) => x.toFixed(1)).join(", ")})`).toBeCloseTo(gaps[0], 4);
  });

  it("gives way to whatever the base chart already drew", () => {
    // NO GUTTER IS RESERVED for this strip — its own note says so — and that is
    // the whole of the remaining family. Sideways it sits above the plot at
    // `plot.y - fs * 1.5`, clamped to y=0, so a chart whose chrome has squeezed
    // the plot to the ceiling pins the strip onto the TITLE. Upright it sits two
    // points right of the plot, where a narrow chart keeps its category names,
    // its primary value axis, its legend and its series names. 184 pairs across
    // the sweep, spread thin over every kind that can carry an overlay.
    //
    // Reserving a gutter is the other option and it is worse: it would shrink
    // the plot on every dual-axis chart, including the roomy ones with nothing
    // to fix. So the add-on yields to the chart it was added to.
    const bad: string[] = [];
    for (const [oname, opt] of RAISERS)
      for (const { kind } of CHART_KINDS)
        for (const [w, h] of FRAMES)
          for (const fs of [10, 18])
            for (const horizontal of [false, true]) {
              let all: TextNode[];
              try {
                all = textsOf(withOption(kind, opt, w, h, fs, horizontal), /./);
              } catch {
                continue;
              }
              const ticks = all.filter((t) => t.name === "secondary-axis");
              const others = all.filter((t) => t.name !== "secondary-axis");
              for (const t of ticks)
                for (const o of others)
                  if (collides(t, o)) bad.push(`${oname} ${kind} ${w}x${h} fs=${fs} H=${horizontal}: on ${o.name}`);
            }
    expect([...new Set(bad)].slice(0, 8), "the secondary strip is printed over the base chart").toEqual([]);
  });

  it("keeps most of the strip while doing it", () => {
    // The yield must not empty the axis. Measured across the same sweep: 1,560
    // tick numbers become 1,237. A fifth of them go, all of them ones that were
    // being drawn through something else.
    let drawn = 0;
    for (const [, opt] of RAISERS)
      for (const { kind } of CHART_KINDS)
        for (const [w, h] of FRAMES)
          for (const fs of [10, 18])
            for (const horizontal of [false, true]) {
              try {
                drawn += textsOf(withOption(kind, opt, w, h, fs, horizontal), /^secondary-axis$/).length;
              } catch {
                /* a kind that refuses this option contributes nothing */
              }
            }
    expect(drawn, "the yield rule emptied the secondary axis").toBeGreaterThan(1000);
  });
});
