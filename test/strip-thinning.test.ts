import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { readFileSync } from "node:fs";
import { textWidth, type SceneNode, type TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * Two strips of chrome that were drawn through themselves, and the guard the
 * overlap ratchet cannot give them.
 *
 * `test/overlap-budget.test.ts` counts collisions and would catch either of
 * these coming back. What it CANNOT tell is a fix from a rout: a rule that
 * dropped every label on both strips would take the count to zero and pass. So
 * the counts belong here, beside the fix.
 *
 *     timeline / timeline               36 pairs -> 0
 *     size-legend-label / size-legend-label   34 -> 0
 *
 * They failed for different reasons and both reasons are ones this engine has
 * met before.
 *
 * AND THE TWO LAYERS EARN EACH OTHER. Mutating the gantt gap back to the old
 * constant passes every test in THIS file — the sample gantt's dates are short
 * enough that `fs * 2.6` still clears them — and the ratchet catches it at once,
 * naming `timeline / timeline` as a shape that has reappeared, sixteen times,
 * under variants this file does not build. The reverse holds too: a rule that
 * dropped every date would sail through the ratchet and fail here. Neither is
 * sufficient; both are cheap.
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
const textsOf = (cfg: ChartConfig, re: RegExp): TextNode[] =>
  buildChart(cfg).nodes.filter(
    (n: SceneNode): n is TextNode => n.kind === "text" && re.test(n.name ?? "") && !!String(n.text).trim(),
  );
const at = (kind: Parameters<typeof sampleConfig>[0], w: number, h: number, fontSize: number): ChartConfig =>
  ({ ...(sampleConfig(kind) as ChartConfig), width: w, height: h, style: { fontSize } }) as ChartConfig;
const pairs = (ls: TextNode[]) => {
  let n = 0;
  for (let i = 0; i < ls.length; i++) for (let j = i + 1; j < ls.length; j++) if (collides(ls[i], ls[j])) n++;
  return n;
};

describe("a gantt's date strip", () => {
  it("draws no date over another", () => {
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 14, 18, 24])
        expect(pairs(textsOf(at("gantt", w, h, fontSize), /^timeline$/)), `${w}x${h} fs=${fontSize}`).toBe(0);
  });

  it("thins by the WIDEST DATE, not by a constant", () => {
    // `minLabelGap` was `fs * 2.6`, and a constant cannot know how wide a date
    // is. "Dec" clears it easily; a full month-and-year is more than twice it at
    // the same font, so two of them came through side by side and were drawn
    // into each other. The floor keeps a strip of SHORT labels thinning exactly
    // as it always did, so an ordinary gantt does not move.
    const dates = textsOf(at("gantt", 480, 300, 10), /^timeline$/);
    expect(dates.length, "the date strip was emptied rather than thinned").toBeGreaterThan(1);
    // Every surviving pair really is at least a label apart.
    const xs = dates.map((d) => d.x + d.w / 2).sort((a, b) => a - b);
    const widest = Math.max(...dates.map((d) => textWidth(d.text, d.fontSize)));
    for (let i = 1; i < xs.length; i++)
      expect(xs[i] - xs[i - 1], "two dates are closer together than one date is wide").toBeGreaterThanOrEqual(widest);
  });

  it("measures the nudge at the size the date is DRAWN", () => {
    // `headFs` is `bandFontSize(fs * 0.9, …)`, so it equals `fs * 0.9` only when
    // the header band can afford it. Measuring the end-nudge at the larger size
    // pushed the first date further right than it had to go — toward its
    // neighbour. The same class of mistake as the gauge total measured unbold.
    const src = readFileSync("src/core/layout/gantt.ts", "utf8");
    expect(src, "the nudge went back to measuring at fs * 0.9").toMatch(/const half = textWidth\(text, headFs\) \/ 2/);
    // And the thinning is tested on the NUDGED position, not the raw tick —
    // otherwise the nudge closes the gap the thinning just verified.
    expect(src, "thinning is back on the raw tick position").toMatch(/if \(at - lastLabelX >= minLabelGap\)/);
    expect(src).toMatch(/lastLabelX = at;/);
  });

  it("keeps both ends of the timeline on an ordinary gantt", () => {
    // THE REGRESSION THAT WOULD MATTER: the end dates are what say what period
    // the chart covers, and the nudge exists to keep them on the canvas.
    const dates = textsOf(at("gantt", 960, 540, 10), /^timeline$/);
    expect(dates.length, "a roomy gantt lost its date strip").toBeGreaterThan(2);
    for (const d of dates) {
      expect(d.x + d.w / 2 - textWidth(d.text, d.fontSize) / 2, `${d.text} runs off the left`).toBeGreaterThanOrEqual(
        0,
      );
      expect(d.x + d.w / 2 + textWidth(d.text, d.fontSize) / 2, `${d.text} runs off the right`).toBeLessThanOrEqual(
        960,
      );
    }
  });
});

describe("a bubble chart's size key", () => {
  it("draws no reference number over another", () => {
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 14, 18, 24])
        expect(pairs(textsOf(at("bubble", w, h, fontSize), /^size-legend-label-/)), `${w}x${h} fs=${fontSize}`).toBe(0);
  });

  it("fits each number to its own circle rather than spilling out of it", () => {
    // The box is `r * 2` — the circle's diameter — and the text was always
    // `fs * 0.8`. The circles are nested reference sizes, so the smallest is a
    // 5pt box, and a centred number wider than its box spills equally off both
    // sides onto the neighbouring key. Exactly the gauge total's defect.
    for (const [w, h] of FRAMES)
      for (const fontSize of [10, 14, 18, 24])
        for (const t of textsOf(at("bubble", w, h, fontSize), /^size-legend-label-/)) {
          expect(
            textWidth(t.text, t.fontSize, t.bold),
            `${w}x${h} fs=${fontSize}: "${t.text}" is wider than the circle it labels`,
          ).toBeLessThanOrEqual(t.w + 1e-9); // the fit makes these EQUAL, so allow float dust
          expect(t.fontSize, `${w}x${h} fs=${fontSize}: below the legibility floor`).toBeGreaterThanOrEqual(5);
        }
  });

  it("still labels the key on a chart with room", () => {
    // THE REGRESSION THAT WOULD MATTER. The key is three nested circles and
    // their values; dropping every value leaves three circles meaning nothing.
    // The smallest circle may lose its number — it is the least informative of
    // the three, and the circle still shows the size — but not all of them.
    const keys = textsOf(at("bubble", 960, 540, 10), /^size-legend-label-/);
    expect(keys.length, "the size key lost every number").toBeGreaterThan(1);
  });
});
