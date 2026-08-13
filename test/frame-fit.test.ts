import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { fitPlot } from "../src/core/layout/frame";
import { textWidth } from "../src/core/scene";
import type { SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/**
 * A chart is a box on a slide, and nothing it draws may leave that box.
 *
 * Neither PowerPoint renderer wraps or clips a text box, so a label wider than
 * its frame is not truncated — it is drawn across whatever the chart happens to
 * be sitting next to. The same is true of a plot: a bar hanging below the frame
 * lands on the slide, not on the chart.
 *
 * The sweep is the point of this file. Every earlier fix in this area was found
 * by looking at one chart at one size, and each time the same defect was sitting
 * in four other layouts — the title alone overflowed 13 of 17 kinds. So the
 * property is asserted over the whole product of kinds and frame sizes, and the
 * unit tests below it name the mechanisms that sweep would otherwise only
 * report as a number.
 */

/** Half a point, for the rounding in the layouts' own arithmetic. */
const SLACK = 0.5;

/**
 * The ink a node actually puts on the slide — NOT its box.
 *
 * For text the two are very different: a label's box is routinely wider than
 * its text and is anchored by `align`, so measuring boxes reports overflow that
 * is not there (a right-aligned axis label whose box starts at x=0) and misses
 * overflow that is (a centred label in a 9pt box carrying 102pt of text). An
 * early version of this sweep measured boxes and produced four false positives
 * and one false negative in a single run.
 */
function inkBox(n: SceneNode): { x0: number; y0: number; x1: number; y1: number } | null {
  const a = n as unknown as Record<string, number>;
  if (n.kind === "text") {
    const t = n as TextNode;
    const w = textWidth(t.text, t.fontSize, t.bold);
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    // Baseline by vertical alignment, then the em box around it: cap height up,
    // descender down.
    const base =
      t.valign === "top"
        ? t.y + t.fontSize
        : t.valign === "bottom"
          ? t.y + t.h - t.fontSize * 0.25
          : t.y + t.h / 2 + t.fontSize * 0.36;
    return { x0: x, y0: base - t.fontSize * 0.8, x1: x + w, y1: base + t.fontSize * 0.21 };
  }
  if (n.kind === "rect" || n.kind === "chevron") return { x0: a.x, y0: a.y, x1: a.x + a.w, y1: a.y + a.h };
  if (n.kind === "line")
    return {
      x0: Math.min(a.x1, a.x2),
      y0: Math.min(a.y1, a.y2),
      x1: Math.max(a.x1, a.x2),
      y1: Math.max(a.y1, a.y2),
    };
  if (n.kind === "ellipse") return { x0: a.cx - a.rx, y0: a.cy - a.ry, x1: a.cx + a.rx, y1: a.cy + a.ry };
  if (n.kind === "wedge") return { x0: a.cx - a.r, y0: a.cy - a.r, x1: a.cx + a.r, y1: a.cy + a.r };
  if (n.kind === "symbol") return { x0: a.cx - a.size, y0: a.cy - a.size, x1: a.cx + a.size, y1: a.cy + a.size };
  // An arrowhead's (x, y) is its TIP and its body runs 1.8*size back along
  // `angle`; the angle is not read here, so this is the disc that bounds it at
  // every rotation.
  if (n.kind === "arrowhead")
    return { x0: a.x - a.size * 1.8, y0: a.y - a.size * 1.8, x1: a.x + a.size * 1.8, y1: a.y + a.size * 1.8 };
  if (n.kind === "polygon") {
    const p = (n as { points?: { x: number; y: number }[] }).points ?? [];
    if (!p.length) return null;
    return {
      x0: Math.min(...p.map((q) => q.x)),
      y0: Math.min(...p.map((q) => q.y)),
      x1: Math.max(...p.map((q) => q.x)),
      y1: Math.max(...p.map((q) => q.y)),
    };
  }
  return null;
}

/** The worst distance any node's ink reaches past an edge of the chart, and which node. */
function worstOverflow(cfg: ChartConfig): { pt: number; node: string; side: string } {
  const scene = buildChart(cfg);
  let pt = 0;
  let node = "";
  let side = "";
  for (const n of scene.nodes) {
    const b = inkBox(n);
    if (!b) continue;
    const sides: [string, number][] = [
      ["left", -b.x0],
      ["top", -b.y0],
      ["right", b.x1 - scene.width],
      ["bottom", b.y1 - scene.height],
    ];
    for (const [which, over] of sides) {
      if (over > pt) {
        pt = over;
        node = n.name ?? n.kind;
        side = which;
      }
    }
  }
  return { pt, node, side };
}

/**
 * Frames a chart is asked to draw in. The middle three are the sizes the pane's
 * gallery and the deck's thumbnails use; the outliers are there because every
 * defect this file guards against was invisible at a comfortable size and
 * obvious at a cramped one — and because a chart's size is a number a caller
 * types, so nothing stops one arriving.
 */
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

describe("no chart draws outside its own frame", () => {
  for (const [w, h] of FRAMES) {
    it(`every kind fits ${w}x${h}`, () => {
      const bad: string[] = [];
      for (const { kind } of CHART_KINDS) {
        const o = worstOverflow({ ...sampleConfig(kind), width: w, height: h } as ChartConfig);
        if (o.pt > SLACK) bad.push(`${kind}: ${o.node} ${o.pt.toFixed(1)}pt past the ${o.side}`);
      }
      expect(bad).toEqual([]);
    });
  }
});

/**
 * The same property over the FONT, which the sweep above holds fixed.
 *
 * Every layout prices its chrome in font sizes, so the font is the other axis a
 * chart can be squeezed along — and it was never swept. Seven overflows were
 * sitting at 24 and 32pt, all of one shape: a label centred on a row, a ring or
 * a legend line, in a box `fontSize * 1.2` to `* 1.5` tall. Once the font
 * outgrew the spacing, the labels overlapped EACH OTHER at any frame size and
 * the last one left the chart at a small one. `CLAUDE.md` records the funnel
 * being fixed for exactly this; the butterfly, the gantt, the radar's ticks and
 * the scatter's legend had never been.
 *
 * A big font on a small chart is not a corner case anyone should have to defend:
 * `style.fontSize` is a number a caller types.
 */
describe("no chart draws outside its own frame at any font", () => {
  const FONTS = [6, 8, 10, 14, 18, 24, 32];
  for (const fontSize of FONTS) {
    it(`every kind at ${fontSize}pt`, () => {
      const bad: string[] = [];
      for (const [w, h] of [
        [200, 150],
        [480, 300],
        [960, 540],
      ] as [number, number][]) {
        for (const { kind } of CHART_KINDS) {
          const o = worstOverflow({ ...sampleConfig(kind), width: w, height: h, style: { fontSize } } as ChartConfig);
          if (o.pt > SLACK) bad.push(`${kind} at ${w}x${h}: ${o.node} ${o.pt.toFixed(1)}pt past the ${o.side}`);
        }
      }
      expect(bad).toEqual([]);
    });
  }
});

/**
 * Text that is drawn ON TOP of other text.
 *
 * A label inside the frame can still be unreadable, and nothing here could see
 * that: the frame sweeps above pass a chart whose every label is stacked on its
 * neighbour. A sweep of the ink boxes against each other found 73 kind/font/frame
 * combinations with overlapping text, and one shape accounted for 30 of them —
 * adjacent CATEGORY AXIS labels, which is one defect rather than seven because
 * that axis is shared by every cartesian kind.
 *
 * This gate is deliberately narrow: it pins the two that were fixed rather than
 * asserting no chart anywhere overlaps, which is not true yet. The rest are
 * listed in the PR that added this.
 */
describe("labels are not drawn on top of each other", () => {
  /**
   * The ONE overlap this gate allows, and it is a decision rather than an
   * oversight.
   *
   * A scatter or bubble point label may touch an axis TICK label. Both fixes for
   * it have been tried and measured: giving the point placer the axis labels as
   * obstacles, and confining its band to the plot. Each removes the overlaps by
   * DROPPING point labels — 56 of 301 on charts as roomy as 480x300 — because
   * the y axis owns the left margin. A point's label is data and a tick label is
   * chrome, so the trade is refused, and the reason is at the call site in
   * `layout/scatter.ts`.
   *
   * Narrow on purpose: only these two kinds, only a tick against a NUMBERED
   * point label. Anything else, in either direction, still fails. Written as an
   * exception the gate STATES rather than a frame the gate avoids — the previous
   * arrangement was the second, and it left 146 real overlapping pairs outside
   * the two frames it happened to check.
   */
  const acceptedTrade = (kind: string, a?: string, b?: string): boolean => {
    if (kind !== "scatter" && kind !== "bubble") return false;
    const tick = (x?: string) => x === "x-axis" || x === "y-axis";
    const point = (x?: string) => !!x && /^label-\d+$/.test(x);
    return (tick(a) && point(b)) || (tick(b) && point(a));
  };

  /** The ink of a text node, as the frame sweep measures it. */
  const inkOf = (t: TextNode) => inkBox(t)!;
  const overlap = (a: ReturnType<typeof inkOf>, b: ReturnType<typeof inkOf>) => {
    const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
    const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
    return w > 0 && h > 0 ? w * h : 0;
  };
  const textsNamed = (cfg: ChartConfig, re: RegExp) =>
    buildChart(cfg).nodes.filter(
      (n): n is TextNode => n.kind === "text" && !!n.name && re.test(n.name) && !!n.text.trim(),
    );

  it("the shared category axis fits each name to its slot", () => {
    // Every label is centred in a slot and none was fitted to it, so a name
    // wider than its slot ran into its neighbours on both sides.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const { kind } of CHART_KINDS) {
        const cfg = { ...sampleConfig(kind), width: 200, height: 150, style: { fontSize } } as ChartConfig;
        const cats = textsNamed(cfg, /^category-\d+$/);
        for (let i = 1; i < cats.length; i++) {
          if (overlap(inkOf(cats[i - 1]), inkOf(cats[i])) > 1) {
            bad.push(`fs=${fontSize} ${kind}: ${cats[i - 1].name} over ${cats[i].name}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("no chart overlaps its own text at the default font", () => {
    // The strongest form this file can assert today. A sweep of the text ink
    // boxes against each other found 237 overlapping pairs across the kinds and
    // fonts; fitting each label to the space it actually has took that to 76,
    // and to ZERO at the font every chart is drawn at unless someone changes it.
    // The remainder are at 18pt and above and are listed in the PR that got the
    // count here — do not widen this to those fonts without fixing them.
    const bad: string[] = [];
    for (const [w, h] of [
      // 160x120 was added after a sweep found this gate was checking TWO of the
      // eight frames the overflow sweep uses, while its name claims the default
      // font generally. See the block comment above.
      //
      // 60x300 followed: a tall sidebar is an ordinary shape for a chart on a
      // slide, and it carried 16 overlapping pairs at the default font — a
      // heatmap's column headers drawn across each other, a butterfly's two
      // series names meeting in the middle, a combo point label pinned to x=0
      // over the category names, a bubble size legend spilling onto the y axis.
      //
      // 120x90 followed once the sunburst's adjacent OUTSIDE labels were fitted
      // to the vertical gap between neighbours instead of to each wedge's own
      // arc. That one pair was the only thing left on this frame upright — it is
      // NOT in the rotated sweep below, which still has four (see there).
      //
      // 300x60 is NOT here yet and that is a statement, not an omission. Three
      // pairs survive on a 60pt-tall letterbox, down from five: a combo point
      // label against the column totals (x2) and the stacked CAGR caption
      // against the title. Both are de-collision work — `FLIPPABLE` has nowhere
      // to send a label on a frame that is almost all title. 80x60 carries nine,
      // the same two shapes plus adjacent column totals. Adding either frame
      // without fixing them would only park a known-red gate.
      [60, 300],
      [120, 90],
      [160, 120],
      [200, 150],
      [480, 300],
    ] as [number, number][]) {
      for (const { kind } of CHART_KINDS) {
        const ts = buildChart({ ...sampleConfig(kind), width: w, height: h } as ChartConfig).nodes.filter(
          (n): n is TextNode => n.kind === "text" && !!n.text.trim(),
        );
        const boxes = ts.map((t) => inkBox(t)!);
        for (let i = 0; i < boxes.length; i++) {
          for (let j = i + 1; j < boxes.length; j++) {
            if (overlap(boxes[i], boxes[j]) > 1 && !acceptedTrade(kind, ts[i].name, ts[j].name)) {
              bad.push(`${kind} at ${w}x${h}: ${ts[i].name} over ${ts[j].name}`);
            }
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("axis tick labels are fitted to the spacing between ticks", () => {
    // One label per tick, each centred on its own tick, and none was fitted to
    // the gap — 71 of the 237 pairs were axis labels drawn through each other,
    // the worst shape in the engine after the category axis.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const kind of ["stacked", "line", "scatter", "bubble", "boxplot"] as const) {
        const cfg = { ...sampleConfig(kind), width: 200, height: 150, style: { fontSize } } as ChartConfig;
        for (const axis of [/^value-axis$/, /^y-axis$/, /^x-axis$/]) {
          const ts = textsNamed(cfg, axis);
          for (let i = 1; i < ts.length; i++) {
            if (overlap(inkOf(ts[i - 1]), inkOf(ts[i])) > 1)
              bad.push(`fs=${fontSize} ${kind}: ${ts[i].name} pair ${i}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a heatmap's row labels are fitted to their row", () => {
    // The label's BOX is the row so the boxes never overlap, but its INK is the
    // font and is centred in that box — so once the font outgrew the row the
    // names were drawn through each other. Twenty pairs, and invisible to a
    // box-based check, which is why this file measures ink.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      const ts = textsNamed(
        { ...sampleConfig("heatmap"), width: 200, height: 150, style: { fontSize } } as ChartConfig,
        /^row-\d+$/,
      );
      for (let i = 1; i < ts.length; i++) {
        if (overlap(inkOf(ts[i - 1]), inkOf(ts[i])) > 1)
          bad.push(`fs=${fontSize}: ${ts[i - 1].name} over ${ts[i].name}`);
      }
    }
    expect(bad).toEqual([]);
  });

  it("a sunburst's inside labels are fitted to their wedge", () => {
    // Drawn verbatim at the ring font and bounded by nothing, so a group name
    // wider than its own arc ran into the next group's — the same defect the
    // pie's inside labels had, in the same shape. The bound is the chord of the
    // wedge at the label's radius.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      const ts = textsNamed(
        { ...sampleConfig("sunburst"), width: 200, height: 150, style: { fontSize } } as ChartConfig,
        /^group-label-\d+$/,
      );
      for (let i = 0; i < ts.length; i++) {
        for (let j = i + 1; j < ts.length; j++) {
          if (overlap(inkOf(ts[i]), inkOf(ts[j])) > 1) bad.push(`fs=${fontSize}: ${ts[i].name} over ${ts[j].name}`);
        }
      }
    }
    expect(bad).toEqual([]);
  });

  /**
   * The same property, sideways — which for a year nothing checked.
   *
   * Every fit above was written for the upright chart and several of the
   * horizontal twins had none at all: `horizontalChrome` drew its category names
   * at the chart font in a box `fs * 1.5` tall centred on the row, the mekko's
   * own fit answered `false` outright when horizontal, and the totals and the
   * combo's point labels were bounded by nothing. Rotating a chart rotates which
   * side of a label is crowded; it does not excuse the label from being fitted.
   *
   * Zero, the same as upright — which was not a foregone conclusion when this
   * was written: a coarser estimate of text width put two scatter point labels
   * in the x-axis strip, and the engine's own `textWidth` says they clear it.
   * Measure with the metric the layouts measure with.
   */
  it("no horizontal chart overlaps its own text at the default font", () => {
    const bad: string[] = [];
    for (const [w, h] of [
      // 160x120 was added after a sweep found this gate was checking TWO of the
      // eight frames the overflow sweep uses, while its name claims the default
      // font generally. See the block comment above.
      //
      // 60x300 followed: a tall sidebar is an ordinary shape for a chart on a
      // slide, and it carried 16 overlapping pairs at the default font — a
      // heatmap's column headers drawn across each other, a butterfly's two
      // series names meeting in the middle, a combo point label pinned to x=0
      // over the category names, a bubble size legend spilling onto the y axis.
      //
      // 120x90 is in the UPRIGHT sweep above and deliberately not here. Rotating
      // the sample leaves four pairs on this frame, all one shape: a mekko's
      // legend against its totals and its category names. Rotating a chart
      // rotates which side of a label is crowded, and this gate has found that
      // the fits were written for the upright chart before — so the asymmetry is
      // recorded rather than smoothed over.
      //
      // 300x60 and 80x60 are absent for the reasons given in the upright sweep,
      // plus the same mekko legend pair here.
      [60, 300],
      [160, 120],
      [200, 150],
      [480, 300],
    ] as [number, number][]) {
      for (const { kind } of CHART_KINDS) {
        const ts = buildChart({
          ...sampleConfig(kind),
          width: w,
          height: h,
          horizontal: true,
        } as ChartConfig).nodes.filter((n): n is TextNode => n.kind === "text" && !!n.text.trim());
        const boxes = ts.map((t) => inkBox(t)!);
        for (let i = 0; i < boxes.length; i++) {
          for (let j = i + 1; j < boxes.length; j++) {
            if (overlap(boxes[i], boxes[j]) > 1 && !acceptedTrade(kind, ts[i].name, ts[j].name))
              bad.push(`${kind} at ${w}x${h}: ${ts[i].name} over ${ts[j].name}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a horizontal chart's category names and totals fit the row they belong to", () => {
    // Both are centred on their row in a box `fs * 1.5` tall and neither was
    // bound by the row PITCH, so at 18pt on a 200x150 frame every name ran into
    // its neighbours — nine kinds, one shape, because they share this chrome.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const { kind } of CHART_KINDS) {
        const cfg = {
          ...sampleConfig(kind),
          width: 200,
          height: 150,
          horizontal: true,
          style: { fontSize },
        } as ChartConfig;
        for (const re of [/^category-\d+$/, /^total-\d+$/]) {
          const ts = textsNamed(cfg, re);
          for (let i = 1; i < ts.length; i++) {
            if (overlap(inkOf(ts[i - 1]), inkOf(ts[i])) > 1)
              bad.push(`fs=${fontSize} ${kind}: ${ts[i - 1].name} over ${ts[i].name}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a waterfall's value labels fit the slot their bar owns", () => {
    // Upright the box was `colThick + 12` however wide the slot was, so every
    // label could bleed 6pt into each neighbour; sideways it was `fs * 1.5` tall
    // on a row that may be thinner than that. Both orientations, one bound.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const horizontal of [false, true]) {
        const ts = textsNamed(
          { ...sampleConfig("waterfall"), width: 200, height: 150, horizontal, style: { fontSize } } as ChartConfig,
          /^label-\d+$/,
        );
        for (let i = 0; i < ts.length; i++) {
          for (let j = i + 1; j < ts.length; j++) {
            if (overlap(inkOf(ts[i]), inkOf(ts[j])) > 1)
              bad.push(`fs=${fontSize} h=${horizontal}: ${ts[i].name} over ${ts[j].name}`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a combo line's point labels fit their row when the chart is horizontal", () => {
    // Sideways a point label is centred on its category's row rather than
    // floating above its point, and it was bounded by nothing — the largest
    // single group of text collisions in the engine when this was measured.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      const ts = textsNamed(
        { ...sampleConfig("combo"), width: 200, height: 150, horizontal: true, style: { fontSize } } as ChartConfig,
        /^combo-label-\d+-\d+$/,
      );
      for (let i = 0; i < ts.length; i++) {
        for (let j = i + 1; j < ts.length; j++) {
          if (overlap(inkOf(ts[i]), inkOf(ts[j])) > 1) bad.push(`fs=${fontSize}: ${ts[i].name} over ${ts[j].name}`);
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("no tick label is drawn too small to read — it is dropped instead", () => {
    // Every axis fit shrinks a label to the room it has and none had a floor, so
    // a plot squeezed to a few points answered with ONE-POINT labels: ink rather
    // than text, stacked in the gutter, from a fit reporting success. Six y ticks
    // 1.6pt apart on a 200x150 scatter at a 26pt font were the case that named
    // it. Asserted over every kind, since the floor is shared chrome.
    const bad: string[] = [];
    for (const fontSize of [10, 18, 24, 32]) {
      for (const { kind } of CHART_KINDS) {
        for (const horizontal of [false, true]) {
          const cfg = {
            ...sampleConfig(kind),
            width: 200,
            height: 150,
            horizontal,
            style: { fontSize },
          } as ChartConfig;
          for (const t of textsNamed(cfg, /^(value-axis|x-axis|y-axis)$/)) {
            if (t.fontSize < 5) bad.push(`fs=${fontSize} ${kind} h=${horizontal}: ${t.name} at ${t.fontSize}pt`);
          }
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("keeps the series labels and the legend out of the title band", () => {
    // Both settle into the gutter above or beside the plot, and both floored
    // themselves at the CANVAS top — so when the gap they wanted could not be
    // honoured, the topmost one came to rest inside the title. The title is a
    // full-width band, so that is two texts crossed, not one label clamped.
    const bad: string[] = [];
    for (const fontSize of [18, 24, 32]) {
      for (const kind of ["area", "clustered", "line", "mekko", "stacked", "stacked100", "scatter"] as const) {
        const cfg = { ...sampleConfig(kind), width: 200, height: 150, style: { fontSize } } as ChartConfig;
        const title = textsNamed(cfg, /^title$/)[0];
        expect(title, `the ${kind} sample must carry a title for this to mean anything`).toBeTruthy();
        for (const t of textsNamed(cfg, /^(series-label-\d+|legend-\d+)$/)) {
          if (overlap(inkOf(t), inkOf(title)) > 1) bad.push(`fs=${fontSize} ${kind}: ${t.name} over the title`);
        }
      }
    }
    expect(bad).toEqual([]);
  });

  it("a cascade's drop caption stays out of the footnote", () => {
    // The caption of a block too thin for text inside it goes underneath, and
    // the plot did not reserve for that — so at the DEFAULT size it was drawn
    // 7.5pt into the footnote's band, dark text over dark text.
    for (const [w, h] of [
      [480, 300],
      [200, 150],
    ] as [number, number][]) {
      const cfg = { ...sampleConfig("cascade"), width: w, height: h } as ChartConfig;
      const foot = textsNamed(cfg, /^footnote$/)[0];
      expect(foot, "the sample cascade must carry a footnote for this to mean anything").toBeTruthy();
      for (const d of textsNamed(cfg, /^drop-label-\d+$/)) {
        expect(overlap(inkOf(d), inkOf(foot)), `${d.name} over the footnote at ${w}x${h}`).toBe(0);
      }
    }
  });
});

describe("fitPlot", () => {
  const cfg = { width: 200, height: 150 } as ChartConfig;

  it("returns a plot that fits UNTOUCHED, so no ordinary chart moves", () => {
    const plot = { x: 34, y: 20, w: 150, h: 100 };
    expect(fitPlot(cfg, plot)).toBe(plot);
  });

  it("never inverts: a chrome overspend cannot produce a negative height", () => {
    // Scatter at 120x90 computed h: -8. `toY` maps the value domain through the
    // height, so a negative one does not squash the plot — it turns the axis
    // upside down and puts the tick labels below the bottom of the chart.
    const fitted = fitPlot(cfg, { x: 34, y: 140, w: 150, h: -8 });
    expect(fitted.h).toBeGreaterThan(0);
    expect(fitted.w).toBeGreaterThan(0);
  });

  it("grows UP from the layout's own bottom edge, never down from its top", () => {
    // The bottom edge is the category axis and the value baseline. Anchoring the
    // other way pins the plot to the foot of the frame, and every label drawn
    // beneath a mark then spills out of it.
    const fitted = fitPlot(cfg, { x: 34, y: 100, w: 150, h: -20 });
    expect(fitted.y + fitted.h).toBeLessThanOrEqual(80);
    expect(fitted.y).toBeGreaterThanOrEqual(0);
  });

  it("keeps the whole plot inside the frame", () => {
    const fitted = fitPlot(cfg, { x: 190, y: 300, w: 400, h: 400 });
    expect(fitted.x).toBeGreaterThanOrEqual(0);
    expect(fitted.y).toBeGreaterThanOrEqual(0);
    expect(fitted.x + fitted.w).toBeLessThanOrEqual(cfg.width);
    expect(fitted.y + fitted.h).toBeLessThanOrEqual(cfg.height);
  });
});

describe("the mechanisms the sweep would only report as a number", () => {
  const at = (kind: string, w: number, h: number) =>
    buildChart({ ...sampleConfig(kind as never), width: w, height: h } as ChartConfig);

  it("a radar drops its perimeter labels rather than draw them off the web", () => {
    // The ring's radius reserves room for the labels on both axes; when the
    // radius floor has to override that reservation there is nowhere left for
    // them, and the bottom one sat 12pt past a 70pt frame.
    const small = at("radar", 90, 70).nodes.filter((n) => n.name?.startsWith("category-"));
    const roomy = at("radar", 480, 300).nodes.filter((n) => n.name?.startsWith("category-"));
    expect(small).toHaveLength(0);
    expect(roomy.length).toBeGreaterThan(0);
  });

  it("a sunburst keeps its INSIDE labels when the outer ring will not fit", () => {
    // Inside labels are bounded by the wedge they sit on, so the frame never
    // takes them — only the ring that hangs outside the arc.
    const nodes = at("sunburst", 90, 70).nodes.filter((n): n is TextNode => n.kind === "text");
    expect(nodes.some((n) => n.name?.startsWith("group-label-"))).toBe(true);
    expect(nodes.some((n) => n.name?.startsWith("label-"))).toBe(false);
  });

  it("a tilemap's legend stays in the frame when the grid outgrows its budget", () => {
    const scene = at("tilemap", 300, 60);
    const min = scene.nodes.find((n) => n.name === "legend-min")!;
    expect(inkBox(min)!.y1).toBeLessThanOrEqual(scene.height + SLACK);
  });

  it("a difference label flips to the other side of its arrow when the right has no room", () => {
    // `computeFrame` reserves a right margin for this label, and only the
    // cartesian frame does: a line chart puts its last category hard against the
    // plot edge, so the label started 8pt past the chart. The frame-clip at the
    // end of `buildChart` then cut "+100%" down to "+…" — inside the frame and
    // saying nothing, which is why this asserts the TEXT and not just the ink.
    const nodes = buildChart({
      kind: "line",
      width: 400,
      height: 300,
      decorations: { difference: { from: 0, to: 1 } },
      data: { categories: ["Q1", "Q2"], series: [{ name: "A", values: [10, 20] }] },
    } as ChartConfig).nodes;
    const label = nodes.find((n): n is TextNode => n.name === "diff-label")!;
    expect(label.text).toContain("100");
    expect(inkBox(label)!.x1).toBeLessThanOrEqual(400 + SLACK);
  });

  it("a CAGR arrow shortens its lift rather than put its head above the chart", () => {
    // Only the lift may move: the arrow's two x anchors and the DIFFERENCE
    // between its ends are what it claims, and shortening the lift moves both
    // ends by the same amount, so the slope survives.
    const tall = at("stacked", 300, 300).nodes;
    const flat = at("stacked", 300, 60).nodes;
    const slope = (ns: SceneNode[]) => {
      const l = ns.find((n) => n.name === "cagr-line") as unknown as Record<string, number>;
      return (l.y2 - l.y1) / (l.x2 - l.x1);
    };
    const head = flat.find((n) => n.name === "cagr-head") as unknown as Record<string, number>;
    expect(head.y - head.size * 1.8).toBeGreaterThanOrEqual(-SLACK);
    // The chart is a fifth the height, so the slope is a fifth as steep — what
    // matters is that it still points the same way and is not flattened to zero
    // by the clamp.
    expect(Math.sign(slope(flat))).toBe(Math.sign(slope(tall)));
    expect(slope(flat)).not.toBe(0);
  });
});
