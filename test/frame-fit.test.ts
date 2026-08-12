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
