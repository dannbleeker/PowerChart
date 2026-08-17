import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { legendRow } from "../src/core/layout/column";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import { horizontalLegendFits } from "../src/core/layout/frame";
import type { RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Legend layout — wrapping, reserved rows, chip mirrors the mark. */

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */
const W = 480;

const H = 300;

describe("legend chips are a miniature of the mark they label", () => {
  const paintedCfg: ChartConfig = {
    kind: "stacked",
    horizontal: true,
    width: W,
    height: H,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "Retail", values: [10, 20], color: "#2a78d6" },
        { name: "Online", values: [5, 8], color: "#2a78d6", pattern: "diagonal" },
        { name: "Plan", values: [4, 6], scenario: "PL" },
      ],
    },
    decorations: { seriesLabels: true },
  };

  it("carries pattern and the IBCS scenario restyle onto the chip", () => {
    const { nodes } = buildChart(paintedCfg);
    const chip = (i: number) => nodes.find((n) => n.name === `legend-chip-${i}`) as RectNode;
    const seg = (i: number) => nodes.find((n) => n.name === `seg-${i}-0`) as RectNode;
    // Before: three identical solid squares — two same-coloured series told
    // apart only by a hatch, and a hollow PL bar keyed by a solid block.
    // (The chip carries no separator stroke, so only the scenario restyle's
    // own outline is compared against the segment.)
    for (const i of [0, 1, 2]) {
      expect(chip(i).fill).toBe(seg(i).fill);
      expect(chip(i).pattern).toBe(seg(i).pattern);
    }
    expect(chip(0).stroke).toBeUndefined();
    expect(chip(1).pattern).toBe("diagonal");
    expect(chip(2).fill).toBe("none");
    expect(chip(2).stroke).toBe(seg(2).stroke);
    expect(chip(2).strokeWidth).toBe(seg(2).strokeWidth);
  });
});

describe("legend wraps instead of marching off-canvas", () => {
  const seriesCfg = (n: number, width: number): ChartConfig => ({
    kind: "stacked",
    width,
    height: 200,
    data: {
      categories: ["A"],
      series: Array.from({ length: n }, (_, i) => ({ name: `Series ${i + 1}`, values: [1] })),
    },
  });
  const chipYs = (nodes: ReturnType<typeof legendRow>) =>
    new Set(nodes.filter((n) => n.name?.startsWith("legend-chip-")).map((n) => (n as RectNode).y));

  it("puts overflowing chips on a second row", () => {
    // 12 wide chips cannot fit one 300pt row — they must span multiple rows.
    expect(chipYs(legendRow(seriesCfg(12, 300), DEFAULT_STYLE, 0, 0, { maxX: 300 })).size).toBeGreaterThan(1);
  });

  it("stays a single row (byte-identical) when everything fits", () => {
    expect(chipYs(legendRow(seriesCfg(2, 800), DEFAULT_STYLE, 0, 0, { maxX: 796 })).size).toBe(1);
  });
});

describe("a wrapped legend reserves its rows and never overlaps the plot", () => {
  // Horizontal bars, a narrow canvas, and six long series names: the legend
  // cannot fit one row, so #139's wrap advances it downward. Until the frame
  // reserved those extra rows, the wrapped rows painted on top of the bars.
  const cfg: ChartConfig = {
    kind: "stacked",
    horizontal: true,
    width: 360,
    height: 320,
    data: {
      categories: ["North", "South"],
      series: [
        { name: "Northern Europe Wholesale", values: [12, 9] },
        { name: "Central Europe Retail", values: [8, 11] },
        { name: "Southern Europe Online", values: [6, 7] },
        { name: "Nordics Direct-to-Consumer", values: [5, 4] },
        { name: "Baltics Partner Channel", values: [3, 6] },
        { name: "Iberia Marketplace", values: [4, 5] },
      ],
    },
    decorations: { seriesLabels: true, categoryAxis: true },
  };
  const nodes = buildChart(cfg).nodes;
  const box = (n: RectNode | TextNode) => ({ x: n.x, y: n.y, r: n.x + n.w, b: n.y + n.h });
  const overlaps = (a: ReturnType<typeof box>, c: ReturnType<typeof box>) =>
    a.x < c.r && c.x < a.r && a.y < c.b && c.y < a.b;
  const legendNodes = nodes.filter(
    (n): n is RectNode | TextNode => (n.kind === "rect" || n.kind === "text") && !!n.name?.startsWith("legend-"),
  );
  const plotRects = nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("seg-"));

  it("actually wraps to more than one row (non-vacuous)", () => {
    const chipRows = new Set(legendNodes.filter((n) => n.name?.startsWith("legend-chip-")).map((n) => n.y));
    expect(chipRows.size).toBeGreaterThan(1);
    expect(plotRects.length).toBeGreaterThan(0);
  });

  it("puts no legend chip or label on top of a bar segment", () => {
    for (const l of legendNodes) {
      for (const p of plotRects) {
        expect(overlaps(box(l), box(p))).toBe(false);
      }
    }
  });

  it("keeps every legend node inside the canvas height", () => {
    for (const l of legendNodes) {
      expect(l.y + l.h).toBeLessThanOrEqual(cfg.height + 1e-6);
    }
  });
});

/**
 * A horizontal line/area chart drew its legend TWICE: `horizontalChrome` emits
 * the shared `legendRow` under `decor.seriesLabels`, and the line layout had a
 * hand-rolled copy of its own under the same condition. The two sat 2.5pt
 * apart, so every series name rendered as a smear — visible in the shipped
 * showcase deck ("Horizontal profile chart (stacked area)"), where "Retail" and
 * "Online" were each drawn over themselves.
 */
describe("a sideways line/area chart legends its series once", () => {
  const cfg = (kind: "line" | "area"): ChartConfig => ({
    kind,
    horizontal: true,
    width: W,
    height: H,
    data: {
      categories: ["North", "South", "East"],
      series: [
        { name: "Retail", values: [10, 20, 15] },
        { name: "Online", values: [5, 8, 12] },
      ],
    },
    decorations: { seriesLabels: true },
  });

  for (const kind of ["line", "area"] as const) {
    it(`emits one chip and one label per series (${kind})`, () => {
      const { nodes } = buildChart(cfg(kind));
      for (const si of [0, 1]) {
        expect(
          nodes.filter((n) => n.name === `legend-chip-${si}`),
          `chips for series ${si}`,
        ).toHaveLength(1);
        expect(
          nodes.filter((n) => n.name === `legend-${si}`),
          `labels for series ${si}`,
        ).toHaveLength(1);
      }
      // Non-vacuous: the legend is actually drawn, not merely not duplicated.
      const labels = nodes.filter((n): n is TextNode => n.kind === "text" && /^legend-\d+$/.test(n.name ?? ""));
      expect(labels.map((n) => n.text)).toEqual(["Retail", "Online"]);
    });
  }
});

/**
 * `seriesLabelNodes` pushes overlapping labels DOWN, then — if the last one
 * overflows the plot — shifts the whole stack up and re-propagates upward, with
 * nothing clamping the top. Twelve series on a 240×160 chart put two labels at
 * negative y; thirty on a 400×300 put the topmost at −123. `collide.ts` refuses
 * to nudge a label off the top for the same reason, but it only moves labels
 * UP, so it cannot rescue one already emitted above the canvas.
 */
describe("right-hand series labels stay on the canvas", () => {
  const sized = (w: number, h: number, count: number): ChartConfig =>
    ({
      kind: "line",
      width: w,
      height: h,
      decorations: { seriesLabels: true },
      data: {
        categories: ["Q1", "Q2"],
        series: Array.from({ length: count }, (_, i) => ({ name: `S${i}`, values: [i + 1, i + 2] })),
      },
    }) as ChartConfig;

  for (const [w, h, count] of [
    [240, 160, 12],
    [400, 300, 30],
    [480, 300, 24],
  ] as const) {
    it(`fits ${count} labels on a ${w}×${h} chart`, () => {
      const labels = buildChart(sized(w, h, count)).nodes.filter(
        (n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("series-label"),
      );
      expect(labels.length, "no labels drawn, so this proves nothing").toBeGreaterThan(count / 2);
      for (const l of labels) {
        expect(l.y, "a label was drawn above the canvas").toBeGreaterThanOrEqual(-0.01);
        expect(l.y + l.h, "a label was drawn below the canvas").toBeLessThanOrEqual(h + 0.01);
      }
    });
  }

  it("keeps the labels in the order of the lines they name", () => {
    // The spread's step is by definition below the gap the labels wanted, so
    // every neighbouring pair overlaps — and `series-label-` is movable, with a
    // nudge that only goes UP. Left alone it pushed each label past the one
    // above it and handed back a full set of REORDERED labels: every one on the
    // canvas, each naming the wrong line. Shrinking the spread set to its own
    // step leaves the de-collision pass nothing to do.
    for (const [w, h, count] of [
      [240, 160, 12],
      [400, 300, 30],
      [480, 300, 24],
    ] as const) {
      const labels = buildChart(sized(w, h, count))
        .nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("series-label"))
        .sort((a, b) => a.y - b.y)
        .map((n) => n.text);
      // Series i has values [i+1, i+2], so the highest index ends highest.
      expect(labels, `${w}x${h}, ${count} series`).toEqual(
        Array.from({ length: count }, (_, i) => `S${count - 1 - i}`),
      );
    }
  });

  it("leaves a chart with room to spare exactly where it was", () => {
    // The negative control: spreading unconditionally would move every label on
    // every chart, and these three are nowhere near needing it.
    const ys = buildChart(sized(480, 300, 3))
      .nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("series-label"))
      .map((n) => Math.round(n.y * 10) / 10);
    expect(new Set(ys).size).toBe(3);
    expect(Math.min(...ys)).toBeGreaterThan(0);
  });
});

/**
 * Scatter's group legend is hand-rolled — it does not go through
 * `legendRow`/`legendRowCount` the way mekko, boxplot, radar, butterfly and
 * column all do — and it never wrapped. Entries marched right until they left
 * the frame: eight groups on the DEFAULT 480pt frame ran to x=520. In SVG the
 * viewBox clips them, so a group goes silently unexplained; in PowerPoint it is
 * worse, because the Office renderer applies no clamp and they become real
 * shapes sitting off the chart on the slide.
 */
describe("the scatter group legend wraps instead of marching off-canvas", () => {
  const scatterWith = (groups: number, width: number) => {
    const cats = Array.from({ length: groups }, (_, i) => `P${i}`);
    return buildChart({
      kind: "scatter",
      width,
      height: 180,
      data: {
        categories: cats,
        series: [
          { name: "X", values: cats.map((_, i) => i + 1) },
          { name: "Y", values: cats.map((_, i) => groups - i) },
          { name: "Group", values: cats.map((_, i) => i + 1) },
        ],
      },
    } as unknown as ChartConfig);
  };

  it("keeps every legend chip and label inside the frame", () => {
    for (const [groups, width] of [
      [3, 180],
      [4, 240],
      [6, 360],
      [8, 480],
      [12, 480],
    ] as const) {
      const scene = scatterWith(groups, width);
      const legend = scene.nodes.filter((n) => /^legend-/.test(n.name ?? ""));
      expect(legend.length, `${groups} groups produced no legend`).toBeGreaterThan(0);
      for (const n of legend) {
        const right = (n as { x: number; w?: number }).x + ((n as { w?: number }).w ?? 0);
        expect(right, `${n.name} runs to ${right} on a ${width}pt frame`).toBeLessThanOrEqual(width + 0.5);
      }
    }
  });

  it("still draws a legend that fits on one row exactly where it always did", () => {
    // The reservation grows only when the walk actually wraps, so the common
    // case must be untouched.
    const one = scatterWith(2, 480);
    const chips = one.nodes.filter((n) => /^legend-chip-/.test(n.name ?? ""));
    expect(chips.length).toBe(2);
    expect(new Set(chips.map((c) => (c as { y: number }).y)).size, "a two-entry legend wrapped").toBe(1);
  });
});

describe("a horizontal legend the frame cannot pay for", () => {
  /**
   * `computeFrameHorizontal` reserves a row per wrapped legend row, and
   * `horizontalChrome` draws them. On an 80x60 frame the legend wraps to two
   * rows — about 39pt of a 60pt chart — so with the title and the value axis
   * there is no plot left, `fitPlot` floors it, and the legend's second row
   * comes down on the value-axis labels.
   *
   * Both sides now read ONE predicate, `horizontalLegendFits`, which is the
   * point of the test. Gating the reservation alone is measurably worse than
   * not gating at all: doing exactly that to the scatter legend took an
   * extreme-frame overlap count from 55 to 63, because the band is reserved as
   * zero and the legend is drawn anyway.
   */
  const legendTexts = (w: number, h: number) =>
    buildChart({ ...sampleConfig("line"), width: w, height: h, horizontal: true } as ChartConfig).nodes.filter(
      (n): n is TextNode => n.kind === "text" && /^legend/.test(String(n.name)),
    );

  it("drops the legend rather than drawing it over the value axis", () => {
    expect(legendTexts(80, 60)).toHaveLength(0);
  });

  it("still draws it on a frame with room", () => {
    // The negative control: this must not become "drop the legend always".
    expect(legendTexts(480, 300).length).toBeGreaterThan(0);
  });
});

/**
 * The horizontal mekko's legend, and the third call site of a predicate that
 * only had two.
 *
 * `horizontalLegendFits` exists so a legend's RESERVATION and its DRAW cannot
 * disagree — `computeFrameHorizontal` asks it before counting any rows, and
 * `horizontalChrome` asks it before drawing any chips. The mekko draws its own
 * legend and never asked, so on a frame the predicate refuses the reservation
 * was ZERO rows and the legend was drawn into that zero anyway.
 *
 * That is not a smaller version of the bug — it is the worst of the three
 * options, and the one a half-fix produces. This repo has measured it before:
 * gating only the reservation took the scatter from 55 overlapping pairs to 63.
 * Here it put a three-row legend across the bars, their totals and the row
 * labels at 120x90 horizontal — four pairs, all inside the frame, so no
 * overflow gate could see them.
 */
describe("the horizontal mekko legend asks the same predicate the reservation does", () => {
  const mekko = (w: number, h: number) =>
    buildChart({ ...sampleConfig("mekko"), width: w, height: h, horizontal: true } as ChartConfig).nodes.filter(
      (n) => !!n.name?.startsWith("legend-"),
    );

  it("still draws the legend on a frame with room", () => {
    // The guard that stops this becoming "the horizontal mekko lost its legend".
    expect(mekko(480, 300).length, "no legend on a roomy horizontal mekko").toBeGreaterThan(0);
  });

  it.each([
    [120, 90],
    [300, 60],
    [80, 60],
  ])("draws no legend at %ix%i, where the frame reserved no rows for one", (w, h) => {
    expect(mekko(w, h)).toEqual([]);
  });

  it("agrees with the reservation at every frame, in both directions", () => {
    // The property rather than the three sizes: whatever the predicate says, the
    // draw matches it. A future change that alters the predicate's threshold
    // moves both sides together or fails here.
    const cfg0 = sampleConfig("mekko") as ChartConfig;
    for (const [w, h] of [
      [80, 60],
      [120, 90],
      [160, 120],
      [300, 60],
      [200, 150],
      [480, 300],
      [960, 540],
    ] as [number, number][]) {
      const cfg = { ...cfg0, width: w, height: h, horizontal: true } as ChartConfig;
      const fits = horizontalLegendFits(cfg, DEFAULT_STYLE, { ...DEFAULT_DECOR, ...cfg.decorations, totals: true });
      const drawn = mekko(w, h).length > 0;
      expect(drawn, `${w}x${h}: predicate says ${fits} but drawn=${drawn}`).toBe(fits);
    }
  });
});

/**
 * A band with no room in it does not get a legend by shrinking.
 *
 * `seriesLabelNodes` spreads its labels over whatever vertical band is left
 * once the title has taken its share, then shrinks them to the step it ended up
 * with — and that shrink is FLOORED at 5pt. So past the floor the spread stops
 * paying for itself: the labels are pitched closer together than their own
 * height, which is exactly the overlap the shrink exists to prevent. At the
 * limit the band has no height at all (`top === bottom`, a short frame whose
 * title has eaten the gutter) and every label lands on the SAME POINT.
 *
 * Measured across 25 kinds × 8 frames × 7 fonts before the fix: 46 pairs of
 * series labels drawn at identical coordinates, on mekko, line, area, combo and
 * the column family. A reader sees the one drawn last and has no way to know
 * which series it names — and `collide.ts` cannot rescue them, because its
 * nudge only goes up, all of them want the same place, and they exhaust the
 * budget still stacked.
 *
 * The answer is the one every other reservation in this engine gives when it
 * cannot be paid for (the radar's ticks, the sunburst's ring, the pie's outside
 * labels): drop them. The same sweep after the fix: 0 coincident pairs, and the
 * total number of overlapping text pairs falls from 838 to 758 — so nothing is
 * traded for it. The default font is untouched at every frame; the drop starts
 * at 14pt.
 */
describe("series labels are dropped rather than stacked when the band cannot hold them", () => {
  const FRAMES: [number, number][] = [
    [60, 300],
    [80, 60],
    [120, 90],
    [160, 120],
    [200, 150],
    [300, 60],
    [480, 300],
    [960, 540],
  ];
  const FONTS = [6, 8, 10, 14, 18, 24, 32];

  const labelsOf = (kind: ChartConfig["kind"], w: number, h: number, fontSize: number) =>
    buildChart({ ...sampleConfig(kind), width: w, height: h, style: { fontSize } } as ChartConfig).nodes.filter(
      (n): n is TextNode => n.kind === "text" && /^series-label-\d+$/.test(n.name ?? ""),
    );

  it("never draws two series labels at the same point", () => {
    const stacked: string[] = [];
    let seen = 0;
    for (const kind of ["mekko", "line", "area", "combo", "stacked", "clustered"] as ChartConfig["kind"][])
      for (const [w, h] of FRAMES)
        for (const fs of FONTS) {
          const labels = labelsOf(kind, w, h, fs);
          seen += labels.length;
          for (let i = 0; i < labels.length; i++)
            for (let j = i + 1; j < labels.length; j++)
              if (Math.abs(labels[i].y - labels[j].y) < 0.01 && Math.abs(labels[i].x - labels[j].x) < 0.01)
                stacked.push(`${kind} ${w}x${h} fs${fs}: "${labels[i].text}" under "${labels[j].text}"`);
        }
    expect(seen, "no series labels drawn at all, so this proves nothing").toBeGreaterThan(200);
    expect(stacked).toEqual([]);
  });

  it("drops the whole set rather than keeping an unreadable subset", () => {
    // The concrete case the sweep found: three series names at one point, on a
    // frame whose title leaves the label gutter no height.
    expect(labelsOf("mekko", 160, 120, 32)).toEqual([]);
    expect(labelsOf("mekko", 300, 60, 24)).toEqual([]);
  });

  it("leaves every chart that has the room exactly as it was", () => {
    // The negative control. At the default font no frame in the sweep loses a
    // label, so an ordinary chart cannot be touched by this.
    for (const [w, h] of FRAMES) {
      const labels = labelsOf("mekko", w, h, 10);
      expect(labels.length, `mekko ${w}x${h} at the default font lost its series labels`).toBe(3);
      expect(new Set(labels.map((l) => Math.round(l.y * 10) / 10)).size, `mekko ${w}x${h}`).toBe(3);
    }
  });
});
