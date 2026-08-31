import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { textWidth } from "../src/core/scene";
import type { RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Funnel — band geometry, short-frame fit, conversion markers. */

/**
 * Degenerate-frame / degenerate-scale guards found by a layout bug-hunt. Each is
 * byte-identical for normal charts (snapshots unchanged) and only repairs an edge
 * that previously emitted negative geometry or NaN coordinates.
 */
const negDims = (nodes: ReturnType<typeof buildChart>["nodes"]) =>
  nodes.filter(
    (n) =>
      ((n.kind === "rect" || n.kind === "text") && ((n as RectNode).w < 0 || (n as RectNode).h < 0)) ||
      Object.entries(n).some(([k, v]) => ["x", "y", "w", "h"].includes(k) && typeof v === "number" && Number.isNaN(v)),
  );

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */
const W = 480;

const H = 300;

const texts = (nodes: SceneNode[], namePrefix: string) =>
  nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith(namePrefix)).map((n) => n.text);

const base = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  ...DEFAULT_SIZE,
  width: W_offframeguards,
  height: H,
  data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 20, 30] }] },
  ...partial,
});

/**
 * Geometry that must stay inside the canvas. These were found by the review as
 * off-frame / wrong-size layout bugs the loose fuzz bound (|c| < 5000) hid.
 */
const W_offframeguards = 480;

describe("funnel", () => {
  const s = buildChart(sampleConfig("funnel"));

  it("draws centered bands with width proportional to value", () => {
    const bands = [0, 1, 4].map((c) => s.nodes.find((n) => n.name === `stage-${c}`) as RectNode);
    // Widths proportional: 720/1200, 120/1200.
    expect(bands[1].w / bands[0].w).toBeCloseTo(720 / 1200, 2);
    expect(bands[2].w / bands[0].w).toBeCloseTo(120 / 1200, 2);
    // Centered: all bands share the same center x.
    const cx = (r: RectNode) => r.x + r.w / 2;
    expect(cx(bands[1])).toBeCloseTo(cx(bands[0]), 5);
    expect(cx(bands[2])).toBeCloseTo(cx(bands[0]), 5);
  });

  it("labels conversion vs the previous stage between bands", () => {
    const conv = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("conversion-"));
    expect(conv).toHaveLength(4);
    expect(conv[0].text).toContain("60.0%"); // 720/1200
    // Stage names on the left, values on/beside the bands.
    expect(s.nodes.some((n) => n.name === "category-0")).toBe(true);
    expect((s.nodes.find((n) => n.name === "stage-value-4") as TextNode).text).toBe("120");
  });

  it("narrow bands put their value beside the band, wide ones inside", () => {
    const wide = s.nodes.find((n) => n.name === "stage-value-0") as TextNode;
    const band0 = s.nodes.find((n) => n.name === "stage-0") as RectNode;
    expect(wide.x).toBeGreaterThanOrEqual(band0.x); // inside
    const tiny = buildChart({
      ...sampleConfig("funnel"),
      data: { categories: ["All", "Won"], series: [{ name: "Deals", values: [10000, 12] }] },
    });
    const narrow = tiny.nodes.find((n) => n.name === "stage-value-1") as TextNode;
    const band = tiny.nodes.find((n) => n.name === "stage-1") as RectNode;
    expect(narrow.x).toBeGreaterThan(band.x + band.w); // beside
  });
});

describe("funnel bands never go negative on a short, crowded frame", () => {
  it("floors band height", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: 640,
      height: 60, // 20 stages + 1.5em gaps can't fit → bands went negative
      data: {
        categories: Array.from({ length: 20 }, (_, i) => `S${i}`),
        series: [{ name: "v", values: Array.from({ length: 20 }, (_, i) => 20 - i) }],
      },
    };
    expect(negDims(buildChart(cfg).nodes)).toHaveLength(0);
  });
});

describe("funnel conversion marker follows the direction of the step", () => {
  it("marks a rise with ▴ and a fall with ▾", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: W,
      height: H,
      // Ascending — the pyramid ordering funnel.ts recommends.
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 50, 0] }] },
    };
    // Before: "▾ 500.0%" — a down arrow on a 5x increase.
    expect(texts(buildChart(cfg).nodes, "conversion-")).toEqual(["▴ 500.0%", "▾ 0.0%"]);
  });

  it("drops the marker when the stage is unchanged", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: W,
      height: H,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 10] }] },
    };
    expect(texts(buildChart(cfg).nodes, "conversion-")).toEqual(["100.0%"]);
  });
});

describe("funnel bands fit a short frame", () => {
  it("many stages on a short frame stay inside the canvas", () => {
    const scene = buildChart(
      base({
        kind: "funnel",
        height: 60, // deliberately short: a fixed gap used to overshoot the bottom
        data: { categories: ["a", "b", "c", "d", "e", "f"], series: [{ name: "S", values: [60, 50, 40, 30, 20, 10] }] },
      }),
    );
    const stages = scene.nodes.filter((n) => /^stage-\d+$/.test(n.name ?? "")); // the bands, not stage-value-*
    expect(stages.length).toBe(6);
    for (const s of stages as any[]) expect(s.y + s.h).toBeLessThanOrEqual(60 + 1);
  });
});

describe("a stage value is never drawn off the chart", () => {
  /**
   * "Outside" means to the RIGHT of the band, and the widest band already
   * reaches the edge of the plot — so at a large font there is no room out there
   * and the label was drawn past the frame. `stage-value-0` landed at x = 480.0
   * on a 480pt frame: the TOP stage showed no value at all while every stage
   * below it did, because the one band big enough to matter is the one with
   * nothing to its right. Rendered and looked at, which is how it was seen: the
   * funnel had four numbers on five bars.
   *
   * The band that cannot fit a label beside it is the WIDEST one, which is
   * exactly the band with the most room inside it — so it goes inside, in
   * contrast ink. A cramped label on the bar beats a missing number.
   */
  const inkSpan = (t: TextNode) => {
    const w = Math.min(t.w, textWidth(t.text, t.fontSize, t.bold));
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    return { x0: x, x1: x + w };
  };
  const values = (fontSize: number) =>
    buildChart({ ...sampleConfig("funnel"), style: { fontSize } } as ChartConfig).nodes.filter(
      (n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("stage-value-"),
    );

  it("keeps every stage's value inside the frame, at every font size", () => {
    for (const fs of [10, 16, 22, 28, 36]) {
      const found = values(fs);
      // Every stage still HAS a value — the rule is not satisfied by dropping
      // the label that would not fit.
      expect(found.length, `fs=${fs} lost a stage value`).toBe(5);
      for (const t of found) {
        const { x0, x1 } = inkSpan(t);
        expect(x1, `fs=${fs} ${t.name} ran past the right edge`).toBeLessThanOrEqual(DEFAULT_SIZE.width);
        expect(x0, `fs=${fs} ${t.name} started left of the frame`).toBeGreaterThanOrEqual(0);
      }
    }
  });

  it("still puts the value beside the band when there is room for it", () => {
    // The fix must not collapse into "always inside": at a small font the narrow
    // lower stages carry their value outside the bar, which is what makes them
    // readable at all.
    const outside = values(10).filter((t) => t.align === "left");
    const inside = values(10).filter((t) => t.align === "center");
    expect(inside.length, "no value sits inside its band").toBeGreaterThan(0);
    expect(outside.length + inside.length).toBe(5);
  });
});

describe("a funnel's row labels stay on the chart at any font", () => {
  /**
   * The bands are solved for the plot, so they always fit. Their LABELS were
   * not, and each escaped a different way once the font outgrew the chart:
   *
   *   - a label is centred in its band and its ink is about `f` tall, so past
   *     the band height the bottom row's label hung BELOW the plot — 4.4pt at a
   *     28pt font, 15.1 at 36;
   *   - a category name is right-aligned in a gutter capped at 28% of the
   *     width, so a name wider than the cap ran off the LEFT edge —
   *     "Negotiation" by 35.9pt at 28, 83.4 at 36.
   *
   * Neither PowerPoint renderer wraps or clips a text box, so the ink measured
   * here is not bounded by the box it sits in — which is why a check that
   * clamped to the box saw the vertical half of this and missed the horizontal.
   */
  const inkOf = (t: TextNode) => {
    const w = textWidth(t.text, t.fontSize, t.bold);
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    const base =
      t.valign === "top"
        ? t.y + t.fontSize
        : t.valign === "bottom"
          ? t.y + t.h - t.fontSize * 0.25
          : t.y + t.h / 2 + t.fontSize * 0.36;
    return { x0: x, x1: x + w, y1: base + t.fontSize * 0.21 };
  };
  const rowText = (fontSize: number) => {
    const scene = buildChart({ ...sampleConfig("funnel"), style: { fontSize } } as ChartConfig);
    return {
      scene,
      nodes: scene.nodes.filter(
        (n): n is TextNode =>
          n.kind === "text" && (!!n.name?.startsWith("category-") || !!n.name?.startsWith("stage-value-")),
      ),
    };
  };

  it("keeps every stage name and value inside the frame", () => {
    for (const fs of [10, 16, 22, 28, 36]) {
      const { nodes } = rowText(fs);
      expect(nodes.length, `fs=${fs} lost row labels`).toBeGreaterThan(5);
      for (const t of nodes) {
        const ink = inkOf(t);
        expect(ink.x0, `fs=${fs} ${t.name} ran off the left edge`).toBeGreaterThanOrEqual(-0.5);
        expect(ink.x1, `fs=${fs} ${t.name} ran off the right edge`).toBeLessThanOrEqual(DEFAULT_SIZE.width + 0.5);
        expect(ink.y1, `fs=${fs} ${t.name} hung below the frame`).toBeLessThanOrEqual(DEFAULT_SIZE.height + 0.5);
      }
    }
  });

  it("shrinks only as a last resort, and keeps the conversion rate subordinate", () => {
    // At a font that fits, nothing moves at all.
    for (const t of rowText(10).nodes) expect(t.fontSize).toBe(10);
    // One size for the whole row, so a name and its number match.
    expect(new Set(rowText(28).nodes.map((t) => t.fontSize)).size).toBe(1);
    // The conversion rate is muted secondary text and must stay SMALLER than the
    // stage names beside it — shrinking only the names inverted that.
    const scene = buildChart({ ...sampleConfig("funnel"), style: { fontSize: 28 } } as ChartConfig);
    const conv = scene.nodes.find((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("conversion-"))!;
    const name = scene.nodes.find((n): n is TextNode => n.kind === "text" && n.name === "category-0")!;
    expect(conv.fontSize).toBeLessThan(name.fontSize);
  });
});

/**
 * A small frame keeps the BANDS, not the gaps between them.
 *
 * The gap holds the conversion label and is priced at `fs * 1.5` whatever the
 * plot is. The split used to be "reserve a point per band, then give the rest to
 * the gaps", which on a 120x90 frame put four gaps in 59 of 64 points of plot
 * and left five HAIRLINE bands at the 1pt floor — a degenerate geometry rather
 * than a small chart, and one no frame or overflow check can see because it is
 * comfortably inside its box.
 */
describe("a funnel too short for its gaps", () => {
  const bands = (width: number, height: number) =>
    buildChart({ ...sampleConfig("funnel"), width, height } as ChartConfig).nodes.filter(
      (n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("stage-") && !n.name.startsWith("stage-value"),
    );

  it("does not shrink every band to the hairline floor", () => {
    const small = bands(120, 90);
    expect(small.length).toBeGreaterThan(1);
    for (const b of small) {
      expect(b.h, `a band fell to the floor: ${b.name}`).toBeGreaterThan(3);
    }
    // And they still fit: the whole point of the clamp this replaced was that a
    // fixed gap drove the last bands off the bottom of the plot.
    const last = small[small.length - 1];
    expect(last.y + last.h).toBeLessThanOrEqual(90);
  });

  it("leaves a frame that can afford its gaps alone", () => {
    // At 300x200 the fixed `fs * 1.5` still fits inside half the plot, so it
    // wins and the geometry is what it always was.
    const roomy = bands(300, 200);
    const pitch = roomy[1].y - roomy[0].y;
    expect(pitch - roomy[0].h).toBeCloseTo(10 * 1.5, 5);
  });
});

describe("the conversion rate stays inside the gap it is written in", () => {
  /**
   * The conversion label sits in the gap BETWEEN two bands (`h: gap`), and its
   * font was `labelFs * 0.85` whatever that gap turned out to be. `gap` is
   * `min(fs * 1.5, half the plot shared between the bands)`, so on a short frame
   * it collapses while the font does not — the ink then spills onto the bands
   * either side and lands on their `stage-value`, the single commonest
   * overlapping pair at small frames.
   *
   * Pinned here rather than in `frame-fit`'s overlap sweep: that gate covers
   * 60x300 and up, and this shows at 80x60. Nothing leaves the frame either, so
   * no overflow gate could ever have seen it.
   *
   * The first version of this test matched rects named `band-` — the funnel
   * calls them `stage-` — found none, and `continue`d past every frame while
   * reporting a pass. It counts what it checked now and fails if that is zero,
   * because a guard that silently examines nothing is worse than no guard.
   */
  it("never draws a conversion label taller than its own gap", () => {
    let checked = 0;
    for (const [w, h] of [
      [80, 60],
      [120, 90],
      [300, 60],
      [480, 300],
    ] as [number, number][]) {
      const nodes = buildChart({ ...sampleConfig("funnel"), width: w, height: h } as ChartConfig).nodes;
      const stages = nodes
        .filter((n): n is RectNode => n.kind === "rect" && /^stage-\d+$/.test(String(n.name)))
        .sort((a, b) => a.y - b.y);
      const convs = nodes.filter((n): n is TextNode => n.kind === "text" && /^conversion-/.test(String(n.name)));
      if (stages.length < 2) continue;
      const gap = stages[1].y - (stages[0].y + stages[0].h);
      // Dropping the labels entirely is a legitimate answer on a frame with no
      // gap to write them in — what must never happen is one drawn bigger than
      // the space it sits in.
      for (const c of convs) {
        checked++;
        expect(
          c.fontSize,
          `a ${c.fontSize.toFixed(1)}pt conversion label in a ${gap.toFixed(1)}pt gap at ${w}x${h}`,
        ).toBeLessThanOrEqual(Math.max(gap, 0) + 0.01);
      }
    }
    expect(checked, "no conversion label was examined at any frame — the test matched nothing").toBeGreaterThan(0);
  });
});

/**
 * A BLANK CELL IS NOT A ZERO, and for three kinds it used to become one.
 *
 * funnel, waffle and cascade all opened with the same line —
 * `Math.max(0, data.series[0]?.values[c] ?? 0)` — and then formatted their
 * labels off the CLAMPED array. So "no data" and "negative" both reached the
 * label as the number zero and the chart asserted it. Measured on
 * `[1000, null, 250]` before the fix:
 *
 *     funnel    stage-value-1  = "0"
 *     waffle    legend-label-1 = "Signups  0%"
 *     cascade   drop-label-1   = "Other: 1,000 (100.0%)"
 *
 * The third is the one that shows why this ranks above a crash: cascade did not
 * merely print a wrong number, it DERIVED a total collapse from an empty cell
 * and captioned it. Every other kind — clustered, line, pie, treemap, sunburst
 * — has always omitted a blank.
 *
 * Geometry may still clamp; a band cannot have negative width. Text may not.
 */
describe("a value nobody supplied is never printed as a measurement", () => {
  const withBlank = (kind: string): SceneNode[] =>
    buildChart({
      kind,
      ...DEFAULT_SIZE,
      data: { categories: ["Visits", "Signups", "Paid"], series: [{ name: "S", values: [1000, null, 250] }] },
    } as unknown as ChartConfig).nodes;

  const texts = (ns: SceneNode[], re: RegExp) =>
    ns.filter((n): n is TextNode => n.kind === "text" && re.test(n.name || "")).map((n) => String(n.text));

  it("omits the funnel's value label for a stage with no value", () => {
    const ns = withBlank("funnel");
    expect(texts(ns, /^stage-value-/), "printed a value for a stage that has none").toEqual(["1,000", "250"]);
    // And no conversion rate either: a percentage of an unknown is unknown.
    const conv = texts(ns, /^conv/).join(" ");
    expect(conv, `stated a conversion against a blank stage: ${conv}`).not.toMatch(/0\.0%|100\.0%/);
  });

  it("omits the waffle's share for a category with no value", () => {
    const legend = texts(withBlank("waffle"), /^legend-label-/);
    expect(legend, "stated a 0% share for a category nobody measured").toEqual(["Visits  80%", "Signups", "Paid  20%"]);
  });

  it("does not invent a cascade drop out of an empty cell", () => {
    const ns = withBlank("cascade");
    const drops = texts(ns, /^drop-label-/).join(" | ");
    expect(drops, `invented a drop from a blank stage: ${drops}`).not.toMatch(/100\.0%/);
    // The stage itself carries no number either.
    expect(texts(ns, /^stage-label-1-/), "labelled a stage that has no value").toEqual([]);
  });

  it("changes nothing when every value is present", () => {
    // The guard must not cost a complete chart its labels — this is the case
    // every existing test and the whole shipped deck depend on.
    const full = (kind: string) =>
      buildChart({
        kind,
        ...DEFAULT_SIZE,
        data: { categories: ["Visits", "Signups", "Paid"], series: [{ name: "S", values: [1000, 600, 250] }] },
      } as unknown as ChartConfig).nodes;
    expect(texts(full("funnel"), /^stage-value-/)).toEqual(["1,000", "600", "250"]);
    expect(texts(full("waffle"), /^legend-label-/).every((t) => /%/.test(t))).toBe(true);
    expect(texts(full("cascade"), /^stage-label-1-/).join(" "), "a complete cascade lost its stage").toMatch(/600/);
  });

  /**
   * THE HALF THAT WAS DEFERRED, and it reaches the same "0%" by the same route.
   *
   * The blank-cell fix above kept `null` out of the labels. `Math.max(0, …)` is
   * still in front of the geometry, though, so a NEGATIVE cell also arrives at
   * the label as zero. Measured on `[1000, -50, 250]` before this:
   *
   *     funnel    conversion-1 = "▾ 0.0%"
   *     waffle    legend-label-1 = "B  0%"
   *     cascade   drop-label-1 = "Other: 1,000 (100.0%)"
   *
   * Cascade is the worst again, and for the same reason: it dropped the stage,
   * then DERIVED a total wipe-out from the clamp and captioned it — on a sheet
   * whose only fault was a sign.
   *
   * These are counts. A negative stage is not a small conversion or a complete
   * loss; it is a number no share can be taken of, and the rule is the null
   * one: unknown in, unknown out. Geometry may still clamp — a band cannot have
   * negative width. Text may not.
   */
  describe("and neither is a value no share can be taken of", () => {
    const withNegative = (kind: string, values: (number | null)[] = [1000, -50, 250]): SceneNode[] =>
      buildChart({
        kind,
        ...DEFAULT_SIZE,
        data: { categories: ["Visits", "Signups", "Paid"], series: [{ name: "S", values }] },
      } as unknown as ChartConfig).nodes;

    it("states no funnel conversion across a negative stage", () => {
      const ns = withNegative("funnel");
      const conv = texts(ns, /^conv/).join(" ");
      expect(conv, `stated a conversion across a negative stage: ${conv}`).not.toMatch(/%/);
      // …and the stage still shows the number the user actually typed, which is
      // what tells them the sheet is wrong.
      expect(texts(ns, /^stage-value-/), "hid the negative instead of showing it").toContain("-50");
    });

    it("claims no cascade drop from a negative stage", () => {
      const drops = texts(withNegative("cascade"), /^drop-label-/).join(" ");
      expect(drops, `derived a drop from a clamped negative: ${drops}`).not.toMatch(/100\.0%/);
    });

    it("gives a negative waffle part no share of the whole", () => {
      const legend = texts(withNegative("waffle", [60, -20, 40]), /^legend-label-/);
      // "Signups" is the middle category `withNegative` uses — the negative one.
      expect(
        legend.find((t) => t.startsWith("Signups")),
        "a negative part was given a share",
      ).toBe("Signups");
      // The parts that DO have a share keep it — the fix must not silence the
      // legend wholesale.
      expect(legend.filter((t) => /%/.test(t)).length).toBe(2);
    });

    it("leaves an all-positive chart of each kind exactly as it was", () => {
      // The guard on the guard: every assertion above is about text NOT being
      // drawn, and deleting the labels outright would satisfy all three.
      const ok = (kind: string, values: number[]) => withNegative(kind, values);
      expect(texts(ok("funnel", [1000, 600, 250]), /^conv/).join(" ")).toMatch(/%/);
      expect(texts(ok("cascade", [1000, 600, 250]), /^stage-label-1-/).join(" ")).toMatch(/600/);
      expect(texts(ok("waffle", [60, 20, 40]), /^legend-label-/).every((t) => /%/.test(t))).toBe(true);
    });
  });
});
