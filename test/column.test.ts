import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { textWidth } from "../src/core/scene";
import { sampleConfig } from "../src/core/samples";
import type { EllipseNode, LineNode, RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Columns — clustered/stacked, gap width, overlap, stacked-100 negatives, bar styles. */

/** Regression tests for the deferred combo / stacked100 / small-multiples fixes. */
const hasNaN = (nodes: { [k: string]: unknown }[]) =>
  nodes.some((n) => Object.values(n).some((v) => typeof v === "number" && Number.isNaN(v)));

describe("bar styles on clustered", () => {
  const base: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "2024", values: [40, 55, 48] },
        { name: "2025", values: [52, 60, 45] },
      ],
    },
  };
  const styled = (barStyle: "lollipop" | "dot" | "range") =>
    buildChart({ ...base, decorations: { barStyle, segmentLabels: true } });

  it("lollipop: stem from the baseline + dot at the value", () => {
    const s = styled("lollipop");
    const stems = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("stem-"));
    const dots = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("seg-"));
    expect(stems).toHaveLength(6);
    expect(dots).toHaveLength(6);
    // Stem ends at the dot; no rectangles drawn for the data.
    expect(stems[0].y2).toBeCloseTo(dots[0].cy, 5);
    expect(s.nodes.some((n) => n.kind === "rect" && n.name?.startsWith("seg-"))).toBe(false);
  });

  it("dot: dots only, no stems", () => {
    const s = styled("dot");
    expect(s.nodes.some((n) => n.name?.startsWith("stem-"))).toBe(false);
    expect(s.nodes.filter((n) => n.kind === "ellipse" && n.name?.startsWith("seg-"))).toHaveLength(6);
  });

  it("range: two series' dots joined by a connector on one shared x", () => {
    const s = styled("range");
    const ranges = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("range-"));
    expect(ranges).toHaveLength(3);
    const d0 = s.nodes.find((n) => n.name === "seg-0-0") as EllipseNode;
    const d1 = s.nodes.find((n) => n.name === "seg-1-0") as EllipseNode;
    expect(d0.cx).toBeCloseTo(d1.cx, 5); // dumbbell: same x per category
    expect(ranges[0].y1).toBeCloseTo(d0.cy, 5);
    expect(ranges[0].y2).toBeCloseTo(d1.cy, 5);
  });

  it("stays plain bars by default and on stacked charts", () => {
    expect(buildChart(base).nodes.some((n) => n.kind === "rect" && n.name === "seg-0-0")).toBe(true);
    const stacked = buildChart({
      ...base,
      kind: "stacked",
      decorations: { barStyle: "lollipop", segmentLabels: true },
    });
    expect(stacked.nodes.some((n) => n.name?.startsWith("stem-"))).toBe(false);
  });
});

describe("column gap width", () => {
  const base: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "v", values: [5, 8, 3] }] },
    decorations: { segmentLabels: false },
  };

  it("gapWidth 0 makes columns touch (1.5× the default width)", () => {
    const def = (buildChart(base).nodes.find((n) => n.name === "seg-0-0") as RectNode).w;
    const touch = (buildChart({ ...base, gapWidth: 0 }).nodes.find((n) => n.name === "seg-0-0") as RectNode).w;
    // Default gapWidth 50 → colThick = slot·2/3; gapWidth 0 → colThick = slot.
    expect(touch / def).toBeCloseTo(1.5, 2);
  });

  it("large gapWidth thins the columns", () => {
    const def = (buildChart(base).nodes.find((n) => n.name === "seg-0-0") as RectNode).w;
    const thin = (buildChart({ ...base, gapWidth: 300 }).nodes.find((n) => n.name === "seg-0-0") as RectNode).w;
    expect(thin).toBeLessThan(def);
  });
});

describe("clustered overlap", () => {
  const base: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "s1", values: [5, 6] },
        { name: "s2", values: [7, 8] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  const stride = (cfg: ChartConfig) => {
    const s = buildChart(cfg);
    const a = s.nodes.find((n) => n.name === "seg-0-0") as RectNode;
    const b = s.nodes.find((n) => n.name === "seg-1-0") as RectNode;
    return { d: b.x - a.x, w: a.w, ax: a.x, bx: b.x };
  };

  it("overlap 0 reproduces the historical edge-to-edge layout", () => {
    const r = stride(base);
    // Two bars filling the column: stride equals a bar's full width (w+gap).
    expect(r.d).toBeGreaterThan(0);
    expect(r.d).toBeCloseTo(r.w + 1, 5);
  });

  it("positive overlap widens bars and shrinks the stride; 100 fully overlaps", () => {
    const zero = stride(base);
    const forty = stride({ ...base, overlap: 40 });
    expect(forty.w).toBeGreaterThan(zero.w);
    expect(forty.d).toBeLessThan(zero.d);
    const full = stride({ ...base, overlap: 100 });
    expect(full.ax).toBeCloseTo(full.bx, 5); // same position
  });

  it("negative overlap opens a gap between bars", () => {
    const zero = stride(base);
    const neg = stride({ ...base, overlap: -50 });
    expect(neg.d).toBeGreaterThan(zero.d);
  });
});

describe("stacked100 negative values", () => {
  const cfg: ChartConfig = {
    kind: "stacked100",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "New", values: [60, 70] },
        { name: "Renewal", values: [40, 45] },
        { name: "Returns", values: [-15, -10] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("renders the negative series below the zero line (not clamped away)", () => {
    const s = buildChart(cfg);
    const seg = (nm: string) => s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === nm)!;
    const returns = seg("seg-2-0");
    expect(returns).toBeTruthy(); // the negative segment is drawn
    const newSeg = seg("seg-0-0"); // bottom of the positive stack, sits at the zero line
    // Returns starts at/below the zero line and extends further down.
    expect(returns.y).toBeGreaterThanOrEqual(newSeg.y + newSeg.h - 1);
    expect(returns.y + returns.h).toBeGreaterThan(newSeg.y + newSeg.h + 1);
  });

  it("positive-only stacked100 is unchanged (fills exactly to the top)", () => {
    const pos = buildChart({
      ...cfg,
      data: {
        categories: ["Q1"],
        series: [
          { name: "A", values: [60] },
          { name: "B", values: [40] },
        ],
      },
    });
    // Two segments, no negative region.
    expect(pos.nodes.filter((n) => n.name?.match(/^seg-\d+-0$/))).toHaveLength(2);
  });
});

describe("stacked100 with an all-negative category", () => {
  it("fills the segments downward instead of collapsing to zero", () => {
    const cfg: ChartConfig = {
      kind: "stacked100",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A"],
        series: [
          { name: "P", values: [-30] },
          { name: "Q", values: [-20] },
        ],
      },
    };
    const scene = buildChart(cfg);
    const segs = scene.nodes.filter(
      (n) => n.kind === "rect" && (n.name ?? "").startsWith("seg") && (n as { h: number }).h > 0.5,
    );
    expect(segs.length).toBe(2); // both shares visible (were 0-height before)
    expect(hasNaN(scene.nodes as never)).toBe(false);
  });
});

describe("clustered level-anchored decorations (regression)", () => {
  // seriesLevels only advanced in the STACKED branch, so on a clustered chart it
  // was an array of zeros — while decor.ts still opted into "stack level" mode
  // because the array existed. The difference arrow silently collapsed onto the
  // baseline and read "0".
  const base = {
    width: 480,
    height: 300,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "S1", values: [100, 150] },
        { name: "S2", values: [50, 50] },
      ],
    },
    decorations: { difference: { from: 0, to: 1, series: 0 } },
  };

  it.each(["stacked", "clustered"])("%s anchors the arrow to the series mark, not the axis", (kind) => {
    const s = buildChart({ ...base, kind } as unknown as ChartConfig);
    const label = s.nodes.find((n) => n.name?.startsWith("diff-label")) as { text?: string } | undefined;
    const line = s.nodes.find((n) => n.name?.startsWith("diff-line")) as { y1?: number; y2?: number } | undefined;
    expect(label?.text, `${kind} difference label`).toBe("+50%");
    expect(line, `${kind} difference line`).toBeTruthy();
    expect(Math.abs(line!.y1! - line!.y2!), `${kind} arrow collapsed to zero length`).toBeGreaterThan(1);
  });
});

describe("stacked100 value axis (regression)", () => {
  const cfg = (negatives: boolean): ChartConfig =>
    ({
      kind: "stacked100",
      width: 480,
      height: 300,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "New", values: [60, 70] },
          { name: "Renewal", values: [40, 45] },
          ...(negatives ? [{ name: "Returns", values: [-40, -10] }] : []),
        ],
      },
      decorations: { valueAxis: true, gridlines: true },
    }) as unknown as ChartConfig;

  it.each([false, true])("keeps every tick and gridline on the canvas (negatives=%s)", (negatives) => {
    const s = buildChart(cfg(negatives));
    const offCanvas = s.nodes
      .map((n) => ({ n, y: (n as { y?: number }).y ?? (n as { y1?: number }).y1 }))
      .filter(({ y }) => typeof y === "number" && (y < -0.5 || y > s.height + 0.5))
      .map(({ n, y }) => `${n.name}@${y!.toFixed(1)}`);
    // niceTicks rounds OUTWARD from pctNegMin, and nothing filtered it back into
    // the domain — a gridline landed at y=301.6 on a 300pt canvas.
    expect(offCanvas, `off-canvas: ${offCanvas.join(", ")}`).toEqual([]);
  });

  it("labels the share axis in percent, each label naming its own tick", () => {
    const labels = (buildChart(cfg(false)).nodes as { name?: string; text?: string }[])
      .filter((n) => n.name === "value-axis")
      .map((n) => n.text);
    // Was 0.0 / 0.3 / 0.5 / 0.8 / 1.0 — fractions on a chart whose segments read
    // "60%", and 0.25/0.75 rounded to labels that name no tick on the axis.
    expect(labels).toEqual(["0%", "25%", "50%", "75%", "100%"]);
  });

  it.each([false, true])("keeps the unit suffix off the share axis (horizontal=%s)", (horizontal) => {
    // A share is unitless. numberFormat.suffix is how an author says "millions"
    // — routed through the axis it labelled the 100% strip "25 m%", while the
    // segments beside it correctly read "25%".
    const labels = (
      buildChart({
        ...cfg(false),
        horizontal,
        numberFormat: { suffix: " m", decimals: "auto" },
      } as unknown as ChartConfig).nodes as { name?: string; text?: string }[]
    )
      .filter((n) => n.name === "value-axis")
      .map((n) => n.text);
    expect(labels).toEqual(["0%", "25%", "50%", "75%", "100%"]);
  });
});

/**
 * An authored "100% =" denominator can be SMALLER than the column it normalises
 * (a multi-select survey — "100% = 500 respondents", answers totalling 180% — or
 * a typo). The 100% scale was hard-pinned to max: 1, so those segments were
 * painted outside the plot AND outside the chart frame. `waffle` already handles
 * the same input (test/spatial-layouts.test.ts).
 */
describe("stacked100 honors an inconsistent 100%= denominator", () => {
  const s = buildChart({
    kind: "stacked100",
    ...DEFAULT_SIZE,
    width: 480,
    height: 300,
    data: {
      categories: ["A"],
      series: [
        { name: "x", values: [300] },
        { name: "y", values: [600] },
      ],
      hundredPercent: [500],
    },
  } as unknown as ChartConfig);

  it("keeps the over-100% stack inside the canvas", () => {
    const off = (s.nodes as { name?: string; x?: number; y?: number; w?: number; h?: number }[])
      .filter((n) => n.name?.startsWith("seg-"))
      .filter((n) => (n.y ?? 0) < -0.5 || (n.y ?? 0) + (n.h ?? 0) > s.height + 0.5)
      .map((n) => `${n.name}@${n.y?.toFixed(1)}`);
    expect(off, `off-canvas: ${off.join(", ")}`).toEqual([]);
  });

  it("keeps the segments proportional to their shares", () => {
    const seg = (name: string) => (s.nodes as { name?: string; h?: number }[]).find((n) => n.name === name)!;
    // 600 : 300 — the taller segment is exactly twice the shorter one.
    expect(seg("seg-1-0").h! / seg("seg-0-0").h!).toBeCloseTo(2, 1);
  });

  it("leaves a well-formed 100% chart on the plain 0/25/50/75/100 strip", () => {
    const ok = buildChart({
      kind: "stacked100",
      ...DEFAULT_SIZE,
      width: 480,
      height: 300,
      data: {
        categories: ["A"],
        series: [
          { name: "x", values: [3] },
          { name: "y", values: [7] },
        ],
      },
      decorations: { valueAxis: true },
    } as unknown as ChartConfig);
    const labels = (ok.nodes as { name?: string; text?: string }[])
      .filter((n) => n.name === "value-axis")
      .map((n) => n.text);
    expect(labels).toEqual(["0%", "25%", "50%", "75%", "100%"]);
  });
});

/**
 * A decoration's pixel anchor and the number it prints have to describe the
 * same mark. `columnTop` is the DRAWN top and `columnValue` is what the label
 * reads, and on a stacked column with a negative segment they are different
 * quantities: the positive total against the net. Two shapes came out of that —
 * a zero-length arrow carrying a non-zero percentage, and an arrow pointing the
 * opposite way to its own sign.
 */
describe("a difference arrow on a mixed-sign stack points where its number says", () => {
  const stack = (a: [number, number], b: [number, number]): ChartConfig =>
    ({
      kind: "stacked",
      width: 400,
      height: 300,
      decorations: { difference: { from: 0, to: 1 } },
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "P", values: [a[0], b[0]] },
          { name: "N", values: [a[1], b[1]] },
        ],
      },
    }) as unknown as ChartConfig;

  const arrow = (cfg: ChartConfig) => {
    const nodes = buildChart(cfg).nodes;
    const line = nodes.find((n): n is LineNode => n.kind === "line" && n.name === "diff-line")!;
    const label = nodes.find((n): n is TextNode => n.kind === "text" && n.name === "diff-label")!;
    return { line, label };
  };

  it("draws a real length when the totals differ", () => {
    // +10/−4 → +10/−8 is a net 6 → 2. Both POSITIVE totals are 10, so anchoring
    // on the drawn top gave a zero-length arrow labelled "−67%".
    const { line, label } = arrow(stack([10, -4], [10, -8]));
    expect(label.text).toContain("-");
    expect(Math.abs(line.y2 - line.y1), "zero-length arrow").toBeGreaterThan(1);
    expect(line.y2).toBeGreaterThan(line.y1); // a fall goes DOWN the canvas
  });

  it("points up for a rise and down for a fall", () => {
    // +10/−8 → +5/0 is a net 2 → 5, a RISE — drawn pointing down before.
    const { line, label } = arrow(stack([10, -8], [5, 0]));
    expect(label.text).toContain("+");
    expect(line.y2).toBeLessThan(line.y1);
  });

  it("leaves an all-positive stack exactly where it was", () => {
    // The negative control: the two anchors agree whenever no segment is
    // negative, so this must be an identity — as it is for every chart in the
    // showcase deck.
    const { line, label } = arrow(stack([10, 5], [20, 6]));
    expect(label.text).toContain("+");
    expect(line.y2).toBeLessThan(line.y1);
  });
});

/**
 * A clustered category whose bars are ALL negative.
 *
 * `drawnTopValue` seeded its max with 0, so such a category published 0 —
 * and `columnValue` is what every value-reading decoration prints. A cash-flow
 * chart of [-40, -20, -10, 60] drew its mean line ABOVE the baseline and
 * labelled it "Ø 15", the mean of [0, 0, 0, 60], while the same data as
 * `stacked`, `line` or `area` correctly said "Ø -2.5". A difference arrow
 * between two all-negative categories came out zero-length and labelled "0",
 * sitting beside the very totals that contradict it.
 *
 * Zero is right for the drawn column's TOP — a bar below the baseline tops out
 * at the baseline — and wrong for its VALUE, which is what the reader is told
 * the category is worth.
 */
describe("a clustered column that only goes down", () => {
  const meanLabel = (kind: string) =>
    (
      buildChart({
        kind,
        width: 520,
        height: 320,
        data: { categories: ["Q1", "Q2", "Q3", "Q4"], series: [{ name: "Net", values: [-40, -20, -10, 60] }] },
        decorations: { categoryAxis: true, valueAxis: true, valueLines: [{ mode: "mean" }] },
      } as unknown as ChartConfig).nodes.find((n) => n.name === "value-line-label-0") as { text?: string } | undefined
    )?.text;

  it("averages what the bars actually say, like every other kind does", () => {
    // The true mean of [-40, -20, -10, 60] is -2.5.
    expect(meanLabel("clustered"), "clustered disagreed with the data").toBe("Ø -2.5");
    // The negative control: the kinds that were already right must stay right.
    for (const k of ["stacked", "line", "area"]) expect(meanLabel(k), k).toBe("Ø -2.5");
  });

  it("draws a difference arrow with a length and a true label", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 520,
      height: 320,
      data: { categories: ["FY22", "FY23", "FY24"], series: [{ name: "Movement", values: [-120, -80, -30] }] },
      decorations: { categoryAxis: true, valueAxis: true, totals: true, difference: { from: 0, to: 2 } },
    } as unknown as ChartConfig);
    const label = scene.nodes.find((n) => n.name === "diff-label") as { text?: string } | undefined;
    const line = scene.nodes.find((n) => n.name === "diff-line") as { y1: number; y2: number } | undefined;
    expect(label?.text, "the arrow labelled an all-negative movement as zero").not.toBe("0");
    expect(line && Math.abs(line.y1 - line.y2), "the arrow had no length").toBeGreaterThan(1);
  });

  it("is unchanged for data that rises", () => {
    // maxima 20, 30, 25 -> mean 25. The old code agreed here and must still.
    const scene = buildChart({
      kind: "clustered",
      width: 520,
      height: 320,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "S", values: [20, 30, 25] },
          { name: "T", values: [18, 22, 24] },
        ],
      },
      decorations: { categoryAxis: true, valueAxis: true, valueLines: [{ mode: "mean" }] },
    } as unknown as ChartConfig);
    expect((scene.nodes.find((n) => n.name === "value-line-label-0") as { text?: string } | undefined)?.text).toBe(
      "Ø 25",
    );
  });
});

/**
 * A column total is centred on its category SLOT and was drawn at the chart font
 * however narrow that slot is — the one label family here that never scaled
 * while the title, the category names and the series labels all did. `"113"` is
 * 17.4pt at the default font in a 13pt slot at 80x60, so the last two totals
 * were drawn through each other.
 *
 * The horizontal branch already bounds itself against its row pitch; this is the
 * same bound for the upright chart, and the same two-step the clustered in-bar
 * labels take: fit to the slot, drop past the floor.
 */
describe("upright column totals are fitted to their slot", () => {
  const totals = (w: number, h: number) =>
    buildChart({ ...sampleConfig("stacked"), width: w, height: h } as ChartConfig).nodes.filter(
      (n): n is TextNode => n.kind === "text" && /^total-\d+$/.test(n.name ?? ""),
    );

  it("no two adjacent totals overlap on a thumbnail", () => {
    const ts = totals(80, 60);
    expect(ts.length, "no totals drawn — the check would be vacuous").toBeGreaterThan(1);
    for (let i = 1; i < ts.length; i++) {
      const a = ts[i - 1];
      const b = ts[i];
      const inkA = textWidth(a.text, a.fontSize, a.bold);
      const inkB = textWidth(b.text, b.fontSize, b.bold);
      const ax1 = a.x + (a.w + inkA) / 2;
      const bx0 = b.x + (b.w - inkB) / 2;
      expect(ax1, `${a.name} runs into ${b.name}`).toBeLessThanOrEqual(bx0 + 0.01);
    }
  });

  it("leaves an ordinary chart's totals at the chart font", () => {
    // Last resort: a total that already fits its slot keeps `fs` exactly, so no
    // ordinary chart moves and the showcase deck does not shift.
    for (const t of totals(480, 300)) expect(t.fontSize).toBeCloseTo(10, 5);
  });
});

describe("a drawn segment's percentage matches the segment", () => {
  it("labels a negative stacked segment with its real share, not 0%", () => {
    /**
     * `Math.max(0, raw)` sent every negative segment's fraction to zero, so a
     * returns row of -25 against a positive total of 100 was DRAWN at a quarter
     * of the stack's height and labelled "0%". The picture said one thing and
     * the caption said the outflow never happened — and the caption is the half
     * a reader quotes.
     *
     * The clamp was right where it came from: a negative must not shorten the
     * positive stack. It had no business reaching a LABEL.
     */
    const scene = buildChart({
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1"],
        series: [
          { name: "New", values: [60] },
          { name: "Renewal", values: [40] },
          { name: "Returns", values: [-25] },
        ],
      },
      decorations: { segmentLabels: true, labelContent: ["percent"] },
    } as unknown as ChartConfig);
    const label = (nm: string) =>
      String(scene.nodes.find((n): n is TextNode => n.kind === "text" && n.name === nm)?.text ?? "");
    expect(label("label-0-0")).toBe("60%");
    expect(label("label-1-0")).toBe("40%");
    expect(label("label-2-0"), "a drawn outflow was labelled as nothing").toBe("-25%");
    // And it is still DRAWN — the fix must not have quietly removed the bar the
    // label is about.
    const seg = scene.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "seg-2-0");
    expect(seg?.h ?? 0, "the negative segment stopped being drawn").toBeGreaterThan(1);
  });
});

describe("the difference arrow's caption agrees with the arrow", () => {
  /**
   * The arrow and its own label come from the same two values, and on a
   * NEGATIVE base they disagreed — the arrow being the half that was right.
   * `vTo / vFrom - 1` inverts its sign under a negative denominator, so a loss
   * halving from -100 to -50 pointed UP and said "-50%", and a swing from -50
   * into 25 of profit pointed UP and said "-150%".
   *
   * A reader who trusts the number over the picture then reads the opposite of
   * what happened, which is the worst outcome available on a slide.
   */
  const caption = (from: number, to: number) => {
    const scene = buildChart({
      kind: "clustered",
      ...DEFAULT_SIZE,
      data: { categories: ["FY23", "FY24"], series: [{ name: "Margin", values: [from, to] }] },
      decorations: { difference: { from: 0, to: 1, percent: true } },
    } as unknown as ChartConfig);
    const label = scene.nodes.find((n): n is TextNode => n.kind === "text" && n.name === "diff-label")!;
    const head = scene.nodes.find((n) => n.kind === "arrowhead" && n.name === "diff-head") as
      { angle: number } | undefined;
    expect(label, `no caption for ${from} -> ${to}`).toBeTruthy();
    expect(head, `no arrow for ${from} -> ${to}`).toBeTruthy();
    return { text: String(label.text), points: head!.angle === -90 ? "up" : "down" };
  };

  it("never labels an upward arrow with a fall, or the reverse", () => {
    for (const [from, to] of [
      [100, 150],
      [100, 50],
      [-100, -50],
      [-100, -200],
      [-50, 25],
      [-50, -75],
    ]) {
      const { text, points } = caption(from, to);
      const rose = to > from;
      expect(points, `${from} -> ${to}: the arrow itself points the wrong way`).toBe(rose ? "up" : "down");
      // The caption's SIGN is the claim a reader takes away from it.
      expect(text.startsWith("-"), `${from} -> ${to}: caption "${text}" contradicts an arrow pointing ${points}`).toBe(
        !rose,
      );
    }
  });

  it("still gives a percentage when the base is positive", () => {
    // The fix falls back to an absolute difference, and it must not do that to
    // the ordinary case — a percentage is what this decoration is FOR.
    expect(caption(100, 150).text).toContain("%");
    expect(caption(100, 50).text).toContain("%");
    // …and drops the percentage exactly where it stops meaning anything.
    expect(caption(-100, -50).text).not.toContain("%");
    expect(caption(-50, 25).text).not.toContain("%");
  });
});
