import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { layoutColumns } from "../src/core/layout/column";
import { layoutMekko } from "../src/core/layout/mekko";
import { layoutButterfly } from "../src/core/layout/butterfly";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";
import { textWidth } from "../src/core/scene";
import { dataToSheet, evaluateFormula, sheetToData, transposeSheet } from "../src/taskpane/datasheet";

const rects = (nodes: { kind: string }[]) => nodes.filter((n): n is RectNode => n.kind === "rect");
const byName = (nodes: { name?: string }[], prefix: string) => nodes.filter((n) => n.name?.startsWith(prefix));

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

describe("horizontal (bar) orientation", () => {
  const c = cfg({
    horizontal: true,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "S1", values: [10, 20] },
        { name: "S2", values: [5, 15] },
      ],
    },
  });
  const { nodes, anchors } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("bars extend along x and stack contiguously", () => {
    const segs = byName(rects(nodes), "seg-") as RectNode[];
    expect(segs).toHaveLength(4);
    const colA = segs.filter((s) => s.name!.endsWith("-0"));
    const [s1, s2] = colA;
    // Same row (y), s2 starts where s1 ends.
    expect(s1.y).toBeCloseTo(s2.y, 5);
    expect(s2.x).toBeCloseTo(s1.x + s1.w, 5);
  });

  it("bar lengths are proportional to totals", () => {
    const lenA = anchors.columnTop[0] - anchors.baselineY; // x extent, total 15
    const lenB = anchors.columnTop[1] - anchors.baselineY; // total 35
    expect(lenB / lenA).toBeCloseTo(35 / 15, 5);
  });

  it("builds through buildChart without decorations blowing up", () => {
    const scene = buildChart({ ...c, decorations: { cagr: { from: 0, to: 1 }, totals: true } });
    // CAGR must be skipped in horizontal orientation.
    expect(scene.nodes.find((n) => n.name === "cagr-label")).toBeUndefined();
  });
});

describe("butterfly", () => {
  const c = cfg({
    kind: "butterfly",
    data: {
      categories: ["A", "B"],
      series: [
        { name: "L", values: [10, 20] },
        { name: "R", values: [20, 40] },
      ],
    },
  });
  const { nodes } = layoutButterfly(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("left series extends left, right series extends right, shared scale", () => {
    const segs = byName(rects(nodes), "seg-") as RectNode[];
    const l = segs.find((s) => s.name === "seg-0-1")!; // L=20
    const r = segs.find((s) => s.name === "seg-1-1")!; // R=40
    expect(r.w / l.w).toBeCloseTo(2, 3);
    expect(l.x + l.w).toBeLessThanOrEqual(r.x); // separated by the gutter
  });

  it("puts category labels in the center gutter", () => {
    const cats = byName(nodes, "category-") as TextNode[];
    expect(cats).toHaveLength(2);
  });
});

describe("100%= row", () => {
  it("leaves columns short of full height when series sum < denominator", () => {
    const c = cfg({
      kind: "stacked100",
      data: {
        categories: ["A"],
        series: [{ name: "S", values: [50] }],
        hundredPercent: [100],
      },
    });
    const { nodes, anchors } = layoutColumns(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const seg = byName(rects(nodes), "seg-")[0] as RectNode;
    expect(seg.h).toBeCloseTo(anchors.plot.h / 2, 1);
    const label = byName(nodes, "label-")[0] as TextNode;
    expect(label.text).toBe("50%");
  });
});

describe("mekko category labels do not run into each other", () => {
  /**
   * A Mekko's category label sits under its own COLUMN and is as wide as it
   * needs to be, so a label wider than its column overflows the box
   * symmetrically — into its neighbour, and off the left edge of the chart for
   * the first one. At a 22pt font the three-category sample had two overlapping
   * pairs and `category-0` starting at x = -2, which reads on screen as one
   * run-on string: "EMEA (32%)Americas (42%)APAC (27%)". Rendered and looked at.
   */
  const inkSpan = (t: TextNode) => {
    const w = Math.min(t.w, textWidth(t.text, t.fontSize, t.bold));
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    return { x0: x, x1: x + w };
  };
  const cats = (fontSize: number) => {
    const scene = buildChart({ ...sampleConfig("mekko"), style: { fontSize } } as ChartConfig);
    const found = scene.nodes
      .filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("category-"))
      .map((t) => ({ node: t, ...inkSpan(t) }))
      .sort((a, b) => a.x0 - b.x0);
    return { scene, found };
  };

  it("keeps them apart and on the chart at any font size", () => {
    for (const fs of [10, 16, 22, 30]) {
      const { scene, found } = cats(fs);
      expect(found.length, `fs=${fs} drew no category labels`).toBe(3);
      for (let i = 1; i < found.length; i++)
        expect(found[i].x0, `fs=${fs} ${found[i].node.name} overlaps ${found[i - 1].node.name}`).toBeGreaterThanOrEqual(
          found[i - 1].x1 - 0.5,
        );
      for (const f of found) {
        expect(f.x0, `fs=${fs} ${f.node.name} started left of the chart`).toBeGreaterThanOrEqual(-0.5);
        expect(f.x1, `fs=${fs} ${f.node.name} ran past the right edge`).toBeLessThanOrEqual(scene.width + 0.5);
      }
    }
  });

  it("leaves an ordinary chart's labels at the chart's own font size", () => {
    // The shrink must be a last resort, not a permanent tax: at the default font
    // the labels fit, so nothing moves.
    for (const t of cats(10).found) expect(t.node.fontSize).toBe(10);
    // …and they stay whole rather than being ellipsized to fit.
    for (const t of cats(22).found) expect(t.node.text.endsWith("…")).toBe(false);
    // One size for all of them — labels that differ in size read as a hierarchy
    // that is not there.
    expect(new Set(cats(22).found.map((t) => t.node.fontSize)).size).toBe(1);
  });
});

describe("mekko with units (X extent)", () => {
  it("takes widths from xExtent and heights from totals", () => {
    const c = cfg({
      kind: "mekko",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [10, 5] },
          { name: "S2", values: [10, 5] },
        ],
        xExtent: [30, 10],
      },
    });
    const { anchors } = layoutMekko(c, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(anchors.categoryWidth[0] / anchors.categoryWidth[1]).toBeCloseTo(3, 3);
    // Column B total (10) is half of A (20) → its top sits lower.
    const hA = anchors.plot.y + anchors.plot.h - anchors.columnTop[0];
    const hB = anchors.plot.y + anchors.plot.h - anchors.columnTop[1];
    expect(hB / hA).toBeCloseTo(0.5, 3);
  });
});

describe("level difference arrow", () => {
  it("compares cumulative stack levels of the given series", () => {
    const c = cfg({
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [100, 150] },
          { name: "S2", values: [50, 50] },
        ],
      },
      decorations: { difference: { from: 0, to: 1, series: 0 } },
    });
    const scene = buildChart(c);
    const label = scene.nodes.find((n) => n.name === "diff-label") as TextNode;
    expect(label.text).toBe("+50%"); // 150 vs 100 at series 0's level
  });
});

describe("multiple value lines", () => {
  it("draws one line per entry", () => {
    const c = cfg({
      data: { categories: ["A", "B"], series: [{ name: "S", values: [50, 100] }] },
      decorations: { valueLines: [{ mode: "mean" }, { mode: "value", value: 90 }] },
    });
    const scene = buildChart(c);
    expect(byName(scene.nodes, "value-line-label")).toHaveLength(2);
    expect(scene.nodes.filter((n) => n.kind === "line" && n.name?.startsWith("value-line-"))).toHaveLength(2);
  });

  it("normalizes the legacy single valueLine", () => {
    const c = cfg({
      data: { categories: ["A"], series: [{ name: "S", values: [50] }] },
      decorations: { valueLine: { mode: "value", value: 25 } },
    });
    const scene = buildChart(c);
    expect(byName(scene.nodes, "value-line-label")).toHaveLength(1);
  });
});

describe("datasheet special rows and transpose", () => {
  it("round-trips 100%= and X extent rows", () => {
    const data = {
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] as (number | null)[] }],
      hundredPercent: [10, 20] as (number | null)[],
      xExtent: [3, 4] as (number | null)[],
    };
    const back = sheetToData(dataToSheet(data));
    expect(back.hundredPercent).toEqual([10, 20]);
    expect(back.xExtent).toEqual([3, 4]);
    expect(back.series).toHaveLength(1);
    expect(back.series[0].values).toEqual([1, 2]);
  });

  it("keeps a scatter/bubble X series out of xExtent on a round trip", () => {
    // A lenient X-extent row pattern swallowed the scatter/bubble "X" SERIES into
    // the Mekko-only xExtent field, so every pane edit silently destroyed the X
    // data (layout/scatter.ts finds its X via /^x$/i) and renamed the visible row
    // to "X extent". Kept beside the Mekko case above so neither can drift.
    for (const kind of ["scatter", "bubble"] as const) {
      const back = sheetToData(dataToSheet(sampleConfig(kind).data));
      expect(
        back.series.map((s) => s.name),
        kind,
      ).toContain("X");
      expect(back.xExtent, kind).toBeUndefined();
    }
  });

  it("transposes rows and columns", () => {
    const t = transposeSheet({
      cells: [
        ["", "A", "B"],
        ["S1", "1", "2"],
        ["S2", "3", "4"],
      ],
    });
    expect(t.cells).toEqual([
      ["", "S1", "S2"],
      ["A", "1", "3"],
      ["B", "2", "4"],
    ]);
    const data = sheetToData(t);
    expect(data.categories).toEqual(["S1", "S2"]);
    expect(data.series.map((s) => s.name)).toEqual(["A", "B"]);
  });

  it("remaps A1 references on transpose, so formula cells keep their numbers", () => {
    // The formula TEXT used to move verbatim, so "=SUM(B2:B3)" summed a whole
    // different pair of cells after the swap: the Total row went from 40/60 to
    // 30/70. A transpose swaps the axes, it must not change the values.
    const t = transposeSheet({
      cells: [
        ["", "Q1", "Q2"],
        ["North", "10", "20"],
        ["South", "30", "40"],
        ["Total", "=SUM(B2:B3)", "=SUM(C2:C3)"],
      ],
    });
    expect(t.cells[1]).toEqual(["Q1", "10", "30", "=SUM(B2:C2)"]);
    expect(sheetToData(t).series.map((s) => s.values)).toEqual([
      [10, 30, 40],
      [20, 40, 60],
    ]);
    // Transposing back is the identity, formulas included.
    expect(transposeSheet(t).cells[3]).toEqual(["Total", "=SUM(B2:B3)", "=SUM(C2:C3)"]);
  });

  it("a whole-column sum keeps its number across a transpose", () => {
    // `=SUM(B2:B999)` is what an Excel user types to total a column without
    // counting its rows, and this engine evaluates it correctly (the range
    // clamps to the grid). Transposing it produced `=SUM(B2:ALK2)` — a THREE
    // letter column, which no pattern in the evaluator could match — so the
    // formula silently evaluated to nothing and the Total column went blank.
    //
    // The cause was that the grid's two coordinate spaces were different sizes:
    // rows ran to 999 and columns stopped at ZZ, so a row above 702 had no
    // column to become. It is not enough for a reference to MOVE correctly; the
    // destination has to be sayable.
    const sheet = {
      cells: [
        ["", "Q1", "Q2", "Total"],
        ["North", "10", "20", "=SUM(B2:B999)"],
        ["South", "30", "40", "=SUM(B3:B999)"],
      ],
    };
    // The whole B column: 10 + 30. Holds either side of the fix — it is the
    // premise, not the guard.
    expect(evaluateFormula(sheet.cells, "SUM(B2:B999)")).toBe(40);
    const t = transposeSheet(sheet);
    const moved = t.cells[3][1];
    expect(moved).toBe("=SUM(B2:ALK2)");
    // The guard: the same number, out of the moved formula.
    expect(evaluateFormula(t.cells, moved.slice(1))).toBe(40);
    // …and it still says the same thing after a round trip.
    expect(transposeSheet(t).cells[1][3]).toBe("=SUM(B2:B999)");
  });

  it("leaves a reference it cannot express alone rather than moving part of it", () => {
    // The reference pattern used to match a PREFIX of a longer token. `A1000`
    // matched as "A100", which the evaluator choked on (leftover "0") and
    // answered null for — while the transposer rewrote just that prefix and left
    // the digit behind, so `=A1000` came back as `=CV10`: a real cell holding a
    // real number where the user had a blank. Wrong in a worse direction than
    // blank is.
    const grid = [
      ["", "a", "b"],
      ["S", "7", "8"],
      ["T", "9", "1"],
    ];
    for (const f of ["=A1000", "=ZZZ1", "=SUM(A1:ZZZ9)"]) {
      expect(transposeSheet({ cells: [[f]] }).cells[0][0], `${f} was rewritten`).toBe(f);
      expect(evaluateFormula(grid, f.slice(1)), `${f} became readable`).toBeNull();
    }
    // A reference INSIDE the address space still moves, so the guard above is
    // not simply "never rewrite anything".
    expect(transposeSheet({ cells: [["=B999"]] }).cells[0][0]).toBe("=ALK2");
  });
});

describe("grand total label (think-cell 14)", () => {
  const base = cfg({
    kind: "stacked",
    data: {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "A", values: [10, 20, 30] },
        { name: "B", values: [5, 10, 15] },
      ],
    },
  });
  const grand = (nodes: { name?: string }[]) => nodes.find((n) => n.name === "grand-total") as TextNode | undefined;

  it("emits one top-right label summing every category total", () => {
    const { nodes } = layoutColumns(base, DEFAULT_STYLE, { ...DEFAULT_DECOR, totals: true, grandTotal: true });
    const g = grand(nodes);
    expect(g).toBeTruthy();
    expect(g!.text).toBe("90"); // 15 + 30 + 45
    // Top-right: right-aligned, near the top of the frame, past the last column's centre.
    expect(g!.align).toBe("right");
    const totals = byName(nodes, "total-") as TextNode[];
    for (const t of totals) expect(g!.y).toBeLessThanOrEqual(t.y + 1); // grand total sits at/above the column totals row
  });

  it("is independent of per-column totals (can show without them)", () => {
    const { nodes } = layoutColumns(
      { ...base, decorations: { ...DEFAULT_DECOR, totals: false, grandTotal: true } },
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, totals: false, grandTotal: true },
    );
    expect(grand(nodes)).toBeTruthy();
    expect(byName(nodes, "total-")).toHaveLength(0); // no per-column totals
  });

  it("clustered sums across every series and category", () => {
    const { nodes } = layoutColumns(
      { ...base, kind: "clustered", decorations: { ...DEFAULT_DECOR, grandTotal: true } },
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, grandTotal: true },
    );
    expect(grand(nodes)!.text).toBe("90");
  });

  it("is suppressed on a 100% chart (every column totals the same) and when horizontal", () => {
    const pct = layoutColumns(
      { ...base, kind: "stacked100", decorations: { ...DEFAULT_DECOR, grandTotal: true } },
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, grandTotal: true },
    );
    expect(grand(pct.nodes)).toBeUndefined();
    const horiz = layoutColumns(
      { ...base, horizontal: true, decorations: { ...DEFAULT_DECOR, grandTotal: true } },
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, grandTotal: true },
    );
    expect(grand(horiz.nodes)).toBeUndefined();
  });
});

describe("IBCS scenario notation (Series.scenario)", () => {
  const seg = (nodes: { name?: string }[], si: number) =>
    nodes.find((n) => n.name === `seg-${si}-0`) as RectNode | undefined;
  const build = (scenarios: (string | undefined)[]) =>
    layoutColumns(
      cfg({
        kind: "clustered",
        data: {
          categories: ["Q1"],
          series: scenarios.map((sc, i) => ({
            name: "Sales",
            color: "#3b6ea5",
            values: [80 + i],
            ...(sc ? { scenario: sc as "AC" | "PY" | "PL" | "BU" | "FC" } : {}),
          })),
        },
      }),
      DEFAULT_STYLE,
      DEFAULT_DECOR,
    ).nodes;

  it("AC is a solid fill of the series colour", () => {
    const r = seg(build(["AC"]), 0)!;
    expect(r.fill).toBe("#3b6ea5");
    expect(r.pattern).toBeUndefined();
  });

  it("PY is a lighter solid (not the base, not white/none)", () => {
    const r = seg(build(["PY"]), 0)!;
    expect(r.fill).not.toBe("#3b6ea5");
    expect(r.fill).not.toBe("none");
    expect(r.fill).toMatch(/^#[0-9a-f]{6}$/i);
  });

  it("PL / BU are outlined/hollow: no fill, coloured border", () => {
    for (const sc of ["PL", "BU"]) {
      const r = seg(build([sc]), 0)!;
      expect(r.fill).toBe("none");
      expect(r.stroke).toBe("#3b6ea5");
      expect(r.strokeWidth).toBeGreaterThan(0);
    }
  });

  it("FC is hatched AND carries a fill/border encoding that survives export", () => {
    const r = seg(build(["FC"]), 0)!;
    // SVG shows the true IBCS hatch…
    expect(r.pattern).toBe("diagonal");
    // …but rect.pattern is SVG-only: both PowerPoint renderers drop it, which
    // made FC pixel-identical to AC in the actual deliverable. So the fill and
    // border must ALSO distinguish it, using only what all three can express.
    expect(r.fill).toMatch(/^#[0-9a-f]{6}$/i);
    expect(r.fill).not.toBe("#3b6ea5"); // not AC's solid
    expect(r.fill).not.toBe("none"); // not PL/BU's hollow
    expect(r.stroke).toBe("#3b6ea5");
    expect(r.strokeWidth).toBeGreaterThan(0);
  });

  it("keeps all four scenario codes visually distinct", () => {
    const nodes = build(["AC", "PY", "PL", "FC"]);
    const sig = [0, 1, 2, 3].map((i) => {
      const r = seg(nodes, i)!;
      return `${r.fill}|${r.stroke}|${r.strokeWidth}`;
    });
    expect(new Set(sig).size).toBe(4);
  });

  it("inks the segment label against the PAINTED fill, not the series colour", () => {
    // #168 restyles the segment (PL/BU hollow, PY/FC tinted) but the label ink
    // was still chosen from the original dark series colour — so a hollow PL bar
    // got white text on the white canvas, and a light PY tint got white on light.
    const nodes = layoutColumns(
      cfg({
        kind: "stacked",
        data: {
          categories: ["Q1"],
          series: [
            { name: "Plan", scenario: "PL" as const, color: "#1f3d7a", values: [30] },
            { name: "Prev", scenario: "PY" as const, color: "#1f3d7a", values: [30] },
          ],
        },
      }),
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, segmentLabels: true },
    ).nodes;
    for (const si of [0, 1]) {
      const label = nodes.find((n) => n.name === `label-${si}-0`) as TextNode | undefined;
      // Dark ink on a hollow bar (white canvas) and on a light tint.
      expect(label?.color, `series ${si}`).toBe("#0b0b0b");
    }
  });

  it("tints PY toward the canvas, so a dark theme does not wash it white", () => {
    // #138 moved every tint onto style.background; #168 reintroduced a hardcoded
    // "#ffffff" at a new site, which blows out on a dark canvas.
    const dark = layoutColumns(
      cfg({
        kind: "clustered",
        data: { categories: ["Q1"], series: [{ name: "S", scenario: "PY" as const, color: "#3b6ea5", values: [80] }] },
      }),
      { ...DEFAULT_STYLE, background: "#1b1b1b" },
      DEFAULT_DECOR,
    ).nodes;
    const r = seg(dark, 0)!;
    // Halfway to a near-black surface must be DARKER than the series colour,
    // never lighter (which is what lerping toward white produced).
    const lum = (c: string) => parseInt(c.slice(1, 3), 16) + parseInt(c.slice(3, 5), 16) + parseInt(c.slice(5, 7), 16);
    expect(lum(r.fill)).toBeLessThan(lum("#3b6ea5"));
  });

  it("appends the two-letter scenario code to the legend label", () => {
    const nodes = layoutColumns(
      cfg({
        kind: "clustered",
        data: { categories: ["Q1"], series: [{ name: "Sales", scenario: "AC", values: [80] }] },
      }),
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, seriesLabels: true },
    ).nodes;
    const labels = nodes.filter((n): n is TextNode => n.kind === "text").map((n) => n.text);
    expect(labels.some((t) => t.includes("Sales (AC)"))).toBe(true);
  });
});

describe("IBCS variance tier (decorations.variance)", () => {
  const base = cfg({
    kind: "clustered",
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "AC", values: [82, 91] },
        { name: "PL", values: [85, 100] },
      ],
    },
  });
  const bar = (nodes: { name?: string }[], c: number) =>
    nodes.find((n) => n.name === `variance-bar-${c}`) as RectNode | undefined;
  const label = (nodes: { name?: string }[], c: number) =>
    nodes.find((n) => n.name === `variance-label-${c}`) as TextNode | undefined;

  it("draws a signed absolute-Δ bar + label per category, plus a zero line", () => {
    const { nodes } = layoutColumns(base, DEFAULT_STYLE, {
      ...DEFAULT_DECOR,
      variance: { actual: 0, reference: 1 },
    });
    expect(nodes.find((n) => n.name === "variance-zero")).toBeTruthy();
    expect(label(nodes, 0)!.text).toBe("-3"); // 82 − 85
    expect(label(nodes, 1)!.text).toBe("-9"); // 91 − 100
    expect(bar(nodes, 0)).toBeTruthy();
  });

  it("colours a favourable delta green and an unfavourable one red (goodIsUp default)", () => {
    const up = cfg({
      kind: "clustered",
      data: {
        categories: ["Q1"],
        series: [
          { name: "AC", values: [90] },
          { name: "PL", values: [80] },
        ],
      },
    });
    const { nodes } = layoutColumns(up, DEFAULT_STYLE, { ...DEFAULT_DECOR, variance: { actual: 0, reference: 1 } });
    expect(label(nodes, 0)!.text).toBe("+10");
    expect(bar(nodes, 0)!.fill).toBe("#0ca30c"); // green (favourable)
  });

  it("flips the sign colouring for cost-like metrics (goodIsUp:false)", () => {
    const { nodes } = layoutColumns(
      cfg({
        kind: "clustered",
        data: {
          categories: ["Q1"],
          series: [
            { name: "AC", values: [90] },
            { name: "PL", values: [80] },
          ],
        },
      }),
      DEFAULT_STYLE,
      { ...DEFAULT_DECOR, variance: { actual: 0, reference: 1, goodIsUp: false } },
    );
    expect(bar(nodes, 0)!.fill).toBe(DEFAULT_STYLE.negative); // +10 is now unfavourable
  });

  it("computes relative Δ% in percent mode", () => {
    const { nodes } = layoutColumns(base, DEFAULT_STYLE, {
      ...DEFAULT_DECOR,
      variance: { actual: 0, reference: 1, mode: "percent" },
    });
    expect(label(nodes, 0)!.text).toBe("-4%"); // (82−85)/85 = −3.5% → −4%
  });

  it("skips a category with a null on either side, and is suppressed when horizontal", () => {
    const withNull = cfg({
      kind: "clustered",
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "AC", values: [82, null] },
          { name: "PL", values: [85, 100] },
        ],
      },
    });
    const { nodes } = layoutColumns(withNull, DEFAULT_STYLE, {
      ...DEFAULT_DECOR,
      variance: { actual: 0, reference: 1 },
    });
    expect(bar(nodes, 1)).toBeUndefined(); // null actual → no bar
    const horiz = layoutColumns({ ...base, horizontal: true }, DEFAULT_STYLE, {
      ...DEFAULT_DECOR,
      variance: { actual: 0, reference: 1 },
    });
    expect(horiz.nodes.find((n) => n.name === "variance-zero")).toBeUndefined();
  });
});

/**
 * Stack membership survives the grid, which can only express it by adjacency.
 */
describe("non-contiguous stack groups", () => {
  const data = {
    categories: ["Q1"],
    series: [
      { name: "Retail", values: [1], stack: 0 },
      { name: "Online", values: [2], stack: 1 },
      { name: "Wholesale", values: [3], stack: 0 },
    ],
  };

  it("round-trips as the same GROUPING, not as three separate columns", () => {
    // `dataToSheet` could only emit a separator when the stack changed from the
    // previous series, and `sheetToData` renumbered by separator count — so
    // [0,1,0] came back [0,1,2]. Two columns became three, and merely loading
    // such a chart into the pane restructured it on the next update.
    const back = sheetToData(dataToSheet(data));
    const groups = new Map<number, string[]>();
    for (const s of back.series) {
      const k = s.stack ?? 0;
      groups.set(k, [...(groups.get(k) ?? []), s.name]);
    }
    expect([...groups.values()].map((g) => g.sort().join("+")).sort()).toEqual(["Online", "Retail+Wholesale"]);
  });

  it("draws the same chart before and after the round trip", () => {
    // The grouping is what the reader sees; the series ORDER inside the config
    // is not, as long as paint order within a group is kept.
    const xs = (d: { categories: string[]; series: { name: string; values: (number | null)[]; stack?: number }[] }) =>
      buildChart({ ...DEFAULT_SIZE, kind: "stacked", data: d } as unknown as ChartConfig)
        .nodes.filter((n) => /^seg-/.test(n.name ?? ""))
        .map((n) => Math.round((n as unknown as { x: number }).x));
    expect(new Set(xs(sheetToData(dataToSheet(data)))).size, "the round trip changed the column count").toBe(
      new Set(xs(data)).size,
    );
  });

  it("leaves a chart that was already contiguous exactly as it was", () => {
    const fine = {
      categories: ["Q1"],
      series: [
        { name: "A", values: [1], stack: 0 },
        { name: "B", values: [2], stack: 0 },
        { name: "C", values: [3], stack: 1 },
      ],
    };
    expect(sheetToData(dataToSheet(fine)).series.map((s) => s.name)).toEqual(["A", "B", "C"]);
  });
});

/**
 * The skill's documented inputs have to work through the engine, not only
 * through the pane's grid.
 */
describe("a Gantt row of ISO dates", () => {
  const plan = (values: (string | number)[][]) =>
    ({
      ...DEFAULT_SIZE,
      kind: "gantt",
      data: {
        categories: ["Design", "Build", "Test"],
        series: [
          { name: "Start", values: values[0] },
          { name: "End", values: values[1] },
        ],
        dates: true,
      },
    }) as unknown as ChartConfig;

  it("draws the same plan whether the dates are ISO strings or epoch days", () => {
    // SKILL.md says "ISO dates supported" and, as a hard Rule, "dates as ISO
    // strings only in Gantt rows" — so this is what an agent writes. Only the
    // datasheet ever parsed them: through the skill, every string fell to null,
    // `layoutGantt` read each task as a section header, and the slide came back
    // with names on grey bands and no bars, no labels, no timeline. Exit 0.
    const iso = buildChart(
      plan([
        ["2026-01-05", "2026-02-02", "2026-03-09"],
        ["2026-01-30", "2026-03-06", "2026-03-27"],
      ]),
    );
    const days = buildChart(
      plan([
        [20458, 20486, 20521],
        [20483, 20518, 20539],
      ]),
    );
    const bars = (s: typeof iso) => s.nodes.filter((n) => n.kind === "rect").length;
    expect(bars(iso), "the ISO plan drew no bars at all").toBeGreaterThan(3);
    expect(bars(iso)).toBe(bars(days));
  });

  it("leaves a non-date row's strings alone", () => {
    // Only rows the Gantt grammar calls dates are parsed — "% complete" and
    // "After" are numbers and must keep falling to null if they arrive as text.
    const cfg = {
      ...DEFAULT_SIZE,
      kind: "gantt",
      data: {
        categories: ["A"],
        series: [
          { name: "Start", values: ["2026-01-05"] },
          { name: "End", values: ["2026-01-30"] },
          { name: "% complete", values: ["2026-01-05"] },
        ],
        dates: true,
      },
    } as unknown as ChartConfig;
    expect(() => buildChart(cfg)).not.toThrow();
  });
});
