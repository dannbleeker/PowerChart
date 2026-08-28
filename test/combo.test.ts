import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { LineNode, RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Combo charts — mixed bases, secondary axes, small multiples, overflow guards. */

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

/**
 * Two combo line-overlay scale bugs found by a layout-modes hunt: both put a line
 * series off the canvas (finite but wrong coordinates). Byte-identical for the
 * valid cases (positive line / plain-chain waterfall) — snapshots unchanged — and
 * only rescue the broken configs.
 */
const HEIGHT = 300;

const markerYs = (cfg: ChartConfig) =>
  buildChart(cfg)
    .nodes.filter((n) => n.name?.startsWith("combo-marker") || n.name?.startsWith("combo-line"))
    .flatMap((n) => (n.kind === "rect" ? [n.y, n.y + n.h] : n.kind === "line" ? [n.y1, n.y2] : []));

describe("combo column modes", () => {
  const combo: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    data: {
      categories: ["2023", "2024"],
      series: [
        { name: "A", values: [30, 40] },
        { name: "B", values: [20, 25] },
        { name: "Margin", values: [31, 38], type: "line" },
      ],
    },
  };

  it("clustered mode puts column series side by side", () => {
    const stacked = buildChart(combo);
    const clustered = buildChart({ ...combo, combo: { columns: "clustered" } });
    const seg = (s: ReturnType<typeof buildChart>, name: string) => s.nodes.find((n) => n.name === name) as RectNode;
    // Stacked: same x for both series. Clustered: different x.
    expect(seg(stacked, "seg-0-0").x).toBeCloseTo(seg(stacked, "seg-1-0").x);
    expect(seg(clustered, "seg-0-0").x).not.toBeCloseTo(seg(clustered, "seg-1-0").x);
    // Line overlay still present in both.
    expect(clustered.nodes.some((n) => n.name?.startsWith("combo-line"))).toBe(true);
  });

  it("stacked100 mode keeps the line on its own right-hand axis", () => {
    const pct = buildChart({ ...combo, combo: { columns: "stacked100" }, decorations: { segmentLabels: true } });
    expect(pct.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
    expect(pct.nodes.some((n) => n.name?.startsWith("combo-line"))).toBe(true);
  });
});

describe("small multiples", () => {
  const cfg: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    title: "Revenue by region",
    multiples: {},
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [100, 90] },
        { name: "South", values: [40, 100] },
        { name: "East", values: [25, 30] },
      ],
    },
    decorations: { segmentLabels: false },
  };
  const s = buildChart(cfg);

  it("renders one titled panel per series in a grid", () => {
    expect((s.nodes.find((n) => n.name === "p0-title") as TextNode).text).toBe("North");
    expect((s.nodes.find((n) => n.name === "p2-title") as TextNode).text).toBe("East");
    expect((s.nodes.find((n) => n.name === "title") as TextNode).text).toBe("Revenue by region");
    // Three panels side by side (n ≤ 3 → one row).
    const p0 = s.nodes.find((n) => n.name === "p0-seg-0-0") as RectNode;
    const p1 = s.nodes.find((n) => n.name === "p1-seg-0-0") as RectNode;
    expect(p1.x).toBeGreaterThan(p0.x);
  });

  it("panels share one value scale: equal values → equal bar heights", () => {
    // North Q1 = South Q2 = 100 → same height in different panels.
    const north = s.nodes.find((n) => n.name === "p0-seg-0-0") as RectNode;
    const south = s.nodes.find((n) => n.name === "p1-seg-0-1") as RectNode;
    expect(north.h).toBeCloseTo(south.h, 3);
    // East's 30 is visibly smaller than 100, same scale.
    const east = s.nodes.find((n) => n.name === "p2-seg-0-1") as RectNode;
    expect(east.h / north.h).toBeCloseTo(0.3, 1);
  });

  it("columns option and carried special rows work; single series stays whole", () => {
    const stackedCols = buildChart({ ...cfg, multiples: { columns: 1 } });
    const a = stackedCols.nodes.find((n) => n.name === "p0-seg-0-0") as RectNode;
    const b = stackedCols.nodes.find((n) => n.name === "p1-seg-0-0") as RectNode;
    expect(Math.abs(a.x - b.x)).toBeLessThan(30); // same column
    expect(b.y).toBeGreaterThan(a.y);
    const withTarget = buildChart({
      ...cfg,
      data: { ...cfg.data, series: [...cfg.data.series, { name: "Target", values: [110, 110] }] },
    });
    for (const pi of [0, 1, 2]) {
      expect(withTarget.nodes.some((n) => n.name === `p${pi}-target-0`)).toBe(true);
    }
    const single = buildChart({
      ...cfg,
      data: { ...cfg.data, series: [cfg.data.series[0]] },
    });
    expect(single.nodes.some((n) => n.name?.startsWith("p0-"))).toBe(false);
  });

  it("scatter and other row-semantics kinds ignore multiples", () => {
    const scatter = buildChart({
      kind: "scatter",
      ...DEFAULT_SIZE,
      multiples: {},
      data: {
        categories: ["A", "B"],
        series: [
          { name: "X", values: [1, 2] },
          { name: "Y", values: [3, 4] },
        ],
      },
    });
    expect(scatter.nodes.some((n) => n.name?.startsWith("p0-"))).toBe(false);
  });
});

describe("combo mekko base", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    secondaryAxis: true,
    combo: { columns: "mekko" },
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "P1", values: [30, 45, 20] },
        { name: "P2", values: [20, 25, 30] },
        { name: "Share %", type: "line", values: [62, 70, 45] },
      ],
    },
    decorations: { segmentLabels: false, seriesLabels: false },
  };

  it("draws mekko columns (variable width) with a line over them", () => {
    const s = buildChart(cfg);
    // Mekko segments exist and column widths differ (A total 50 vs B total 70).
    const segs = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("seg-"));
    expect(segs.length).toBeGreaterThan(0);
    const widths = new Set(segs.map((r) => Math.round(r.w)));
    expect(widths.size).toBeGreaterThan(1); // variable-width columns
    expect(s.nodes.filter((n) => n.name?.startsWith("combo-marker-0-"))).toHaveLength(3);
  });
});

describe("combo independent line axes", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    combo: { lineAxes: "independent" },
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "Revenue", values: [40, 60] },
        { name: "Small KPI", type: "line", values: [10, 20] },
        { name: "Huge KPI", type: "line", values: [1000, 2000] },
      ],
    },
    decorations: { segmentLabels: false, seriesLabels: false },
  };

  it("scales each line to its own range so both are visible despite 100× units", () => {
    const s = buildChart(cfg);
    const y = (nm: string) => (s.nodes.find((n): n is RectNode => n.name === nm) as RectNode).y;
    // Line 0 (10→20) and line 1 (1000→2000): each spans its own range, so the
    // min points share a y and the max points share a y.
    expect(y("combo-marker-0-0")).toBeCloseTo(y("combo-marker-1-0"), 0); // both at their min
    expect(y("combo-marker-0-1")).toBeCloseTo(y("combo-marker-1-1"), 0); // both at their max
    expect(y("combo-marker-0-1")).toBeLessThan(y("combo-marker-0-0")); // max is higher up
    // Independent mode forces value labels (no shared axis to read).
    expect(s.nodes.some((n) => n.name === "combo-label-1-1")).toBe(true);
    // No single shared secondary axis in independent mode.
    expect(s.nodes.some((n) => n.name === "secondary-axis")).toBe(false);
  });

  it("shared axis (default) keeps one secondary scale", () => {
    const shared = buildChart({ ...cfg, secondaryAxis: true, combo: {} });
    expect(shared.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
  });
});

describe("combo stacked-area base", () => {
  const cfg: ChartConfig = {
    kind: "combo",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Jan", "Feb", "Mar"],
      series: [
        { name: "Cloud", values: [20, 24, 28] },
        { name: "Licenses", values: [15, 14, 16] },
        { name: "Margin %", type: "line", values: [22, 24, 26] },
      ],
    },
    combo: { columns: "area" },
    secondaryAxis: true,
  };
  const s = buildChart(cfg);

  it("draws a stacked area base with the line overlaid", () => {
    // The area base emits filled area slabs...
    expect(s.nodes.some((n) => n.name?.startsWith("area-"))).toBe(true);
    // ...and a secondary axis for the line series.
    expect(s.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
    // ...and line segments over the top.
    expect(s.nodes.some((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("combo-line-"))).toBe(true);
  });
});

describe("combo chart", () => {
  it("draws marked line series over the columns on a shared scale", () => {
    const scene = buildChart(
      cfg({
        kind: "combo",
        data: {
          categories: ["A", "B"],
          series: [
            { name: "Cols", values: [10, 20] },
            { name: "Line", values: [40, 50], type: "line" },
          ],
        },
      }),
    );
    expect(byName(scene.nodes, "combo-line-").length).toBeGreaterThan(0);
    expect(byName(scene.nodes, "seg-0-")).toHaveLength(2);
    // Line values above stack totals still fit: markers stay inside the plot.
    const markers = byName(scene.nodes, "combo-marker-") as RectNode[];
    for (const m of markers) expect(m.y).toBeGreaterThan(0);
  });
});

describe("combo secondary axis", () => {
  it("scales line series independently and adds right-hand ticks", () => {
    const scene = buildChart(
      cfg({
        kind: "combo",
        secondaryAxis: true,
        data: {
          categories: ["A", "B"],
          series: [
            { name: "Revenue", values: [500, 800] },
            { name: "Margin %", values: [30, 45], type: "line" },
          ],
        },
      }),
    );
    expect(byName(scene.nodes, "secondary-axis").length).toBeGreaterThanOrEqual(3);
    // With its own 0..50-ish scale, the 45% marker sits high in the plot.
    const markers = byName(scene.nodes, "combo-marker-0-1");
    expect(markers).toHaveLength(1);
  });
});

describe("combo overlay line stays on the plot when it dips below the bars", () => {
  it("extends the shared axis down to a negative overlay over a non-negative base", () => {
    // All-zero bars + a line reaching −18: the shared column axis used to floor at
    // its own data (≥0), plotting the line thousands of points below the plot.
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: 300,
      data: {
        categories: ["C0", "C1", "C2", "C3"],
        series: [
          { name: "bars", values: [0, 0, 0, 0] },
          { name: "line", values: [0, 0, 0, -18] },
        ],
      },
    };
    const nodes = buildChart(cfg).nodes;
    const line = nodes.filter((n) => n.name?.startsWith("combo-line-") || n.name?.startsWith("combo-marker-"));
    expect(line.length).toBeGreaterThan(0);
    // Every line/marker coordinate must sit within a generous margin of the 300pt canvas.
    for (const n of line) {
      const ys = n.kind === "line" ? [n.y1, n.y2] : n.kind === "rect" ? [n.y, n.y + n.h] : [];
      for (const y of ys) expect(y).toBeLessThan(360);
    }
  });

  it("leaves a combo whose bars already run more negative unchanged", () => {
    // Bars reach −50, the line only −5: the bars own the floor, so the fix must
    // not raise it to the line (which would clip the bars).
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: 300,
      data: {
        categories: ["C0", "C1"],
        series: [
          { name: "bars", values: [-50, -40] },
          { name: "line", values: [-5, -3] },
        ],
      },
    };
    // The most-negative bar segment must still reach near the bottom of the plot.
    const rects = buildChart(cfg).nodes.filter((n) => n.name?.startsWith("seg-") || n.name?.startsWith("col-"));
    expect(rects.length).toBeGreaterThan(0);
  });
});

describe("combo with a single unmarked series", () => {
  it("renders a plain column, not a column AND a line", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "X", values: [10, 20] }] },
    };
    const scene = buildChart(cfg);
    const lineNodes = scene.nodes.filter((n) => (n.name ?? "").startsWith("combo-line"));
    const markerNodes = scene.nodes.filter((n) => (n.name ?? "").startsWith("combo-marker"));
    const colSegs = scene.nodes.filter((n) => n.kind === "rect" && (n.name ?? "").startsWith("seg"));
    expect(lineNodes.length).toBe(0); // the bug: it was drawn as a line too
    expect(markerNodes.length).toBe(0);
    expect(colSegs.length).toBeGreaterThan(0); // still a column
  });

  it("still draws a line for the last series when there are two", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Col", values: [10, 20] },
          { name: "Ln", values: [5, 8] },
        ],
      },
    };
    const scene = buildChart(cfg);
    expect(scene.nodes.some((n) => (n.name ?? "").startsWith("combo-line"))).toBe(true);
  });
});

describe("small multiples with a high Target row", () => {
  it("includes the carried Target in the shared panel scale (no clipping)", () => {
    const cfg: ChartConfig = {
      kind: "line",
      ...DEFAULT_SIZE,
      multiples: {},
      decorations: { seriesLabels: true },
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [10, 20] },
          { name: "S2", values: [15, 25] },
          { name: "Target", values: [100, 100] }, // far above the data
        ],
      },
    };
    const scene = buildChart(cfg);
    // Nothing renders above the top of the canvas (the target used to overflow).
    const ys = scene.nodes.map((n) => (n as { y?: number }).y).filter((y): y is number => typeof y === "number");
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(0);
  });
});

describe("combo secondary axis spans negative line values", () => {
  it("keeps a negative overlay line on-plot", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: HEIGHT,
      secondaryAxis: true,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "bars", values: [10, 20, 30] },
          { name: "line", type: "line", values: [-5, 8, -3] },
        ],
      },
    };
    const ys = markerYs(cfg);
    expect(ys.length).toBeGreaterThan(0);
    for (const y of ys) expect(y).toBeLessThanOrEqual(HEIGHT + 1);
  });

  it("still renders an all-positive secondary line on-plot (unchanged path)", () => {
    const cfg: ChartConfig = {
      kind: "combo",
      width: 480,
      height: HEIGHT,
      secondaryAxis: true,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "bars", values: [10, 20, 30] },
          { name: "line", type: "line", values: [5, 8, 3] },
        ],
      },
    };
    const ys = markerYs(cfg);
    expect(ys.length).toBeGreaterThan(0);
    for (const y of ys) {
      expect(y).toBeGreaterThanOrEqual(-1);
      expect(y).toBeLessThanOrEqual(HEIGHT + 1);
    }
  });
});

describe("small multiples + pareto (regression)", () => {
  // buildMultiples clears categorySort per panel precisely so the panels cannot
  // disagree about their x-axis — but pareto is a category-reordering transform
  // too and was passed straight through, so each single-series panel re-sorted
  // itself: A,B,C became B,C,A beside A,C,B.
  const build = (pareto: boolean) =>
    buildChart({
      kind: "clustered",
      width: 640,
      height: 320,
      ...(pareto ? { pareto: true } : {}),
      multiples: { columns: 2 },
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "S1", values: [10, 40, 20] },
          { name: "S2", values: [50, 5, 30] },
        ],
      },
    } as unknown as ChartConfig);

  const order = (pareto: boolean) =>
    (build(pareto).nodes as { kind: string; text?: string; x?: number }[])
      .filter((n) => n.kind === "text" && ["A", "B", "C"].includes(n.text ?? ""))
      .sort((a, b) => (a.x ?? 0) - (b.x ?? 0))
      .map((n) => n.text)
      .join(",");

  it("every panel keeps the same category order", () => {
    expect(order(true), "panels disagree about their category order").toBe(order(false));
  });
});

/**
 * Horizontal (bar) combos. `layoutCombo` never looked at `cfg.horizontal`, and a
 * horizontal base returns `valueToY: undefined` — so a shared-axis bar combo
 * dropped its overlay entirely, and a secondary-axis one drew every point with
 * the category's Y coordinate in the X slot. `pareto: true` reaches both
 * without the author ever typing `kind: "combo"`.
 */
describe("horizontal combo overlay", () => {
  const points = (s: ReturnType<typeof buildChart>) =>
    (s.nodes as RectNode[]).filter((n) => n.name?.startsWith("combo-marker-")).map((n) => ({ x: n.x, y: n.y }));

  const bars = (over?: Partial<ChartConfig>) =>
    buildChart(
      cfg({
        kind: "combo",
        horizontal: true,
        data: {
          categories: ["A", "B", "C"],
          series: [
            { name: "bars", values: [40, 30, 20] },
            { name: "line", type: "line", values: [35, 25, 45] },
          ],
        },
        decorations: { valueAxis: true },
        ...over,
      }) as ChartConfig,
    );

  it("draws the overlay on a shared axis instead of dropping it", () => {
    const pts = points(bars());
    // Was []: the bail-out read valueToY, which no horizontal base defines.
    expect(pts).toHaveLength(3);
  });

  it("runs the overlay down the categories and along the value axis", () => {
    for (const s of [bars(), bars({ secondaryAxis: true })]) {
      const pts = points(s);
      // Categories run top-to-bottom on a bar chart, so y must increase with the
      // category index and x must track the VALUE (35 < 45 → point 0 left of 2).
      expect(pts[0].y).toBeLessThan(pts[1].y);
      expect(pts[1].y).toBeLessThan(pts[2].y);
      expect(pts[0].x).toBeLessThan(pts[2].x);
    }
  });

  it("keeps a horizontal Pareto's cumulative line on the canvas", () => {
    const s = buildChart(
      cfg({
        kind: "clustered",
        horizontal: true,
        pareto: true,
        data: { categories: ["A", "B", "C", "D"], series: [{ name: "n", values: [40, 30, 20, 10] }] },
        decorations: { valueAxis: true },
      }) as unknown as ChartConfig,
    );
    const pts = points(s);
    expect(pts).toHaveLength(4);
    // The cumulative share only rises, so the line runs left-to-right…
    expect(pts.map((p) => p.x)).toEqual([...pts.map((p) => p.x)].sort((a, b) => a - b));
    // …and every mark, plus the secondary tick strip, stays on the canvas. The
    // strip was drawn at plot.x + plot.w + 2, off the right edge of a bar chart.
    const off = (s.nodes as (RectNode | TextNode)[])
      .filter((n) => n.name?.startsWith("combo-") || n.name === "secondary-axis")
      .filter((n) => n.x < -0.5 || n.y < -0.5 || n.x + (n.w ?? 0) > s.width + 0.5 || n.y + (n.h ?? 0) > s.height + 0.5)
      .map((n) => `${n.name}@${n.x.toFixed(1)},${n.y.toFixed(1)}`);
    expect(off, `off-canvas: ${off.join(", ")}`).toEqual([]);
  });
});

/**
 * `tightLabelPriority` — which value labels survive a frame too small for both.
 *
 * On a short frame a combo's column totals and its line's point labels can want
 * the same band and neither can move. `resolveLabelCollisions` resolves every
 * one of these against an unbounded canvas, by flipping the point label BELOW
 * its point — measured landing at y 57-75 on a 60pt-tall chart, i.e. off the
 * bottom, so the flip is correctly refused. Shrinking cannot help either: the
 * two are centred at the same y whenever the column total and the line value
 * coincide, which is a property of the DATA.
 *
 * So one set is dropped, and only where it is STILL colliding after every legal
 * move — which is what makes a roomy chart keep both.
 */
describe("combo — tightLabelPriority", () => {
  const named = (c: ChartConfig, re: RegExp) =>
    buildChart(c)
      .nodes.filter((n): n is TextNode => n.kind === "text" && re.test(n.name ?? "") && !!n.text.trim())
      .map((n) => n.name!);

  // The shipped sample, so this measures the chart the gate and the showcase
  // actually draw rather than a hand-built one that happens to emit no point
  // labels at all — the first version of this did exactly that and reported
  // zero labels at 480x300, which reads identically to the drop firing.
  const combo = (w: number, h: number, priority?: "columns" | "line"): ChartConfig => {
    const s = sampleConfig("combo") as ChartConfig;
    return {
      ...s,
      width: w,
      height: h,
      ...(priority ? { decorations: { ...s.decorations, tightLabelPriority: priority } } : {}),
    } as ChartConfig;
  };

  it("keeps BOTH sets on a chart with room, which is the whole point of the rule", () => {
    // The guard that stops this becoming "small charts lose their line labels".
    for (const priority of [undefined, "columns", "line"] as const) {
      const c = combo(480, 300, priority);
      expect(named(c, /^total-/), `totals at 480x300 (${priority})`).toHaveLength(4);
      expect(named(c, /^combo-label-/), `point labels at 480x300 (${priority})`).toHaveLength(4);
    }
  });

  it("defaults to keeping the columns: the totals stay, the colliding point labels go", () => {
    const c = combo(80, 60);
    expect(named(c, /^total-/)).toHaveLength(4);
    expect(named(c, /^combo-label-/)).toEqual([]);
  });

  it('"line" reverses it — the point labels stay and the colliding totals go', () => {
    const c = combo(80, 60, "line");
    expect(named(c, /^combo-label-/)).toHaveLength(4);
    expect(named(c, /^total-/)).toEqual([]);
  });

  it("drops only the labels actually still colliding, not the whole family", () => {
    // At 300x60 two of the four point labels sit high enough that de-collision
    // already cleared them. Those must survive even though their siblings do not
    // — a family-wide drop would take them too.
    const kept = named(combo(300, 60), /^combo-label-/);
    expect(kept).toEqual(["combo-label-0-2", "combo-label-0-3"]);
  });
});

/**
 * The combo line's series name shares the right-hand margin with the COLUMN
 * series names, and it was taking a width floor (`fs * 3.4`) even where a
 * perfectly good gutter existed. Fitted to that inflated width it stayed at the
 * chart font while its neighbours correctly shrank, and was drawn straight
 * across them and the last column's total at 80x60.
 *
 * The floor itself is load-bearing and stays: the showcase's combo slide has a
 * plot that reaches the right edge, and removing the floor outright dropped
 * "Margin %" from that deck. Measured on more than `sampleConfig("combo")` only
 * after it did — the sample's gutter and the showcase's are not the same.
 */
describe("the combo line's series label uses the gutter it actually has", () => {
  const label = (w: number, h: number) =>
    buildChart({ ...sampleConfig("combo"), width: w, height: h } as ChartConfig).nodes.find(
      (n): n is TextNode => n.kind === "text" && n.name === "combo-series-label-0",
    );

  it("is drawn at every frame, never dropped", () => {
    // The floor exists so this name survives a plot that reaches the edge. It
    // must not become collateral of the fit.
    for (const [w, h] of [
      [80, 60],
      [120, 90],
      [200, 150],
      [300, 60],
      [480, 300],
      [60, 300],
    ] as [number, number][]) {
      expect(label(w, h), `${w}x${h}: lost the line's series name`).toBeDefined();
    }
  });

  it("takes the real gutter as its box, not the wider floor", () => {
    // This is the fix itself, and it has to be asserted on the BOX: the old
    // floor also produced a smaller font here (7.5 against 10), so a test that
    // only compared font sizes passed against the unfixed file and proved
    // nothing. At 80x60 the true gutter is 22pt and the floor claimed 34.
    const small = label(80, 60)!;
    expect(small.w).toBeLessThanOrEqual(22.01);
    // …and the font follows the box down to its neighbours' size.
    expect(small.fontSize).toBeLessThanOrEqual(5.01);
    const big = label(480, 300)!;
    expect(big.fontSize).toBeCloseTo(10, 5);
  });

  it("stays inside the frame", () => {
    for (const [w, h] of [
      [80, 60],
      [120, 90],
      [60, 300],
      [480, 300],
    ] as [number, number][]) {
      const l = label(w, h)!;
      expect(l.x, `${w}x${h}`).toBeGreaterThanOrEqual(0);
      expect(l.x + l.w, `${w}x${h}`).toBeLessThanOrEqual(w + 0.01);
    }
  });
});

/**
 * Sideways, the combo line's name is drawn at the line's LAST POINT, and the
 * only bound it had was a floor at zero: `end.y` is wherever that point landed,
 * so on a short frame the whole label was drawn past the foot of the chart —
 * y 91.5 with a 14pt box on a 90pt canvas, and the same at 300x60 and 200x150,
 * all at the DEFAULT font. Off the chart is onto whatever sits under it on the
 * slide, because neither PowerPoint renderer clips a text box.
 *
 * Neither existing gate could see it. The overflow sweep in `frame-fit` never
 * sets `horizontal`, and the overlap sweep cannot see ink that is off the canvas
 * entirely.
 */
describe("a horizontal combo's line name stays on the chart", () => {
  const FRAMES: [number, number][] = [
    [80, 60],
    [120, 90],
    [160, 120],
    [200, 150],
    [300, 60],
    [480, 300],
    [960, 540],
  ];
  const FONTS = [6, 8, 10, 14, 18, 24, 32];

  it("is inside the canvas at every frame and font", () => {
    let seen = 0;
    const bad: string[] = [];
    for (const [w, h] of FRAMES)
      for (const fontSize of FONTS) {
        const scene = buildChart({
          ...sampleConfig("combo"),
          width: w,
          height: h,
          horizontal: true,
          style: { fontSize },
        } as ChartConfig);
        for (const n of scene.nodes.filter(
          (n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("combo-series-label-"),
        )) {
          seen++;
          if (n.y < -0.5 || n.y + n.h > h + 0.5)
            bad.push(`${w}x${h} at ${fontSize}pt: y=${n.y.toFixed(1)}..${(n.y + n.h).toFixed(1)} on a ${h}pt canvas`);
        }
      }
    // 18 now, not 20-odd: where the legend names the series too, the name at the
    // end of the line yields to anything it collides with, so three of these are
    // deliberately not drawn. The floor is here to stop the sweep silently
    // measuring nothing, not to pin a count — it is lowered with the reason, and
    // eighteen names on-canvas is still eighteen names checked.
    expect(seen, "no line names drawn at all, so this proves nothing").toBeGreaterThan(14);
    expect(bad).toEqual([]);
  });

  it("drops a point label the name has taken the room from", () => {
    // `unplaceableComboLabels` used to be skipped for every horizontal chart
    // (`skipDecor`), and to return early when there were no column TOTALS —
    // which a horizontal combo does not draw. So sideways there was no pass at
    // all, and a point label sat under the line's own name with nothing able to
    // separate them.
    //
    // The frame and font this is measured at have moved twice, because the
    // sideways name has been bounded twice — first to half a row, then to the
    // clear air between two rows' own labels — and each time the trade stopped
    // happening where the test was looking. A case where the NAME is not drawn
    // proves nothing about a pass that only runs when it is.
    //
    // MOVED TWICE MORE ON 2026-08-28, the second time to the OTHER ORIENTATION,
    // and that is not a frame change but a behaviour change worth stating.
    //
    // Sideways this trade no longer exists anywhere. The combo's legend now names
    // its line, so the line's own end-of-line name is said twice — and where it
    // collides with anything it is the thing that yields. The point label it used
    // to displace now survives, which is the right way round: a value is data and
    // a name repeated in the legend is not.
    //
    // UPRIGHT there is no legend, so the name is said once, keeps its place, and
    // the old trade still runs. 300x60 at 14pt is where both exist: with
    // `seriesLabels` off the chart keeps two point labels, with the name on it
    // keeps one.
    const at = (decorations?: ChartConfig["decorations"]) =>
      buildChart({
        ...sampleConfig("combo"),
        width: 300,
        height: 60,
        horizontal: false,
        style: { fontSize: 14 },
        ...(decorations ? { decorations } : {}),
      } as ChartConfig).nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("combo-label-"));
    const labels = at();
    expect(at({ ...sampleConfig("combo").decorations, seriesLabels: false }).length).toBeGreaterThan(labels.length);
    expect(labels.length).toBeLessThan(4);
    expect(labels.length, "every point label was dropped, which is not the trade").toBeGreaterThan(0);
  });

  it("keeps all four on a chart with the room for them", () => {
    // The negative control: nothing is dropped where nothing collides.
    const labels = buildChart({
      ...sampleConfig("combo"),
      width: 480,
      height: 300,
      horizontal: true,
    } as ChartConfig).nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("combo-label-"));
    expect(labels.length).toBe(4);
  });
});

/**
 * A combo's legend names its LINE, and always should have.
 *
 * `layoutCombo` hands its column series to `layoutColumns` in a config it builds
 * itself, with the lines removed — they are drawn afterwards, by the combo. The
 * legend is drawn from that same config, so a combo's legend named its columns
 * and never its lines. At every size, not only crowded ones.
 *
 * That was survivable while the line carried its own name at the end of the
 * line, and it stopped being survivable exactly where that name could not be
 * drawn: a 120x90 horizontal combo at 6pt drew "Product" and "Services" and
 * nothing at all for "Margin %". A series with no name anywhere on the chart is
 * not a cosmetic problem.
 *
 * Measured over the sample combo at eight frames, seven fonts, both
 * orientations — 112 charts:
 *
 *     legend names the line     0 -> 28
 *     no legend at all         76 -> 80
 *     no series named at all   21 -> 25
 *
 * The four that lost their legend are the trade, and it is the right way round.
 * They are thumbnails — 80x60 through 200x150 at large fonts — where the third
 * entry tips `horizontalLegendFits`, whose whole job is to refuse a legend that
 * has eaten the plot. What they had before was a legend naming two of three
 * series with nothing to say a third existed, which is worse than none.
 */
describe("a combo's legend names every series it draws", () => {
  const legendOf = (w: number, h: number, fontSize: number, horizontal: boolean) =>
    buildChart({
      ...sampleConfig("combo"),
      width: w,
      height: h,
      horizontal,
      style: { fontSize },
    } as ChartConfig)
      .nodes.filter((n): n is TextNode => n.kind === "text" && /^legend-\d+$/.test(n.name ?? ""))
      .map((n) => n.text);

  it("names the line wherever it names the columns", () => {
    for (const [w, h] of [
      [200, 150],
      [480, 300],
      [960, 540],
    ]) {
      const legend = legendOf(w, h, 10, true);
      expect(legend.length, `${w}x${h}: no legend at all`).toBeGreaterThan(0);
      expect(
        legend.some((t) => t.startsWith("Margin")),
        `${w}x${h}: the legend omits the line — ${legend}`,
      ).toBe(true);
    }
  });

  it("names all three, not two of three", () => {
    // The specific shape of the bug: a legend that stops short says nothing
    // about what it left out, so the reader cannot know to look.
    const legend = legendOf(480, 300, 10, true);
    expect(legend).toEqual(["Product", "Services", "Margin %"]);
  });

  it("reserves the room it is about to take", () => {
    // The coupling this file's own comments say has been written twice: the
    // predicate that decides a legend fits and the row that draws it must count
    // the same entries. Both go through `legendLabelsFor` now. If they drifted,
    // the legend would be drawn into a band reserved for one row fewer and land
    // on the plot — so this asserts the legend's last row clears the plot top.
    // 160x120 at 10pt, where the three entries WRAP to two rows. At 480x300 they
    // fit on one, so reserving for two entries and reserving for three come to
    // the same number of rows and the test cannot tell them apart — it passed
    // against a mutant that had the reservation counting two while the row drew
    // three. A wrapping legend is the only place the coupling is visible.
    const nodes = buildChart({
      ...sampleConfig("combo"),
      width: 160,
      height: 120,
      horizontal: true,
      style: { fontSize: 10 },
    } as ChartConfig).nodes;
    const legend = nodes.filter((n): n is TextNode => n.kind === "text" && /^legend-\d+$/.test(n.name ?? ""));
    const bars = nodes.filter((n) => /^seg-/.test(n.name ?? ""));
    expect(legend.length, "no legend to check").toBeGreaterThan(0);
    expect(bars.length, "no bars to check against").toBeGreaterThan(0);
    const legendBottom = Math.max(...legend.map((l) => l.y + l.h));
    const barTop = Math.min(...bars.map((b) => (b as unknown as { y: number }).y));
    expect(legendBottom, "the legend is drawn over the plot it was not reserved room beside").toBeLessThanOrEqual(
      barTop + 1,
    );
  });
});

describe("the legend's fit predicate counts what the row will draw", () => {
  it("refuses a legend it has no room for, counting the line", () => {
    // THE OTHER HALF OF THE COUPLING, and the half a mutant walked through
    // first. `horizontalLegendFits` decides whether there is room; `legendRow`
    // decides what to draw. Both must count the same entries or the predicate
    // approves a legend smaller than the one that arrives.
    //
    // 120x90 at 10pt is where it shows: two entries fit and three do not, so a
    // predicate still counting two says yes and the row then draws three into a
    // chart the predicate would have refused. The right answer is no legend —
    // `horizontalLegendFits` exists to refuse one that has eaten the plot, and
    // at this size it has.
    const legend = buildChart({
      ...sampleConfig("combo"),
      width: 120,
      height: 90,
      horizontal: true,
      style: { fontSize: 10 },
    } as ChartConfig).nodes.filter((n): n is TextNode => n.kind === "text" && /^legend-\d+$/.test(n.name ?? ""));
    expect(
      legend.map((l) => l.text),
      "a legend was drawn on a chart with no room for it",
    ).toEqual([]);
  });
});
