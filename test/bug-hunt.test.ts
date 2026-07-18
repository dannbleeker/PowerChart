import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE, valueExtent } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";
import { formatNumber, parseDateToken } from "../src/core/format";
import { textWidth } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Regression tests for defects found during a codebase bug-hunt pass. */

describe("full-circle pie / doughnut (single slice)", () => {
  it("renders a visible circle path for a 360° pie", () => {
    const cfg: ChartConfig = {
      kind: "pie",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [100] }] },
    };
    const scene = buildChart(cfg);
    const wedge = scene.nodes.find((n) => n.kind === "wedge");
    expect(wedge).toBeTruthy();
    // The lone slice spans the whole circle.
    expect((wedge as any).endAngle - (wedge as any).startAngle).toBeCloseTo(360);
    const svg = sceneToSvg(scene);
    // Must emit a real path with area — the old code drew a degenerate arc
    // between coincident points, producing nothing.
    expect(svg).toMatch(/fill-rule="evenodd"/);
    expect(svg).toMatch(/<path[^>]*A /);
  });

  it("renders a full disc plus an overpainted hole for a 360° doughnut", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [100] }] },
    };
    const scene = buildChart(cfg);
    // The lone slice is a full-circle wedge (the hole is faked with an
    // overpainted background ellipse for Office.js compatibility).
    const wedge = scene.nodes.find((n) => n.kind === "wedge") as any;
    expect(wedge.endAngle - wedge.startAngle).toBeCloseTo(360);
    expect(scene.nodes.some((n) => n.kind === "ellipse" && (n as any).name === "hole")).toBe(true);
    const svg = sceneToSvg(scene);
    expect(svg).toMatch(/fill-rule="evenodd"/);
    // A real disc, not a degenerate zero-area arc.
    const path = svg.match(/<path d="([^"]*)"[^>]*fill-rule="evenodd"/)!;
    expect(path[1]).not.toMatch(/NaN/);
    expect((path[1].match(/A /g) || []).length).toBeGreaterThanOrEqual(2);
  });
});

describe("formatNumber negative-zero", () => {
  it("normalises a rounded -0 to 0", () => {
    expect(formatNumber(-0.4, { decimals: 0 })).toBe("0");
    expect(formatNumber(-0.001, { decimals: 0 })).toBe("0");
  });
  it("keeps the sign on genuine negatives", () => {
    expect(formatNumber(-5.2)).toBe("-5.2");
    expect(formatNumber(-0.4, { decimals: 0, forceSign: true })).toBe("0");
  });
});

describe("parseDateToken numeric ranges", () => {
  it("rejects hyphenated numeric ranges as category labels", () => {
    expect(parseDateToken("3-5")).toBeNull();
    expect(parseDateToken("10-20")).toBeNull();
    expect(parseDateToken("18–24")).toBeNull(); // en dash
  });
  it("still parses real dates", () => {
    expect(parseDateToken("2026-01-15")).toBeTypeOf("number");
    expect(parseDateToken("2026-01")).toBeTypeOf("number");
  });
});

describe("categorySort excludes carried rows", () => {
  it("ranks by real stack totals, ignoring Error/Target rows", () => {
    // Category B has the larger real total (30 vs 20) but a huge Target on A.
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      categorySort: "descending",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Value", values: [20, 30] },
          { name: "Target", values: [999, 1] },
        ],
      },
    };
    const scene = buildChart(cfg);
    // Descending by real total → B (30) before A (20). Find the category axis
    // labels in order.
    const labels = scene.nodes
      .filter((n) => n.kind === "text" && (n as any).name?.startsWith("category-"))
      .map((n) => (n as any).text);
    expect(labels[0]).toBe("B");
    expect(labels[1]).toBe("A");
  });
});

describe("degenerate inputs do not throw", () => {
  it("boxplot with an empty series", () => {
    const cfg: ChartConfig = {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [] }] },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
  it("area chart with a single empty category set", () => {
    const cfg: ChartConfig = {
      kind: "area",
      ...DEFAULT_SIZE,
      decorations: { seriesLabels: true },
      data: { categories: [], series: [{ name: "S", values: [] }] },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
});

/** Helper: the vertical span of every rect in a scene. */
const rectSpan = (scene: { nodes: any[] }) => {
  const ys = scene.nodes.flatMap((n) => (n.kind === "rect" ? [n.y, n.y + n.h] : []));
  return { top: Math.min(...ys), bottom: Math.max(...ys) };
};

describe("value extents and auto-scales", () => {
  it("area extent honours the negative stack, like stacked (was floored at 0)", () => {
    const data = {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "A", values: [10, -40, 20] },
        { name: "B", values: [5, -30, 10] },
      ],
    };
    expect(valueExtent({ kind: "area", ...DEFAULT_SIZE, data } as ChartConfig)).toEqual({ min: -70, max: 30 });
    // identical data must give the same extent as the stacked sibling
    expect(valueExtent({ kind: "area", ...DEFAULT_SIZE, data } as ChartConfig)).toEqual(
      valueExtent({ kind: "stacked", ...DEFAULT_SIZE, data } as ChartConfig),
    );
  });

  it("negative small-multiples areas stay inside the scene", () => {
    const scene = buildChart({
      kind: "area",
      width: 480,
      height: 300,
      multiples: {},
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "A", values: [10, -40, 20] },
          { name: "B", values: [5, -30, 10] },
        ],
      },
    } as ChartConfig);
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1);
    expect(bottom).toBeLessThanOrEqual(301); // was ~775 — rendered far below the canvas
  });

  it("a target below the data range widens the scale down, not just up", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Actual", values: [10, 20, 30] },
          { name: "Target", values: [-50, -50, -50] },
        ],
      },
    } as ChartConfig);
    const ticks = scene.nodes.filter((n: any) => n.kind === "line" && n.name?.startsWith("target-"));
    expect(ticks.length).toBe(3);
    for (const t of ticks as any[]) {
      expect(t.y1).toBeGreaterThanOrEqual(0);
      expect(t.y1).toBeLessThanOrEqual(300); // target tick used to land outside the plot
    }
  });

  it("a multi-series waterfall combo keeps its columns on canvas", () => {
    const scene = buildChart({
      kind: "combo",
      ...DEFAULT_SIZE,
      combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "W1", values: [10, 10, 10] },
          { name: "W2", values: [10, 10, 10] },
          { name: "L", type: "line", values: [40, 40, 40] },
        ],
      },
    } as ChartConfig);
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1); // was -129 — stacked peak understated
    expect(bottom).toBeLessThanOrEqual(scene.height + 1);
  });

  it("small-multiples panels share one category order under categorySort", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 480,
      height: 300,
      multiples: { columns: 2 },
      categorySort: "descending",
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "S1", values: [1, 9, 5] },
          { name: "S2", values: [9, 1, 5] },
        ],
      },
      decorations: { categoryAxis: true },
    } as ChartConfig);
    const axisOf = (p: string) =>
      scene.nodes
        .filter((n: any) => n.kind === "text" && n.name?.startsWith(`${p}-category-`))
        .sort((a: any, b: any) => a.x - b.x)
        .map((n: any) => n.text);
    const p0 = axisOf("p0"),
      p1 = axisOf("p1");
    expect(p0.length).toBeGreaterThan(0);
    expect(p1).toEqual(p0); // panels used to rank by their own series and disagree
  });
});

describe("layout indexing", () => {
  it("grouped treemap tiles get the rectangle for their own value", () => {
    const scene = buildChart({
      kind: "treemap",
      width: 600,
      height: 400,
      data: { categories: ["G | a", "G | b", "G | c"], series: [{ name: "S", values: [1, 50, 100] }] },
    } as ChartConfig);
    const area = (name: string) => {
      const t = scene.nodes.find((n: any) => n.kind === "rect" && n.name === name) as any;
      return t ? t.w * t.h : 0;
    };
    // values 1 < 50 < 100 must give areas tile-0 < tile-1 < tile-2.
    // (Members are listed ascending, so the sort reorders them — which is
    // exactly when the post-sort-index lookup handed over the wrong rect.)
    expect(area("tile-0")).toBeLessThan(area("tile-1"));
    expect(area("tile-1")).toBeLessThan(area("tile-2"));
  });

  it("stacked connectors join the same series across a zero segment", () => {
    const scene = buildChart({
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [0, 10] },
          { name: "S2", values: [20, 30] },
        ],
      },
      decorations: { connectors: true },
    } as ChartConfig);
    const conns = scene.nodes.filter((n: any) => n.kind === "line" && n.name?.startsWith("connector-")) as any[];
    // S1 is absent in category A, so only S2's boundary can be joined: exactly one
    // connector, and it must link the two S2 tops (20 -> 30+... ), never S1<->S2.
    expect(conns).toHaveLength(1);
    expect(conns[0].name).toBe("connector-0-1"); // series index 1 = S2, not push-order 0
  });

  it("a tight-spread violin category still renders", () => {
    const scene = buildChart({
      kind: "violin",
      width: 480,
      height: 300,
      data: {
        categories: ["Wide", "Tight"],
        series: [
          { name: "o1", values: [0, 3] },
          { name: "o2", values: [30, 3.1] },
          { name: "o3", values: [60, 3] },
          { name: "o4", values: [90, 3.05] },
          { name: "o5", values: [120, 3.02] },
        ],
      },
    } as ChartConfig);
    // Both categories have valid observations; neither may be silently dropped.
    expect(scene.nodes.some((n: any) => n.name === "violin-0")).toBe(true);
    expect(scene.nodes.some((n: any) => n.name === "violin-1")).toBe(true);
  });

  it("horizontal mekko hides labels in rows thinner than the font", () => {
    const cats = Array.from({ length: 16 }, (_, i) => `C${i + 1}`);
    const cfg = {
      kind: "mekko",
      width: 600,
      height: 160,
      horizontal: true,
      data: {
        categories: cats,
        series: [
          { name: "S1", values: cats.map(() => 50) },
          { name: "S2", values: cats.map(() => 50) },
        ],
      },
      decorations: { segmentLabels: true },
    } as ChartConfig;
    const scene = buildChart(cfg);
    const segs = scene.nodes.filter((n: any) => n.kind === "rect" && n.name?.startsWith("seg-")) as any[];
    const labels = scene.nodes.filter((n: any) => n.kind === "text" && n.name?.startsWith("label-")) as any[];
    // Rows here are ~6.6pt thick — far under the 11pt font. The old gate measured
    // the segment's 250pt value-axis length instead, so it stamped a label into
    // every one of them.
    expect(segs.length).toBe(32);
    expect(segs.every((s) => s.h < 11 * 1.25)).toBe(true);
    expect(labels).toHaveLength(0);
    // Vertical mekko is unaffected: its gate always measured r.h.
    const tall = buildChart({ ...cfg, horizontal: false } as ChartConfig);
    expect(tall.nodes.some((n: any) => n.kind === "text" && n.name?.startsWith("label-"))).toBe(true);
  });

  it("horizontal mekko keeps labels out of segments shorter than the font", () => {
    // Thick rows, hairline-short segments — the case the thickness gate alone
    // does not catch. The fit check tolerates 2pt of bleed (the text box is
    // drawn 4pt wider than the segment), which is fine for a vertical mekko's
    // wide columns but let three 3.8pt segments each print a 5.4pt "4",
    // overlapping their neighbours by 1.6pt.
    const scene = buildChart({
      kind: "mekko",
      width: 400,
      height: 300,
      horizontal: true,
      data: {
        categories: ["EMEA"],
        series: [
          { name: "A", values: [4] },
          { name: "B", values: [4] },
          { name: "C", values: [4] },
          { name: "D", values: [300] },
        ],
      },
    } as ChartConfig);
    const ink = (n: any) => {
      const w = textWidth(n.text, n.fontSize);
      const cx = n.x + n.w / 2;
      return { lo: cx - w / 2, hi: cx + w / 2 };
    };
    const labels = scene.nodes.filter((n: any) => n.kind === "text" && n.name?.startsWith("label-")) as any[];
    // Only D is wide enough to carry a label.
    expect(labels.map((l) => l.text)).toEqual(["300"]);
    // And no two labels may ever overlap.
    const inks = labels.map(ink).sort((a, b) => a.lo - b.lo);
    for (let i = 1; i < inks.length; i++) expect(inks[i].lo).toBeGreaterThanOrEqual(inks[i - 1].hi);
  });
});

describe("boxplot extent reports the boxes, not the rows", () => {
  const meanSd = (values: number[]): ChartConfig => ({
    kind: "boxplot",
    ...DEFAULT_SIZE,
    boxplot: { meanSd: true },
    data: { categories: ["A"], series: values.map((v, i) => ({ name: `o${i}`, values: [v] })) },
  });

  it("covers the mean±2·SD whiskers, which sit outside the samples", () => {
    // samples 0..100 → mean 50, sample SD 57.7 → whiskers at -65.5 and 165.5.
    const ext = valueExtent(meanSd([0, 100, 0, 100]))!;
    expect(ext.max).toBeCloseTo(165.47, 1); // was 100 — the largest raw sample
    expect(ext.min).toBeCloseTo(-65.47, 1); // was 0
  });

  it("keeps mean±SD whiskers on the plot when Same scale applies the extent", () => {
    // "Same scale" writes valueExtent's result back as a hard scale override,
    // so an understated extent clipped the whiskers off the plot.
    const cfg = meanSd([0, 100, 0, 100]);
    const ext = valueExtent(cfg)!;
    const scene = buildChart({ ...cfg, scale: { min: ext.min < 0 ? ext.min : undefined, max: ext.max } });
    const ys = scene.nodes.flatMap((n: any) =>
      n.kind === "line" ? [n.y1, n.y2] : n.kind === "rect" ? [n.y, n.y + n.h] : [],
    );
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(-1);
    expect(Math.max(...ys)).toBeLessThanOrEqual(scene.height + 1);
  });

  it("still reports the raw range for a plain Tukey boxplot", () => {
    // Tukey whiskers and outliers are real observations, so nothing moves here.
    const cfg: ChartConfig = {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [10, 20, 30, 40].map((v, i) => ({ name: `o${i}`, values: [v] })) },
    };
    expect(valueExtent(cfg)).toEqual({ min: 0, max: 40 });
  });
});

describe("valueExtent reports what the layout draws", () => {
  /** What "Same scale" does: write the extent back as a hard scale override. */
  const underSameScale = (cfg: ChartConfig) => {
    const e = valueExtent(cfg)!;
    return buildChart({ ...cfg, scale: { min: e.min < 0 ? e.min : undefined, max: e.max } });
  };
  const inkSpan = (scene: { nodes: any[] }) => {
    const ys = scene.nodes.flatMap((n) =>
      n.kind === "line" ? [n.y1, n.y2] : n.kind === "rect" ? [n.y, n.y + n.h] : [],
    );
    return { top: Math.min(...ys), bottom: Math.max(...ys) };
  };

  it("treats an Error row as a whisker, not as a data point", () => {
    const cfg = {
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A"],
        series: [
          { name: "S", values: [10] },
          { name: "Error", values: [30] },
        ],
      },
    } as ChartConfig;
    // The whisker spans 10±30. The old extent was {0,30}: 30 was the Error row's
    // own magnitude mistaken for a value — neither the data range nor the drawn one.
    const ext = valueExtent(cfg)!;
    expect(ext.max).toBeGreaterThanOrEqual(40);
    expect(ext.min).toBeLessThanOrEqual(-20);
    const { top, bottom } = inkSpan(underSameScale(cfg));
    expect(top).toBeGreaterThanOrEqual(-1); // whisker used to reach y=-83
    expect(bottom).toBeLessThanOrEqual(301); // and y=465
  });

  it("covers a waterfall's Target row", () => {
    const cfg = {
      kind: "waterfall",
      width: 480,
      height: 300,
      waterfall: { totalIndices: [2] },
      data: {
        categories: ["Start", "Up", "End"],
        series: [
          { name: "V", values: [100, 20, 0] },
          { name: "Target", values: [null, null, 200] },
        ],
      },
    } as ChartConfig;
    expect(valueExtent(cfg)!.max).toBeGreaterThanOrEqual(200); // was 120 — the running total only
    const { bottom } = inkSpan(underSameScale(cfg));
    expect(bottom).toBeLessThanOrEqual(301);
  });

  it("does not sum a Target row into a stack", () => {
    const cfg = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A"],
        series: [
          { name: "S", values: [10] },
          { name: "Target", values: [5] },
        ],
      },
    } as ChartConfig;
    // The Target is a tick at 5, not another 5pt of stack: the column totals 10.
    expect(valueExtent(cfg)).toEqual({ min: 0, max: 10 });
  });

  it("covers an explicit threshold line above the data", () => {
    const cfg = {
      kind: "clustered",
      width: 480,
      height: 300,
      decorations: { valueLines: [{ mode: "value", value: 500 }] },
      data: { categories: ["A"], series: [{ name: "S", values: [10] }] },
    } as ChartConfig;
    expect(valueExtent(cfg)!.max).toBeGreaterThanOrEqual(500); // was 10
  });

  it("a mean value line needs no widening — it is inside the data by construction", () => {
    const plain = {
      kind: "clustered",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 20] }] },
    } as ChartConfig;
    const withMean = { ...plain, decorations: { valueLines: [{ mode: "mean" }] } } as ChartConfig;
    expect(valueExtent(withMean)).toEqual(valueExtent(plain));
  });
});

describe("reordering categories carries the per-category colors", () => {
  /** The fill of each column, left to right. */
  const fills = (scene: { nodes: any[] }) =>
    scene.nodes
      .filter((n) => n.kind === "rect" && /^(seg|col|bar)-/.test(n.name ?? ""))
      .sort((a, b) => a.x - b.x)
      .map((n) => n.fill);

  const data = {
    categories: ["A", "B", "C"],
    // The highlight is declared on A, the SMALLEST value — so any sort moves it.
    series: [{ name: "S1", values: [10, 50, 30], colors: ["#ff0000", null, null] }],
  };

  it("categorySort moves a highlight with its data point", () => {
    const scene = buildChart({
      kind: "clustered",
      width: 480,
      height: 300,
      categorySort: "descending",
      data,
    } as ChartConfig);
    const order = scene.nodes
      .filter((n: any) => n.kind === "text" && n.name?.startsWith("category-"))
      .sort((a: any, b: any) => a.x - b.x)
      .map((n: any) => n.text);
    expect(order).toEqual(["B", "C", "A"]);
    // Red belongs to A, now rightmost. It used to stay at position 0 and paint B.
    const red = fills(scene)
      .map((f, i) => [i, f])
      .filter(([, f]) => f === "#ff0000");
    expect(red).toEqual([[order.indexOf("A"), "#ff0000"]]);
  });

  it("pareto moves a highlight with its data point", () => {
    const scene = buildChart({ kind: "clustered", width: 480, height: 300, pareto: true, data } as ChartConfig);
    const order = scene.nodes
      .filter((n: any) => n.kind === "text" && n.name?.startsWith("category-"))
      .sort((a: any, b: any) => a.x - b.x)
      .map((n: any) => n.text);
    expect(order).toEqual(["B", "C", "A"]);
    const red = fills(scene)
      .map((f, i) => [i, f])
      .filter(([, f]) => f === "#ff0000");
    expect(red).toEqual([[order.indexOf("A"), "#ff0000"]]);
  });
});

describe("boxplot jitter dots are inside the plot", () => {
  // 19 tight samples plus one at 100: mean ~5.95, SD ~22.1, so the mean±SD box
  // spans only [-38.3, 50.2] and reports no outliers. The jitter overlay still
  // plots the 100.
  const cfg = {
    kind: "boxplot",
    width: 480,
    height: 300,
    boxplot: { meanSd: true, jitter: true },
    data: {
      categories: ["A"],
      series: [...Array(19).fill(1), 100].map((v, i) => ({ name: `o${i}`, values: [v] })),
    },
  } as ChartConfig;

  const dotSpan = (scene: { nodes: any[] }) => {
    const dots = scene.nodes.filter((n) => n.name?.startsWith("dot-"));
    expect(dots.length).toBeGreaterThan(0);
    const ys = dots.map((d: any) => d.cy ?? d.y);
    return { top: Math.min(...ys), bottom: Math.max(...ys) };
  };

  it("the extent covers the samples the overlay draws", () => {
    expect(valueExtent(cfg)!.max).toBeGreaterThanOrEqual(100); // was ~50.2 (mean+2SD)
  });

  it("dots stay on the plot on the auto path", () => {
    const { top, bottom } = dotSpan(buildChart(cfg));
    expect(top).toBeGreaterThanOrEqual(-1); // the 100 sample used to sit at y=-101.6
    expect(bottom).toBeLessThanOrEqual(301);
  });

  it("dots stay on the plot under Same scale", () => {
    const e = valueExtent(cfg)!;
    const scene = buildChart({ ...cfg, scale: { min: e.min < 0 ? e.min : undefined, max: e.max } });
    const { top, bottom } = dotSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1); // and y=-146.0
    expect(bottom).toBeLessThanOrEqual(301);
  });

  it("without jitter the scale still describes the box, not the samples", () => {
    // Nothing plots the far sample, so it must not stretch the axis.
    const noJitter = { ...cfg, boxplot: { meanSd: true } } as ChartConfig;
    expect(valueExtent(noJitter)!.max).toBeLessThan(100);
  });
});

describe("waterfall extent walks the same chain the layout draws", () => {
  const stacked = {
    kind: "waterfall",
    ...DEFAULT_SIZE,
    waterfall: { totalIndices: [3] },
    data: {
      categories: ["FY23", "Organic", "M&A", "FY24"],
      series: [
        { name: "Europe", values: [50, 8, 5, 0] },
        { name: "Americas", values: [36, 6, 9, 0] },
      ],
    },
  } as ChartConfig;

  it("counts every stacked series, not just the first", () => {
    // The bridge runs (50+36) + (8+6) + (5+9) = 114. A second chain that added
    // only series[0] reported 63 — the first series' total, called the chart's.
    expect(valueExtent(stacked)).toEqual({ min: 0, max: 114 });
  });

  it("keeps a stacked bridge on the shape under Same scale", () => {
    const e = valueExtent(stacked)!;
    const scene = buildChart({ ...stacked, scale: { min: e.min < 0 ? e.min : undefined, max: e.max } });
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1); // was -213.8 on a 300pt canvas
    expect(bottom).toBeLessThanOrEqual(scene.height + 1);
  });

  it("carries the running total across a spacer, as the bars do", () => {
    // The old extent ignored spacerIndices entirely; it only got away with it
    // because a spacer's cell is usually null.
    const withSpacer = {
      kind: "waterfall",
      ...DEFAULT_SIZE,
      waterfall: { totalIndices: [4], spacerIndices: [2] },
      data: {
        categories: ["Start", "Up", "", "Up2", "End"],
        series: [{ name: "V", values: [40, 10, 999, 10, 0] }],
      },
    } as ChartConfig;
    // The spacer draws no bar and must not advance the total: 40+10+10 = 60.
    expect(valueExtent(withSpacer)).toEqual({ min: 0, max: 60 });
  });
});

describe('waterfall "of which" detail groups', () => {
  /** Cost (-12) decomposed into Labour/Freight/Energy (-7/-3/-2). */
  const bridge = (detail: boolean) =>
    ({
      kind: "waterfall",
      width: 560,
      height: 300,
      data: {
        categories: ["FY23", "Volume", "Cost", "> Labour", "> Freight", "> Energy", "FX", "FY24"],
        series: [{ name: "Delta", values: [86, 14, -12, -7, -3, -2, -4, 0] }],
      },
      waterfall: {
        totalIndices: [7],
        ...(detail ? { detailGroups: [{ of: 2, indices: [3, 4, 5] }] } : {}),
      },
    }) as ChartConfig;
  const node = (c: ChartConfig, name: string) => buildChart(c).nodes.find((n) => n.name === name) as any;

  it("keeps detail columns off the chain, so the totals stay right", () => {
    // 86 + 14 - 12 - 4 = 84. Without the grouping the details join the walk and
    // the same rows total 72 — the breakdown counted twice.
    expect(node(bridge(true), "label-7").text).toBe("84");
    expect(node(bridge(false), "label-7").text).toBe("72");
  });

  it("decomposes the parent's delta, from the parent's own base", () => {
    const c = bridge(true);
    const cost = node(c, "bar-2");
    const energy = node(c, "bar-5"); // the last detail
    const labour = node(c, "bar-3"); // the first
    // The sub-bridge starts where Cost starts and ends where Cost ends: the
    // group IS that column taken apart, not more steps in the walk.
    expect(labour.y).toBeCloseTo(cost.y, 6);
    expect(energy.y + energy.h).toBeCloseTo(cost.y + cost.h, 6);
  });

  it("steps the connector over the group, without burying it", () => {
    const c = bridge(true);
    const names = buildChart(c)
      .nodes.filter((n) => n.name?.startsWith("connector-"))
      .map((n) => n.name);
    // A detail has no outgoing level to carry, so it draws no connector.
    expect(names).toEqual(["connector-0", "connector-1", "connector-2", "connector-6"]);
    // The parent's connector reaches the next CHAIN column, not the next index.
    const conn = node(c, "connector-2");
    const fx = node(c, "bar-6");
    expect(conn.x2).toBeLessThanOrEqual(fx.x + 0.01);
    expect(conn.x2 - conn.x1).toBeGreaterThan(150); // it spans the whole group
    // Anchoring the sub-bridge at the parent's base (rather than at the level
    // the chain carries) is what keeps this line clear of the bars it skips.
    for (const i of [3, 4, 5]) {
      const b = node(c, `bar-${i}`);
      const through = conn.y1 > b.y && conn.y1 < b.y + b.h;
      expect(through, `bar-${i}`).toBe(false);
    }
  });

  it("renders a group that does not sum to its parent exactly as authored", () => {
    // The engine draws your numbers; it does not reconcile them. The chain is
    // unaffected either way, so the totals cannot silently drift.
    const c = bridge(true);
    (c.data.series[0].values as (number | null)[])[3] = -1; // 1+3+2 != 12
    expect(node(c, "label-7").text).toBe("84");
    expect(node(c, "bar-3")).toBeTruthy();
  });

  it("covers the detail bars in the value extent", () => {
    expect(valueExtent(bridge(true))!.max).toBeGreaterThanOrEqual(100);
  });
});
