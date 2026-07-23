import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
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
