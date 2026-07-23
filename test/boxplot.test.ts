import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart, valueExtent } from "../src/core/chart";
import { boxplotExtent, layoutBoxplot } from "../src/core/layout/boxplot";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { EllipseNode, LineNode, PolygonNode, RectNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Box plots — grouped, notched, jittered dots, mean±SD, data-driven domain. */

describe("grouped boxplots", () => {
  const cfg: ChartConfig = {
    kind: "boxplot",
    ...DEFAULT_SIZE,
    data: {
      categories: ["North", "South"],
      series: [
        { name: "Min | 2024", values: [2, 3] },
        { name: "Q1 | 2024", values: [3, 5] },
        { name: "Median | 2024", values: [4, 7] },
        { name: "Q3 | 2024", values: [6, 9] },
        { name: "Max | 2024", values: [8, 12] },
        { name: "Min | 2025", values: [3, 4] },
        { name: "Q1 | 2025", values: [4, 6] },
        { name: "Median | 2025", values: [5, 8] },
        { name: "Q3 | 2025", values: [7, 10] },
        { name: "Max | 2025", values: [9, 13] },
      ],
    },
  };
  const s = buildChart(cfg);

  it("draws one box per group per category, side by side on one shared scale", () => {
    const b24 = s.nodes.find((n) => n.name === "box-0-g0") as RectNode;
    const b25 = s.nodes.find((n) => n.name === "box-0-g1") as RectNode;
    expect(b24).toBeDefined();
    expect(b25).toBeDefined();
    expect(b24.x + b24.w).toBeLessThanOrEqual(b25.x + 1); // side by side, no overlap
    // Shared scale: 2025's IQR (4–7) sits higher than 2024's (3–6).
    expect(b25.y).toBeLessThan(b24.y);
    // Distinct group colors.
    expect(b24.stroke).not.toBe(b25.stroke);
  });

  it("adds a group legend and keeps ungrouped sheets on the single-box path", () => {
    expect(s.nodes.filter((n) => n.name?.startsWith("legend-chip"))).toHaveLength(2);
    const single = buildChart({
      ...cfg,
      data: {
        categories: ["A"],
        series: [
          { name: "Min", values: [2] },
          { name: "Q1", values: [3] },
          { name: "Median", values: [4] },
          { name: "Q3", values: [6] },
          { name: "Max", values: [8] },
        ],
      },
    });
    expect(single.nodes.some((n) => n.name === "box-0")).toBe(true); // un-suffixed names
    expect(single.nodes.some((n) => n.name?.startsWith("legend-chip"))).toBe(false);
  });

  it("groups raw-sample rows by the same pipe suffix", () => {
    const raw = buildChart({
      ...cfg,
      data: {
        categories: ["A"],
        series: [
          ...[4, 5, 6, 7, 8].map((v) => ({ name: "obs | X", values: [v] })),
          ...[8, 9, 10, 11, 12].map((v) => ({ name: "obs | Y", values: [v] })),
        ],
      },
    });
    const bx = raw.nodes.find((n) => n.name === "box-0-g0") as RectNode;
    const by = raw.nodes.find((n) => n.name === "box-0-g1") as RectNode;
    expect(bx).toBeDefined();
    expect(by.y).toBeLessThan(bx.y); // Y's higher values sit higher
  });
});

describe("boxplot jittered dots", () => {
  const rows = Array.from({ length: 12 }, (_, i) => ({
    name: `s${i + 1}`,
    values: [3 + (i % 5), 5 + (i % 4), 4 + (i % 6)],
  }));
  const base: ChartConfig = {
    kind: "boxplot",
    ...DEFAULT_SIZE,
    data: { categories: ["North", "South", "East"], series: rows },
    decorations: { categoryAxis: true, valueAxis: true },
  };

  it("adds no dots by default", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.name?.startsWith("dot-"))).toBe(false);
  });

  it("overlays one jittered dot per observation, spread across the box", () => {
    const s = buildChart({ ...base, boxplot: { jitter: true } });
    const dots = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("dot-0-"));
    expect(dots).toHaveLength(12); // one per raw observation in category 0
    // Deterministic jitter spreads them horizontally (not all on center).
    const xs = new Set(dots.map((d) => Math.round(d.cx * 10)));
    expect(xs.size).toBeGreaterThan(4);
    // Jitter subsumes the separate outlier dots.
    expect(s.nodes.some((n) => n.name?.startsWith("outlier-"))).toBe(false);
  });

  it("is deterministic (same layout twice)", () => {
    const a = buildChart({ ...base, boxplot: { jitter: true } }).nodes.filter((n) => n.name?.startsWith("dot-"));
    const b = buildChart({ ...base, boxplot: { jitter: true } }).nodes.filter((n) => n.name?.startsWith("dot-"));
    expect(JSON.stringify(a)).toBe(JSON.stringify(b));
  });
});

describe("notched boxplots", () => {
  const rows = Array.from({ length: 10 }, (_, i) => ({
    name: `o${i + 1}`,
    values: [10 + i, 12 + i * 1.5, 20 + i * 2],
  }));
  const base: ChartConfig = {
    kind: "boxplot",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: rows },
    decorations: { categoryAxis: true, valueAxis: true },
  };

  it("renders the box as a 10-point notched polygon in raw-sample mode", () => {
    const s = buildChart({ ...base, boxplot: { notch: true } });
    const box = s.nodes.find((n): n is PolygonNode => n.kind === "polygon" && n.name === "box-0");
    expect(box).toBeTruthy();
    expect(box!.points).toHaveLength(10);
    // The notch pinches inward at the median: mid-height x-extent < box extent.
    const xs = box!.points.map((p) => p.x);
    const boxExtent = Math.max(...xs) - Math.min(...xs);
    const notchXs = box!.points.filter((_, i) => i === 3 || i === 8).map((p) => p.x);
    expect(Math.abs(notchXs[0] - notchXs[1])).toBeLessThan(boxExtent);
  });

  it("plain box (rect) without the notch flag", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.kind === "rect" && n.name === "box-0")).toBe(true);
    expect(s.nodes.some((n) => n.kind === "polygon" && n.name === "box-0")).toBe(false);
  });

  it("precomputed boxes (no sample size) stay rectangular even with notch on", () => {
    const s = buildChart({
      kind: "boxplot",
      ...DEFAULT_SIZE,
      boxplot: { notch: true },
      data: {
        categories: ["X"],
        series: [
          { name: "Min", values: [2] },
          { name: "Q1", values: [4] },
          { name: "Median", values: [6] },
          { name: "Q3", values: [9] },
          { name: "Max", values: [12] },
        ],
      },
    });
    expect(s.nodes.some((n) => n.kind === "rect" && n.name === "box-0")).toBe(true);
  });
});

describe("mean±SD boxplot", () => {
  // Symmetric sample around 10: mean 10, sample SD 2.
  const cfg: ChartConfig = {
    kind: "boxplot",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A"],
      series: [
        { name: "o1", values: [8] },
        { name: "o2", values: [10] },
        { name: "o3", values: [12] },
        { name: "o4", values: [8] },
        { name: "o5", values: [10] },
        { name: "o6", values: [12] },
      ],
    },
    boxplot: { meanSd: true },
    decorations: { valueAxis: true },
  };
  const s = buildChart(cfg);
  const valueToY = s.nodes.length; // touch to avoid unused; real check below

  it("centres the box on the mean with a ±1·SD body", () => {
    const box = s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "box-0")!;
    const median = s.nodes.find((n): n is LineNode => n.kind === "line" && n.name === "median-0")!;
    // Median line is the mean (y at value 10); the box spans q1..q3 = mean±SD.
    const meanY = median.y1;
    // Box top corresponds to the higher value (mean+SD=12), so box.y < meanY.
    expect(box.y).toBeLessThan(meanY);
    expect(box.y + box.h).toBeGreaterThan(meanY);
    // Symmetric: mean line sits at the vertical centre of the box (±1px).
    expect(Math.abs(box.y + box.h / 2 - meanY)).toBeLessThan(1);
  });

  it("has no mean × marker (the centre line already is the mean)", () => {
    expect(s.nodes.some((n) => n.name?.startsWith("mean-"))).toBe(false);
    expect(valueToY).toBeGreaterThan(0);
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

describe("boxplot data-driven domain (no forced zero)", () => {
  it("does not pin the value axis to zero for far-from-zero data", () => {
    const cfg: ChartConfig = {
      kind: "boxplot",
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "o1", values: [40, 55] },
          { name: "o2", values: [60, 80] },
          { name: "o3", values: [50, 95] },
        ],
      },
    };
    const ext = boxplotExtent(cfg)!;
    // The old code forced min:0; the domain must now start near the data (40).
    expect(ext.min).toBeGreaterThan(30);
    expect(ext.max).toBeLessThanOrEqual(95);
  });
});

describe("grouped boxplot reserves its wrapped group-legend rows", () => {
  // Grouped boxplots draw their own wrapping group legend; the frame must drop
  // the boxes below the extra rows the same way the column frame does.
  const boxplotCfg = (groups: string[]): ChartConfig => ({
    kind: "boxplot",
    width: 320,
    height: 320,
    data: {
      categories: ["Scores"],
      series: groups.flatMap((g) => [
        { name: `Min | ${g}`, values: [10] },
        { name: `Q1 | ${g}`, values: [20] },
        { name: `Median | ${g}`, values: [30] },
        { name: `Q3 | ${g}`, values: [40] },
        { name: `Max | ${g}`, values: [50] },
      ]),
    },
  });
  const laid = (cfg: ChartConfig) => layoutBoxplot(cfg, DEFAULT_STYLE, DEFAULT_DECOR);
  const chipRowCount = (cfg: ChartConfig) =>
    new Set(
      laid(cfg)
        .nodes.filter((n) => n.name?.startsWith("legend-chip-"))
        .map((n) => (n as RectNode).y),
    ).size;
  const rowH = DEFAULT_STYLE.fontSize * 1.6;

  it("drops the plot exactly one row per wrapped legend row (nothing extra at one row)", () => {
    const oneRow = boxplotCfg(["A", "B"]); // two short groups fit a single row
    const wrapped = boxplotCfg([
      "Northern Europe Wholesale",
      "Central Europe Retail Partners",
      "Southern Europe Online Direct",
      "Nordics Consumer Marketplace",
    ]);
    expect(chipRowCount(oneRow)).toBe(1);
    const rows = chipRowCount(wrapped);
    expect(rows).toBeGreaterThan(1);
    // The wrapped plot top sits exactly (rows-1) legend rows below the one-row
    // plot top — the extra rows are reserved, not painted over the boxes. At one
    // row the reservation is zero, so existing decks stay byte-identical.
    expect(laid(wrapped).anchors.plot.y - laid(oneRow).anchors.plot.y).toBeCloseTo((rows - 1) * rowH, 5);
  });
});
