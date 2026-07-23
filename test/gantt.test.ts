import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { layoutGantt } from "../src/core/layout/gantt";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { EllipseNode, LineNode, RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig, ChartStyle } from "../src/core/types";

/** Gantt — progress, baselines, critical path, dependencies, gutters, footnote row. */

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

/**
 * Gantt bug-hunt guards: frame clearance (baselines, footnote, gutter columns),
 * spans that collapse to nothing, milestone dependencies, and the two fills that
 * ignored the palette/theme.
 */
const day = (iso: string) => Math.round(Date.parse(iso) / 86400000);

function gantt(partial: Partial<ChartConfig>): ChartConfig {
  return {
    kind: "gantt",
    width: 600,
    height: 220,
    data: { categories: [], series: [] },
    ...partial,
  } as ChartConfig;
}

const node = (nodes: { name?: string }[], name: string) => nodes.find((n) => n.name === name);

/** Channel sum — enough to compare two shades of the same hue. */
const ink = (hex: string) =>
  parseInt(hex.slice(1, 3), 16) + parseInt(hex.slice(3, 5), 16) + parseInt(hex.slice(5, 7), 16);

/** Backlog batch C: gantt progress + baselines, grouped boxplots. */
describe("gantt progress and baselines", () => {
  const cfg: ChartConfig = {
    kind: "gantt",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Design", "Build", "Test"],
      series: [
        { name: "Start", values: [1, 4, 9] },
        { name: "End", values: [4, 10, 13] },
        { name: "% Complete", values: [100, 40, null] },
        { name: "Baseline start", values: [1, 3, 8] },
        { name: "Baseline end", values: [3, 8, 12] },
      ],
    },
  };
  const s = buildChart(cfg);

  it("fills the elapsed share of each bar", () => {
    const bar = s.nodes.find((n) => n.name === "bar-1") as RectNode;
    const progress = s.nodes.find((n) => n.name === "progress-1") as RectNode;
    expect(progress.x).toBeCloseTo(bar.x, 5);
    expect(progress.w / bar.w).toBeCloseTo(0.4, 5);
    expect(progress.h).toBeCloseTo(bar.h, 5);
    // 100% fills fully; missing % draws nothing.
    expect((s.nodes.find((n) => n.name === "progress-0") as RectNode).w).toBeCloseTo(
      (s.nodes.find((n) => n.name === "bar-0") as RectNode).w,
      5,
    );
    expect(s.nodes.some((n) => n.name === "progress-2")).toBe(false);
  });

  it("accepts 0–1 fractions too, clamped", () => {
    const frac = buildChart({
      ...cfg,
      data: {
        ...cfg.data,
        series: [cfg.data.series[0], cfg.data.series[1], { name: "% Complete", values: [0.5, 150, null] }],
      },
    });
    const bar0 = frac.nodes.find((n) => n.name === "bar-0") as RectNode;
    expect((frac.nodes.find((n) => n.name === "progress-0") as RectNode).w / bar0.w).toBeCloseTo(0.5, 5);
    // 150% clamps to full.
    const bar1 = frac.nodes.find((n) => n.name === "bar-1") as RectNode;
    expect((frac.nodes.find((n) => n.name === "progress-1") as RectNode).w / bar1.w).toBeCloseTo(1, 2);
  });

  it("draws thin baseline ghost bars beneath the actual bars", () => {
    const bar = s.nodes.find((n) => n.name === "bar-1") as RectNode;
    const ghost = s.nodes.find((n) => n.name === "gantt-baseline-1") as RectNode;
    expect(ghost.y).toBeGreaterThan(bar.y + bar.h - 1); // below the bar
    expect(ghost.h).toBeLessThan(bar.h); // thinner
    // Baseline spans its own dates (3→8), shifted vs the actual (4→10).
    expect(ghost.x).toBeLessThan(bar.x);
    // The plan-vs-actual slip is visible: ghost ends before the bar does.
    expect(ghost.x + ghost.w).toBeLessThan(bar.x + bar.w);
  });

  it("baseline rows never render as task bars", () => {
    expect(s.nodes.filter((n) => n.name?.startsWith("bar-"))).toHaveLength(6); // 3 bars + 3 bar-labels
  });
});

/**
 * Backlog batch J — more §2 within-kind gaps: gantt auto-summary bars,
 * notched boxplots, radar min–max (peer range + us) band.
 */
describe("gantt auto-summary bars", () => {
  const cfg: ChartConfig = {
    kind: "gantt",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Phase 1", "> Research", "> Interviews", "Phase 2", "> Build"],
      series: [
        { name: "Start", values: [null, 1, 3, null, 6] },
        { name: "End", values: [null, 4, 8, null, 12] },
      ],
    },
    decorations: { summaryBars: true, segmentLabels: false },
  };

  it("draws a capped summary bar spanning the section's children", () => {
    const s = buildChart(cfg);
    const summary = s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "summary-0");
    expect(summary).toBeTruthy();
    // Children of Phase 1 are rows 1–2: min start = 1, max end = 8.
    const r1 = s.nodes.find((n): n is RectNode => n.name === "bar-1")!;
    const r2 = s.nodes.find((n): n is RectNode => n.name === "bar-2")!;
    expect(summary!.x).toBeCloseTo(r1.x, 1); // starts at the earliest child start
    expect(summary!.x + summary!.w).toBeCloseTo(r2.x + r2.w, 1); // ends at the latest child end
    expect(s.nodes.some((n) => n.name === "summary-cap-a-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "summary-cap-b-0")).toBe(true);
  });

  it("no summary bars without the decoration", () => {
    const s = buildChart({ ...cfg, decorations: { segmentLabels: false } });
    expect(s.nodes.some((n) => n.name?.startsWith("summary-"))).toBe(false);
  });
});

/**
 * Batch R — three low/niche §2 gaps: critical-path Gantt highlight, the
 * mean±SD boxplot variant, and sparklines.
 */
describe("critical-path Gantt", () => {
  // Two chains from Research: the long one (Design→Build→QA→Launch) and a
  // short branch (Copywriting). Durations make the long chain critical.
  const cfg: ChartConfig = {
    kind: "gantt",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Research", "Design", "Build", "Copywriting", "QA", "Launch"],
      series: [
        { name: "Start", values: [1, 3, 8, 3, 14, 18] },
        { name: "End", values: [3, 8, 14, 5, 18, 20] },
        { name: "After", values: [null, 1, 2, 1, 3, 5] },
      ],
    },
    decorations: { criticalPath: true },
  };
  const s = buildChart(cfg);
  const bar = (c: number) => s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === `bar-${c}`)!;

  it("outlines the critical activities in red and leaves the branch plain", () => {
    // Critical chain: Research(0) Design(1) Build(2) QA(4) Launch(5).
    for (const c of [0, 1, 2, 4, 5]) {
      expect(bar(c).stroke).toBe("#e34948");
    }
    // Copywriting(3) is on the shorter branch — no critical outline.
    expect(bar(3).stroke).toBeUndefined();
  });

  it("draws the critical dependency arrows thicker and red", () => {
    const critHead = s.nodes.find((n) => n.name === "dep-head-2")!; // Build ← Design edge
    const branchHead = s.nodes.find((n) => n.name === "dep-head-3")!; // Copywriting ← Research edge
    expect((critHead as { fill: string }).fill).toBe("#e34948");
    expect((branchHead as { fill: string }).fill).not.toBe("#e34948");
  });

  it("is a no-op without the decoration", () => {
    const plain = buildChart({ ...cfg, decorations: {} });
    const b0 = plain.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === "bar-0")!;
    expect(b0.stroke).toBeUndefined();
  });
});

describe("gantt holidays & brackets", () => {
  const day = (iso: string) => Math.round(Date.parse(iso) / 86400000);
  const c = cfg({
    kind: "gantt",
    data: {
      categories: ["Build"],
      series: [
        { name: "Start", values: [day("2026-01-05")] },
        { name: "End", values: [day("2026-01-30")] },
        { name: "Holiday", values: [day("2026-01-15")] },
        { name: "Bracket Sprint 1", values: [day("2026-01-05"), day("2026-01-19")] },
      ],
      dates: true,
    },
  });
  const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("shades holidays", () => {
    expect(byName(nodes, "holiday-")).toHaveLength(1);
  });
  it("draws labelled bracket annotations", () => {
    expect(nodes.find((n) => n.name === "bracket-0")).toBeTruthy();
    const label = nodes.find((n) => n.name === "bracket-label-0") as TextNode;
    expect(label.text).toBe("Sprint 1");
  });
});

describe("gantt baselines join the timeline extent", () => {
  const slipped = gantt({
    data: {
      categories: ["Task A"],
      dates: true,
      series: [
        { name: "Start", values: [day("2026-02-02")] },
        { name: "End", values: [day("2026-02-20")] },
        { name: "Baseline start", values: [day("2026-01-19")] },
        { name: "Baseline end", values: [day("2026-02-06")] },
      ],
    },
  });

  it("keeps the ghost bar of a slipped plan inside the plot", () => {
    const { nodes, anchors } = layoutGantt(slipped, DEFAULT_STYLE, DEFAULT_DECOR);
    const ghost = node(nodes, "gantt-baseline-0") as RectNode;
    const plot = anchors.plot!;
    // The baseline started two weeks before the actual: excluded from the
    // extent it was drawn at x=-95.5, through the label gutter and off-canvas.
    expect(ghost.x).toBeGreaterThanOrEqual(plot.x - 0.01);
    expect(ghost.x + ghost.w).toBeLessThanOrEqual(plot.x + plot.w + 0.01);
    expect(ghost.x).toBeLessThan((node(nodes, "bar-0") as RectNode).x);
  });

  it("still starts the timeline at the actual span when the baseline is inside it", () => {
    // A baseline within the actual dates must not widen the timeline.
    const inside: ChartConfig = {
      ...slipped,
      data: {
        ...slipped.data,
        series: [
          slipped.data.series[0],
          slipped.data.series[1],
          { name: "Baseline start", values: [day("2026-02-05")] },
          { name: "Baseline end", values: [day("2026-02-15")] },
        ],
      },
    };
    const bare: ChartConfig = { ...slipped, data: { ...slipped.data, series: slipped.data.series.slice(0, 2) } };
    const barX = (c: ChartConfig) => (node(layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR).nodes, "bar-0") as RectNode).x;
    expect(barX(inside)).toBeCloseTo(barX(bare), 6);
  });
});

describe("gantt reserves the footnote row", () => {
  const tasks = (n: number) => ({
    categories: Array.from({ length: n }, (_, i) => `Task ${i + 1}`),
    series: [
      { name: "Start", values: Array.from({ length: n }, (_, i) => i) },
      { name: "End", values: Array.from({ length: n }, (_, i) => i + 2) },
    ],
  });

  it("keeps the Today label off the footnote line", () => {
    const scene = buildChart(
      gantt({
        footnote: "Source: PMO plan v4",
        data: {
          categories: ["Design", "Build", "Test"],
          dates: true,
          series: [
            {
              name: "Start",
              values: [day("2026-01-05"), day("2026-01-12"), day("2026-01-19")],
            },
            { name: "End", values: [day("2026-01-12"), day("2026-01-19"), day("2026-01-26")] },
            { name: "Today", values: [day("2026-01-12"), null, null] },
          ],
        },
      }),
    );
    const foot = node(scene.nodes, "footnote") as TextNode;
    const today = node(scene.nodes, "today-label") as TextNode;
    expect(foot).toBeTruthy();
    expect(today.y + today.h).toBeLessThanOrEqual(foot.y + 0.01);
  });

  it("keeps the last row off the footnote line", () => {
    const scene = buildChart(gantt({ footnote: "Source: PMO plan v4", data: tasks(10) }));
    const foot = node(scene.nodes, "footnote") as TextNode;
    const bar = node(scene.nodes, "bar-9") as RectNode;
    const label = node(scene.nodes, "category-9") as TextNode;
    expect(bar.y + bar.h).toBeLessThanOrEqual(foot.y + 0.01);
    expect(label.y + label.h).toBeLessThanOrEqual(foot.y + 0.01);
  });

  it("gives the plot the space back when there is no footnote", () => {
    const plotH = (c: ChartConfig) => layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR).anchors.plot!.h;
    expect(plotH(gantt({ data: tasks(3) }))).toBeGreaterThan(
      plotH(gantt({ footnote: "Source: PMO plan v4", data: tasks(3) })),
    );
  });
});

describe("gantt spans that collapse to nothing", () => {
  // A one-day cutover typed with the same date twice, a task typed
  // end-before-start, and a normal 4-unit span for scale.
  const { nodes } = layoutGantt(
    gantt({
      data: {
        categories: ["Cutover", "Review", "Normal"],
        series: [
          { name: "Start", values: [15, 20, 5] },
          { name: "End", values: [15, 19, 9] },
        ],
      },
    }),
    DEFAULT_STYLE,
    DEFAULT_DECOR,
  );
  const bar = (c: number) => node(nodes, `bar-${c}`) as RectNode;

  it("draws a hairline for a single-day activity instead of an empty row", () => {
    expect(bar(0)).toBeTruthy();
    expect(bar(0).w).toBeGreaterThan(0);
    expect(bar(0).w).toBeLessThan(3); // a hairline, not a fabricated span
  });

  it("draws an end-before-start pair as the span it names", () => {
    // 19→20 is a quarter of Normal's 5→9, drawn from the earlier date.
    expect(bar(1)).toBeTruthy();
    expect(bar(1).w).toBeCloseTo(bar(2).w / 4, 5);
    expect(bar(1).x + bar(1).w).toBeCloseTo(bar(0).x + (bar(2).w / 4) * 5, 5);
  });
});

describe("gantt dependencies on a milestone row", () => {
  const { nodes } = layoutGantt(
    gantt({
      data: {
        categories: ["Kickoff", "Build"],
        series: [
          { name: "Milestone", values: [0, null] },
          { name: "Start", values: [null, 5] },
          { name: "End", values: [null, 10] },
          { name: "After", values: [null, 1] },
        ],
      },
    }),
    DEFAULT_STYLE,
    DEFAULT_DECOR,
  );

  it("anchors the arrow on the gate when the predecessor has no End", () => {
    const gate = node(nodes, "milestone-0") as EllipseNode;
    const vertical = node(nodes, "dep-v-1") as LineNode;
    expect(vertical).toBeTruthy();
    expect(node(nodes, "dep-h-1")).toBeTruthy();
    expect(node(nodes, "dep-head-1")).toBeTruthy();
    expect(vertical.x1).toBeCloseTo(gate.cx, 5);
  });
});

describe("gantt gutter columns keep the timeline a plot", () => {
  const packed = gantt({
    width: 480,
    height: 240,
    data: {
      categories: [
        "A very long activity name here | Owner Name | some remark",
        "Another long activity | Owner Two | x",
      ],
      series: [
        { name: "Start", values: [0, 4] },
        { name: "End", values: [4, 10] },
        ...["Cost", "FTE", "Risk", "Budget", "Actual", "Var"].map((n) => ({
          name: `Column ${n}`,
          values: [123456, 234567],
        })),
      ],
    },
  });

  it("drops the columns that do not fit rather than inverting the plot", () => {
    const { nodes, anchors } = layoutGantt(packed, DEFAULT_STYLE, DEFAULT_DECOR);
    const plot = anchors.plot!;
    expect(plot.w).toBeGreaterThan(0);
    const heads = nodes.filter((n) => /^col-head-\d+$/.test(n.name ?? ""));
    expect(heads.length).toBeGreaterThan(0);
    expect(heads.length).toBeLessThan(6); // the overflowing ones are dropped
    // Every drawn column still sits left of the timeline, in order.
    for (const h of heads) expect((h as TextNode).x).toBeLessThan(plot.x);
    for (const c of [0, 1]) {
      const b = node(nodes, `bar-${c}`) as RectNode;
      expect(b.x).toBeGreaterThanOrEqual(plot.x - 0.01);
      expect(b.w).toBeGreaterThan(0);
    }
  });

  it("leaves a plan whose gutters fit untouched", () => {
    const roomy: ChartConfig = { ...packed, width: 900 };
    const { nodes } = layoutGantt(roomy, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(nodes.filter((n) => /^col-head-\d+$/.test(n.name ?? ""))).toHaveLength(6);
  });
});

describe("gantt bar fills follow the palette and the theme", () => {
  const withProgress = gantt({
    data: {
      categories: ["Design"],
      series: [
        { name: "Start", values: [0] },
        { name: "End", values: [10] },
        { name: "% Complete", values: [50] },
        { name: "Baseline start", values: [0] },
        { name: "Baseline end", values: [8] },
      ],
    },
  });
  const styled = (over: Partial<ChartStyle>): ChartStyle => ({ ...DEFAULT_STYLE, ...over });
  const run = (over: Partial<ChartStyle> = {}) => layoutGantt(withProgress, styled(over), DEFAULT_DECOR).nodes;

  it("inks the bar label for the bar it sits on", () => {
    expect((node(run(), "bar-label-0") as TextNode).color).toBe("#ffffff"); // default blue
    expect((node(run({ palette: ["#ffe066"] }), "bar-label-0") as TextNode).color).toBe("#0b0b0b");
  });

  it("derives the percent-complete fill from the bar, darker on a light canvas", () => {
    const dflt = run();
    const progress = (node(dflt, "progress-0") as RectNode).fill;
    const bar = (node(dflt, "bar-0") as RectNode).fill;
    expect(progress).not.toBe("#1b4e8a"); // no longer a fixed blue
    expect(ink(progress)).toBeLessThan(ink(bar));
    // On a red palette it stays a red, not a blue block inside a red bar.
    const red = run({ palette: ["#e34948"] });
    const rp = (node(red, "progress-0") as RectNode).fill;
    const [r, , b] = [rp.slice(1, 3), rp.slice(3, 5), rp.slice(5, 7)].map((h) => parseInt(h, 16));
    expect(r).toBeGreaterThan(b);
    // On a dark canvas "denser" is brighter — never darker than the canvas.
    const dark = run({ background: "#1a1a1a", text: "#f5f5f5" });
    expect(ink((node(dark, "progress-0") as RectNode).fill)).toBeGreaterThan(
      ink((node(dark, "bar-0") as RectNode).fill),
    );
  });

  it("adapts the baseline ghost to the canvas like every other tint", () => {
    expect((node(run(), "gantt-baseline-0") as RectNode).fill).toBe("#cfcdc5"); // light theme unchanged
    const onDark = (node(run({ background: "#1a1a1a", text: "#f5f5f5" }), "gantt-baseline-0") as RectNode).fill;
    expect(onDark).not.toBe("#cfcdc5");
    expect(ink(onDark)).toBeLessThan(ink("#cfcdc5"));
    expect(ink(onDark)).toBeGreaterThan(ink("#1a1a1a"));
  });
});
