import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode } from "../src/core/scene";

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
