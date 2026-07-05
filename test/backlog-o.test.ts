import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, RectNode, WedgeNode } from "../src/core/scene";

/**
 * Backlog §3 (new chart types), batch O: semi-circle gauge, Pareto helper,
 * bump chart.
 */

describe("semi-circle gauge", () => {
  const base: ChartConfig = {
    kind: "doughnut",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "S", values: [40, 30, 20, 10] }] },
    decorations: { segmentLabels: false },
  };

  it("fills a 180° arc with a doughnut hole and a centre total", () => {
    const s = buildChart({ ...base, pie: { semi: true } });
    const wedges = s.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    expect(wedges).toHaveLength(4);
    const span = wedges.reduce((a, w) => a + (w.endAngle - w.startAngle), 0);
    expect(span).toBeCloseTo(180, 1); // half circle
    expect(wedges.every((w) => w.innerR > 0)).toBe(true); // it's a doughnut
    expect(s.nodes.some((n) => n.name === "gauge-total")).toBe(true);
  });

  it("a plain doughnut is still a full 360° ring", () => {
    const s = buildChart(base);
    const wedges = s.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    const span = wedges.reduce((a, w) => a + (w.endAngle - w.startAngle), 0);
    expect(span).toBeCloseTo(360, 1);
  });
});

describe("Pareto helper", () => {
  const cfg: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    pareto: true,
    data: {
      categories: ["Low", "High", "Mid"],
      series: [{ name: "Count", values: [10, 60, 30] }],
    },
    decorations: { segmentLabels: false },
  };

  it("sorts descending and overlays a cumulative-% line on a secondary axis", () => {
    const s = buildChart(cfg);
    // Became a combo: bars + a cumulative line + secondary axis.
    expect(s.nodes.some((n) => n.name?.startsWith("combo-marker-0-"))).toBe(true);
    expect(s.nodes.some((n) => n.name === "secondary-axis")).toBe(true);
    // Bars sorted descending: first bar (High=60) taller than last (Low=10).
    const first = s.nodes.find((n): n is RectNode => n.name === "seg-0-0")!;
    const last = s.nodes.find((n): n is RectNode => n.name === "seg-0-2")!;
    expect(first.h).toBeGreaterThan(last.h);
    // Six-per-category not asserted; the cumulative line has one marker per bar.
    expect(s.nodes.filter((n) => n.name?.startsWith("combo-marker-0-"))).toHaveLength(3);
  });

  it("no-op without the pareto flag", () => {
    const s = buildChart({ ...cfg, pareto: undefined });
    expect(s.nodes.some((n) => n.name?.startsWith("combo-marker-"))).toBe(false);
  });
});

describe("bump chart", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    decorations: { bump: true },
    data: {
      categories: ["Y1", "Y2", "Y3"],
      series: [
        { name: "Top", values: [1, 1, 2] },
        { name: "Bottom", values: [3, 3, 3] },
        { name: "Mid", values: [2, 2, 1] },
      ],
    },
  };
  const s = buildChart(cfg);

  it("draws inverted rank lines (rank 1 highest) with markers and end labels", () => {
    // Rank 1 sits above rank 3 (smaller y).
    const top = s.nodes.find((n): n is EllipseNode => n.name === "bump-marker-0-0")!; // Top, rank 1
    const bottom = s.nodes.find((n): n is EllipseNode => n.name === "bump-marker-1-0")!; // Bottom, rank 3
    expect(top.cy).toBeLessThan(bottom.cy);
    // Thick connecting lines and both-end name labels.
    expect(s.nodes.some((n) => n.name === "bump-0-1")).toBe(true);
    expect(s.nodes.some((n) => n.name === "bump-label-l-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "bump-label-r-0")).toBe(true);
  });

  it("period headers label each category", () => {
    expect(s.nodes.some((n) => n.name === "period-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "period-2")).toBe(true);
  });
});
