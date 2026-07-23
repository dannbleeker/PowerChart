import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { WedgeNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Semi-circle gauge. */

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
