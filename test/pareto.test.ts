import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { RectNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Pareto helper (sorted bars + cumulative line). */

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
