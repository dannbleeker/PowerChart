import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { EllipseNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Bump chart (rank-over-time). */

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
