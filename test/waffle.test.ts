import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { NO_DATA } from "../src/core/color";
import { sampleConfig } from "../src/core/samples";
import { PALETTE } from "../src/core/style";
import type { RectNode, TextNode } from "../src/core/scene";

/** Waffle grid. */

describe("waffle chart", () => {
  it("fills exactly the rounded share of 100 cells per category, gray remainder", () => {
    const s = buildChart({
      kind: "waffle",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B", "C"],
        series: [{ name: "Share", values: [45, 30, 25] }],
      },
    });
    const cells = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("waffle-cell"));
    expect(cells).toHaveLength(100);
    expect(cells.filter((c) => c.fill === PALETTE[0])).toHaveLength(45);
    expect(cells.filter((c) => c.fill === PALETTE[1])).toHaveLength(30);
    expect(cells.filter((c) => c.fill === PALETTE[2])).toHaveLength(25);
    expect(cells.filter((c) => c.fill === NO_DATA)).toHaveLength(0);
  });

  it("single category reads as a literal % with a big-number legend", () => {
    const s = buildChart(sampleConfig("waffle")); // Subscription 68
    const cells = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("waffle-cell"));
    expect(cells.filter((c) => c.fill === PALETTE[0])).toHaveLength(68);
    expect(cells.filter((c) => c.fill === NO_DATA)).toHaveLength(32);
    expect((s.nodes.find((n) => n.name === "waffle-big-pct") as TextNode).text).toBe("68%");
    // Fill starts at the bottom-left: cell 0 is on the lowest row, leftmost.
    const c0 = cells.find((c) => c.name === "waffle-cell-0")!;
    expect(Math.max(...cells.map((c) => c.y))).toBeCloseTo(c0.y, 5);
    expect(Math.min(...cells.map((c) => c.x))).toBeCloseTo(c0.x, 5);
  });

  it("largest remainder keeps the filled total exact; 100%= overrides the denominator", () => {
    const thirds = buildChart({
      kind: "waffle",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B", "C"],
        series: [{ name: "Share", values: [1, 1, 1] }],
      },
    });
    const filled = thirds.nodes.filter(
      (n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("waffle-cell") && n.fill !== NO_DATA,
    );
    expect(filled).toHaveLength(100);
    const scaled = buildChart({
      kind: "waffle",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Won"],
        series: [{ name: "Deals", values: [50] }],
        hundredPercent: [200],
      },
    });
    const won = scaled.nodes.filter(
      (n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("waffle-cell") && n.fill !== NO_DATA,
    );
    expect(won).toHaveLength(25);
    expect((scaled.nodes.find((n) => n.name === "waffle-big-pct") as TextNode).text).toBe("25%");
  });
});
