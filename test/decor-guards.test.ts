import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { buildTableScene } from "../src/core/elements";
import type { ChartConfig } from "../src/core/types";

/** Decoration/annotation guards found by a bug-hunt. */

const anyNaN = (nodes: ReturnType<typeof buildChart>["nodes"]) =>
  nodes.some((n) =>
    Object.entries(n).some(
      ([k, v]) =>
        ["x", "y", "w", "h", "x1", "y1", "x2", "y2", "cx", "cy"].includes(k) &&
        typeof v === "number" &&
        Number.isNaN(v),
    ),
  );

const base: ChartConfig = {
  kind: "stacked",
  width: 480,
  height: 300,
  data: {
    categories: ["A", "B"],
    series: [
      { name: "S1", values: [10, 20] },
      { name: "S2", values: [5, 8] },
    ],
  },
};

describe("callout series index is bounds-checked both ways", () => {
  it("ignores a negative series index instead of emitting NaN geometry", () => {
    const cfg = { ...base, decorations: { callouts: [{ text: "note", category: 0, series: -1 }] } };
    const nodes = buildChart(cfg).nodes;
    expect(anyNaN(nodes)).toBe(false);
    expect(nodes.some((n) => n.kind === "text" && n.text === "note")).toBe(true); // still drawn, anchored to the column top
  });

  it("still anchors a valid series-level callout", () => {
    const cfg = { ...base, decorations: { callouts: [{ text: "note", category: 0, series: 1 }] } };
    const nodes = buildChart(cfg).nodes;
    expect(anyNaN(nodes)).toBe(false);
    expect(nodes.some((n) => n.kind === "text" && n.text === "note")).toBe(true);
  });
});

describe("harvey-ball token ignores a non-numeric fraction", () => {
  const nan = (s: ReturnType<typeof buildTableScene>) =>
    s.nodes.some((n) => Object.values(n).some((v) => typeof v === "number" && Number.isNaN(v)));

  it("does not leak a NaN harvey fraction into geometry for a bare-dot token", () => {
    expect(nan(buildTableScene([["[hb:.] Progress"]], 480))).toBe(false);
  });

  it("still renders a valid harvey fraction (a partial-fill wedge)", () => {
    const half = buildTableScene([["[hb:0.5] Progress"]], 480);
    expect(nan(half)).toBe(false);
    expect(half.nodes.some((n) => n.kind === "wedge")).toBe(true); // the filled arc
  });
});
