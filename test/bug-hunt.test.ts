import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE, valueExtent } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";
import { formatNumber, parseDateToken } from "../src/core/format";
import type { ChartConfig } from "../src/core/types";

/** Regression tests for defects found during a codebase bug-hunt pass. */

describe("full-circle pie / doughnut (single slice)", () => {
  it("renders a visible circle path for a 360° pie", () => {
    const cfg: ChartConfig = {
      kind: "pie",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [100] }] },
    };
    const scene = buildChart(cfg);
    const wedge = scene.nodes.find((n) => n.kind === "wedge");
    expect(wedge).toBeTruthy();
    // The lone slice spans the whole circle.
    expect((wedge as any).endAngle - (wedge as any).startAngle).toBeCloseTo(360);
    const svg = sceneToSvg(scene);
    // Must emit a real path with area — the old code drew a degenerate arc
    // between coincident points, producing nothing.
    expect(svg).toMatch(/fill-rule="evenodd"/);
    expect(svg).toMatch(/<path[^>]*A /);
  });

  it("renders a full disc plus an overpainted hole for a 360° doughnut", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [100] }] },
    };
    const scene = buildChart(cfg);
    // The lone slice is a full-circle wedge (the hole is faked with an
    // overpainted background ellipse for Office.js compatibility).
    const wedge = scene.nodes.find((n) => n.kind === "wedge") as any;
    expect(wedge.endAngle - wedge.startAngle).toBeCloseTo(360);
    expect(scene.nodes.some((n) => n.kind === "ellipse" && (n as any).name === "hole")).toBe(true);
    const svg = sceneToSvg(scene);
    expect(svg).toMatch(/fill-rule="evenodd"/);
    // A real disc, not a degenerate zero-area arc.
    const path = svg.match(/<path d="([^"]*)"[^>]*fill-rule="evenodd"/)!;
    expect(path[1]).not.toMatch(/NaN/);
    expect((path[1].match(/A /g) || []).length).toBeGreaterThanOrEqual(2);
  });
});

describe("formatNumber negative-zero", () => {
  it("normalises a rounded -0 to 0", () => {
    expect(formatNumber(-0.4, { decimals: 0 })).toBe("0");
    expect(formatNumber(-0.001, { decimals: 0 })).toBe("0");
  });
  it("keeps the sign on genuine negatives", () => {
    expect(formatNumber(-5.2)).toBe("-5.2");
    expect(formatNumber(-0.4, { decimals: 0, forceSign: true })).toBe("0");
  });
});

describe("parseDateToken numeric ranges", () => {
  it("rejects hyphenated numeric ranges as category labels", () => {
    expect(parseDateToken("3-5")).toBeNull();
    expect(parseDateToken("10-20")).toBeNull();
    expect(parseDateToken("18–24")).toBeNull(); // en dash
  });
  it("still parses real dates", () => {
    expect(parseDateToken("2026-01-15")).toBeTypeOf("number");
    expect(parseDateToken("2026-01")).toBeTypeOf("number");
  });
});

describe("categorySort excludes carried rows", () => {
  it("ranks by real stack totals, ignoring Error/Target rows", () => {
    // Category B has the larger real total (30 vs 20) but a huge Target on A.
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      categorySort: "descending",
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Value", values: [20, 30] },
          { name: "Target", values: [999, 1] },
        ],
      },
    };
    const scene = buildChart(cfg);
    // Descending by real total → B (30) before A (20). Find the category axis
    // labels in order.
    const labels = scene.nodes
      .filter((n) => n.kind === "text" && (n as any).name?.startsWith("category-"))
      .map((n) => (n as any).text);
    expect(labels[0]).toBe("B");
    expect(labels[1]).toBe("A");
  });
});

describe("degenerate inputs do not throw", () => {
  it("boxplot with an empty series", () => {
    const cfg: ChartConfig = {
      kind: "boxplot",
      ...DEFAULT_SIZE,
      data: { categories: ["A"], series: [{ name: "S", values: [] }] },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
  it("area chart with a single empty category set", () => {
    const cfg: ChartConfig = {
      kind: "area",
      ...DEFAULT_SIZE,
      decorations: { seriesLabels: true },
      data: { categories: [], series: [{ name: "S", values: [] }] },
    };
    expect(() => buildChart(cfg)).not.toThrow();
  });
});

/** Helper: the vertical span of every rect in a scene. */
const rectSpan = (scene: { nodes: any[] }) => {
  const ys = scene.nodes.flatMap((n) => (n.kind === "rect" ? [n.y, n.y + n.h] : []));
  return { top: Math.min(...ys), bottom: Math.max(...ys) };
};

describe("value extents and auto-scales", () => {
  it("area extent honours the negative stack, like stacked (was floored at 0)", () => {
    const data = {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "A", values: [10, -40, 20] },
        { name: "B", values: [5, -30, 10] },
      ],
    };
    expect(valueExtent({ kind: "area", ...DEFAULT_SIZE, data } as ChartConfig)).toEqual({ min: -70, max: 30 });
    // identical data must give the same extent as the stacked sibling
    expect(valueExtent({ kind: "area", ...DEFAULT_SIZE, data } as ChartConfig)).toEqual(
      valueExtent({ kind: "stacked", ...DEFAULT_SIZE, data } as ChartConfig),
    );
  });

  it("negative small-multiples areas stay inside the scene", () => {
    const scene = buildChart({
      kind: "area", width: 480, height: 300, multiples: {},
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [{ name: "A", values: [10, -40, 20] }, { name: "B", values: [5, -30, 10] }],
      },
    } as ChartConfig);
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1);
    expect(bottom).toBeLessThanOrEqual(301); // was ~775 — rendered far below the canvas
  });

  it("a target below the data range widens the scale down, not just up", () => {
    const scene = buildChart({
      kind: "clustered", width: 480, height: 300,
      data: {
        categories: ["A", "B", "C"],
        series: [{ name: "Actual", values: [10, 20, 30] }, { name: "Target", values: [-50, -50, -50] }],
      },
    } as ChartConfig);
    const ticks = scene.nodes.filter((n: any) => n.kind === "line" && n.name?.startsWith("target-"));
    expect(ticks.length).toBe(3);
    for (const t of ticks as any[]) {
      expect(t.y1).toBeGreaterThanOrEqual(0);
      expect(t.y1).toBeLessThanOrEqual(300); // target tick used to land outside the plot
    }
  });

  it("a multi-series waterfall combo keeps its columns on canvas", () => {
    const scene = buildChart({
      kind: "combo", ...DEFAULT_SIZE, combo: { columns: "waterfall" },
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "W1", values: [10, 10, 10] },
          { name: "W2", values: [10, 10, 10] },
          { name: "L", type: "line", values: [40, 40, 40] },
        ],
      },
    } as ChartConfig);
    const { top, bottom } = rectSpan(scene);
    expect(top).toBeGreaterThanOrEqual(-1); // was -129 — stacked peak understated
    expect(bottom).toBeLessThanOrEqual(scene.height + 1);
  });

  it("small-multiples panels share one category order under categorySort", () => {
    const scene = buildChart({
      kind: "clustered", width: 480, height: 300, multiples: { columns: 2 }, categorySort: "descending",
      data: {
        categories: ["A", "B", "C"],
        series: [{ name: "S1", values: [1, 9, 5] }, { name: "S2", values: [9, 1, 5] }],
      },
      decorations: { categoryAxis: true },
    } as ChartConfig);
    const axisOf = (p: string) =>
      scene.nodes
        .filter((n: any) => n.kind === "text" && n.name?.startsWith(`${p}-category-`))
        .sort((a: any, b: any) => a.x - b.x)
        .map((n: any) => n.text);
    const p0 = axisOf("p0"), p1 = axisOf("p1");
    expect(p0.length).toBeGreaterThan(0);
    expect(p1).toEqual(p0); // panels used to rank by their own series and disagree
  });
});
