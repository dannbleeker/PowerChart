import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { buildKpiTile } from "../src/core/elements";
import { sampleConfig } from "../src/core/samples";
import { NO_DATA } from "../src/core/color";
import { PALETTE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { ArrowheadNode, LineNode, RectNode, TextNode } from "../src/core/scene";

/** Backlog batch E: slope chart option, waffle kind, KPI tile element. */

describe("slope chart (decorations.slope on line)", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: {
      categories: ["2020", "2025"],
      series: [
        { name: "Brand A", values: [48, 66] },
        { name: "Brand B", values: [40, 49] },
      ],
    },
    decorations: { slope: true },
  };
  const s = buildChart(cfg);

  it("draws end rails, straight lines, and 'Name value' labels at both ends", () => {
    expect(s.nodes.filter((n) => n.name?.startsWith("slope-rail"))).toHaveLength(2);
    expect((s.nodes.find((n) => n.name === "slope-left-0") as TextNode).text).toBe("Brand A 48");
    expect((s.nodes.find((n) => n.name === "slope-right-0") as TextNode).text).toBe("Brand A 66");
    const line = s.nodes.find((n) => n.name === "line-0-1") as LineNode;
    expect(line.y1).toBeGreaterThan(line.y2); // 48 → 66 rises
    // No value axis or gridlines in slope mode.
    expect(s.nodes.some((n) => n.name?.startsWith("grid"))).toBe(false);
    expect(s.nodes.some((n) => n.name?.startsWith("vaxis"))).toBe(false);
  });

  it("labels take the series color and de-overlap when ends are close", () => {
    const crowded = buildChart({
      ...cfg,
      data: {
        categories: ["2020", "2025"],
        series: [
          { name: "Alpha", values: [50, 60] },
          { name: "Beta", values: [50.5, 60.5] },
          { name: "Gamma", values: [51, 61] },
        ],
      },
    });
    const ys = [0, 1, 2]
      .map((si) => (crowded.nodes.find((n) => n.name === `slope-left-${si}`) as TextNode).y)
      .sort((a, b) => a - b);
    expect(ys[1] - ys[0]).toBeGreaterThanOrEqual(14);
    expect(ys[2] - ys[1]).toBeGreaterThanOrEqual(14);
    const label = s.nodes.find((n) => n.name === "slope-left-0") as TextNode;
    const line = s.nodes.find((n) => n.name === "line-0-1") as LineNode;
    expect(label.color).toBe(line.stroke);
  });

  it("handles >2 categories as a polyline and ignores slope on area charts", () => {
    const three = buildChart({
      ...cfg,
      data: {
        categories: ["2020", "2022", "2025"],
        series: [{ name: "A", values: [40, 55, 66] }],
      },
    });
    expect(three.nodes.filter((n) => n.name?.startsWith("line-0-"))).toHaveLength(2);
    expect(three.nodes.filter((n) => n.name?.startsWith("slope-rail"))).toHaveLength(2); // ends only
    const area = buildChart({ ...cfg, kind: "area" });
    expect(area.nodes.some((n) => n.name?.startsWith("slope-rail"))).toBe(false);
  });
});

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

describe("KPI tile element", () => {
  it("shows caption, big value, and a colored delta with arrow", () => {
    const s = buildKpiTile({ label: "Revenue", value: "€4.2m", delta: "+12% vs LY" });
    expect((s.nodes.find((n) => n.name === "kpi-value") as TextNode).text).toBe("€4.2m");
    expect((s.nodes.find((n) => n.name === "kpi-label") as TextNode).text).toBe("Revenue");
    const arrow = s.nodes.find((n) => n.name === "kpi-arrow") as ArrowheadNode;
    expect(arrow.angle).toBe(-90); // up
    expect(arrow.fill).toBe("#0ca30c"); // up is good by default
    expect((s.nodes.find((n) => n.name === "kpi-delta") as TextNode).color).toBe("#0ca30c");
  });

  it("goodIsUp:false colors a falling metric green", () => {
    const s = buildKpiTile({ value: "2.1%", delta: "-0.4pp churn", goodIsUp: false });
    const arrow = s.nodes.find((n) => n.name === "kpi-arrow") as ArrowheadNode;
    expect(arrow.angle).toBe(90); // down
    expect(arrow.fill).toBe("#0ca30c"); // …and that's good
    const up = buildKpiTile({ value: "2.1%", delta: "+0.4pp churn", goodIsUp: false });
    expect((up.nodes.find((n) => n.name === "kpi-arrow") as ArrowheadNode).fill).toBe("#d03b3b");
  });

  it("flat or missing deltas stay neutral, long values shrink to fit", () => {
    const flat = buildKpiTile({ value: "87", delta: "unchanged", direction: "flat" });
    expect(flat.nodes.some((n) => n.name === "kpi-arrow")).toBe(false);
    expect((flat.nodes.find((n) => n.name === "kpi-delta") as TextNode).color).not.toBe("#0ca30c");
    const none = buildKpiTile({ value: "87" });
    expect(none.nodes.some((n) => n.name === "kpi-delta")).toBe(false);
    const long = buildKpiTile({ value: "€1,234,567.89 total" });
    const short = buildKpiTile({ value: "€4m" });
    const fsOf = (s: ReturnType<typeof buildKpiTile>) => (s.nodes.find((n) => n.name === "kpi-value") as TextNode).fontSize;
    expect(fsOf(long)).toBeLessThan(fsOf(short));
  });
});
