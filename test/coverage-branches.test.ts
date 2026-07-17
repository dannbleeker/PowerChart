import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { niceTicks, segmentLabel, resolveFormat } from "../src/core/format";
import { sceneToSvg } from "../src/render/svg";
import type { SymbolNode } from "../src/core/scene";
import { sheetToData } from "../src/taskpane/datasheet";
import type { ChartConfig, ChartData } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

const cfg = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
  ...DEFAULT_SIZE,
  ...partial,
});

const texts = (c: ChartConfig) =>
  buildChart(c)
    .nodes.filter((n): n is TextNode => n.kind === "text")
    .map((t) => t.text);

describe("niceTicks edge cases", () => {
  it("degenerate ranges still produce a usable axis", () => {
    expect(niceTicks(5, 5).length).toBeGreaterThan(1);
    expect(niceTicks(0, 0)).toEqual([0, 1]);
  });
  it("picks 1/2/5/10 steps across magnitudes", () => {
    expect(niceTicks(0, 10, 5)).toEqual([0, 5, 10]);
    expect(niceTicks(0, 7, 5)).toEqual([0, 2, 4, 6, 8]);
    expect(niceTicks(0, 3, 5)).toEqual([0, 1, 2, 3]);
    expect(niceTicks(0, 0.4, 5)).toContain(0.1);
  });
});

describe("segmentLabel percent without a fraction", () => {
  it("drops the percent part when no denominator exists", () => {
    const label = segmentLabel(["value", "percent"], {
      value: 5,
      fraction: null,
      series: "S",
      category: "A",
      fmt: resolveFormat([5]),
    });
    expect(label).toBe("5.0");
  });
});

describe("pie fallbacks", () => {
  it("tolerates missing values and empty series names", () => {
    const s = buildChart(
      cfg({
        kind: "pie",
        data: { categories: ["A", "B", "C"], series: [{ name: "", values: [3, 2] }] },
        decorations: { segmentLabels: true },
      }),
    );
    expect(s.nodes.some((n) => n.kind === "wedge")).toBe(true);
  });
});

describe("scatter edge cases", () => {
  const scatterData = (over: Partial<ChartData>): ChartData => ({
    categories: ["P1", "P2"],
    series: [
      { name: "X", values: [1, 4] },
      { name: "Y", values: [2, 8] },
    ],
    ...over,
  });

  it("a single point renders without dividing by zero", () => {
    const s = buildChart(
      cfg({ kind: "scatter", data: { categories: ["P"], series: [{ name: "X", values: [3] }, { name: "Y", values: [3] }] } }),
    );
    expect(s.nodes.some((n) => n.name?.startsWith("point"))).toBe(true);
  });

  it("uses a custom palette for group colors", () => {
    const s = buildChart(
      cfg({
        kind: "scatter",
        style: { palette: ["#111111", "#222222"] },
        data: scatterData({ series: [...scatterData({}).series, { name: "Group", values: [1, 2] }] }),
      }),
    );
    const chips = s.nodes.filter((n) => n.name?.startsWith("legend-chip"));
    expect(chips.length).toBe(2);
    expect(chips.some((c) => c.kind === "rect" && c.fill === "#111111")).toBe(true);
  });

  it("labelContent controls point labels", () => {
    const c = cfg({
      kind: "scatter",
      data: scatterData({}),
      decorations: { segmentLabels: true, labelContent: ["category", "value"] },
    });
    const s = buildChart(c);
    const labels = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("label-"));
    expect(labels.map((l) => l.text)).toContain("P1 (1.0, 2.0)");
  });
});

describe("gantt edge cases", () => {
  const gantt = (extraRows: { name: string; values: (number | null)[] }[], startDay = 20500, span = 30): ChartConfig =>
    cfg({
      kind: "gantt",
      data: {
        categories: ["Design", "Build"],
        dates: true,
        series: [
          { name: "Start", values: [startDay, startDay + span / 3] },
          { name: "End", values: [startDay + span / 2, startDay + span] },
          ...extraRows,
        ],
      },
    });

  it("uses quarter labels for long timelines", () => {
    const s = buildChart(gantt([], 20500, 700));
    const ticks = s.nodes.filter((n): n is TextNode => n.kind === "text" && n.name === "timeline");
    expect(ticks.some((t) => /^Q[1-4] \d\d$/.test(t.text))).toBe(true);
  });

  it("labels an unlabeled bracket with its date span", () => {
    const s = buildChart(gantt([{ name: "Bracket", values: [20500, 20515] }]));
    expect(s.nodes.some((n) => n.name === "bracket-label-0")).toBe(true);
  });

  it("ignores invalid dependency references", () => {
    // Self-reference, out-of-range, and missing predecessors must not draw elbows.
    const s = buildChart(gantt([{ name: "After", values: [1, 99] }]));
    expect(s.nodes.some((n) => n.name?.startsWith("dep-"))).toBe(false);
  });

  it("draws a dependency elbow for a valid predecessor", () => {
    const s = buildChart(gantt([{ name: "After", values: [null, 1] }]));
    expect(s.nodes.some((n) => n.name?.startsWith("dep"))).toBe(true);
  });
});

describe("horizontal-frame branches", () => {
  it("reserves a legend row on horizontal charts with series labels", () => {
    const c = cfg({
      horizontal: true,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [1, 2] },
          { name: "S2", values: [3, 4] },
        ],
      },
      decorations: { seriesLabels: true, segmentLabels: true },
    });
    expect(texts(c)).toContain("S1");
  });
});

describe("sheetToData parsing fallbacks", () => {
  it("parses calendar dates and rejects garbage text", () => {
    const data = sheetToData({
      cells: [
        ["", "T1", "T2"],
        ["Start", "2026-01-15", "not-a-date"],
      ],
    });
    expect(data.dates).toBe(true);
    expect(data.series[0].values[0]).toBeGreaterThan(20000);
    expect(data.series[0].values[1]).toBeNull();
  });

  it("evaluates formulas and survives bad references", () => {
    const data = sheetToData({
      cells: [
        ["", "A", "B", "C"],
        ["S", "2", "=B2*2", "=SUM(B2:C2)"],
      ],
    });
    expect(data.series[0].values).toEqual([2, 4, 6]);
    // Out-of-range references read as 0; unparseable formulas become null.
    expect(sheetToData({ cells: [["", "A"], ["S", "=ZZ99"]] }).series[0].values).toEqual([0]);
    expect(sheetToData({ cells: [["", "A"], ["S", "=1/"]] }).series[0].values).toEqual([null]);
  });
});

describe("SVG annular wedge path", () => {
  it("emits inner and outer arcs for wedges with a hole", () => {
    const svg = sceneToSvg({
      width: 100,
      height: 100,
      nodes: [
        { kind: "wedge", cx: 50, cy: 50, r: 40, innerR: 20, startAngle: 0, endAngle: 120, fill: "#123456", stroke: "#000000", strokeWidth: 1 },
      ],
    });
    expect(svg.match(/A /g)!.length).toBe(2);
    expect(svg).toContain('stroke="#000000"');
  });
});

describe("SVG marker symbols", () => {
  const svgOf = (n: Partial<SymbolNode> = {}) =>
    sceneToSvg({
      width: 100,
      height: 100,
      nodes: [{ kind: "symbol", shape: "diamond", cx: 50, cy: 50, size: 10, fill: "#123456", ...n } as SymbolNode],
    });

  it("draws a filled polygon on the symbol's own points", () => {
    const svg = svgOf();
    expect(svg).toContain('fill="#123456"');
    // The four diamond vertices, at the box edge midpoints.
    expect(svg).toContain('points="50,40 60,50 50,60 40,50"');
  });

  it("carries an optional stroke, and omits it when absent", () => {
    expect(svgOf({ stroke: "#ffffff", strokeWidth: 2 })).toContain('stroke="#ffffff" stroke-width="2"');
    expect(svgOf()).not.toContain("stroke=");
  });

  it("emits the data-name so a symbol is addressable like any other node", () => {
    expect(svgOf({ name: "point-3" })).toContain('data-name="point-3"');
  });

  it("draws every shape, and rounds coordinates like the rest of the renderer", () => {
    for (const shape of ["diamond", "triangle", "plus", "star5"] as const) {
      const svg = svgOf({ shape, cx: 33.333333, cy: 12.126, size: 7.77 });
      expect(svg).toContain("<polygon");
      // r() quantises to 2dp; a raw float here would be snapshot noise.
      const pts = svg.match(/points="([^"]+)"/)![1];
      for (const n of pts.split(/[ ,]/)) expect(n).toMatch(/^-?\d+(\.\d{1,2})?$/);
    }
  });
});
