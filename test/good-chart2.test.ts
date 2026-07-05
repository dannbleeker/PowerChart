import { describe, expect, it } from "vitest";
import { buildTableScene } from "../src/core/elements";
import { trendStats, formatP } from "../src/core/format";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";
import type { ChartConfig } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

/** Batch 2: rule-based tables, in-cell effects, trend statistics, pattern fills. */

describe("rule-based table", () => {
  const cells = [
    ["Region", "FY23", "FY24"],
    ["North", "120", "134"],
    ["South", "98", "91"],
    ["East", "77", "80"],
    ["West", "65", "64"],
    ["Central", "51", "58"],
    ["Islands", "20", "22"],
    ["Total", "431", "449"],
  ];

  it("draws only horizontal rules — never side borders or row lines", () => {
    const s = buildTableScene(cells, 400);
    const lines = s.nodes.filter((n) => n.kind === "line");
    expect(lines.map((l) => l.name).sort()).toEqual(["rule-bottom", "rule-header", "rule-top"]);
    // All rules are horizontal and full-width.
    for (const l of lines) if (l.kind === "line") expect(l.y1).toBe(l.y2);
    expect(s.nodes.some((n) => n.kind === "rect")).toBe(false);
  });

  it("adds a separator gap after every 5 body rows", () => {
    const s = buildTableScene(cells, 400);
    const rowY = (ri: number) => (s.nodes.find((n) => n.name === `cell-text-${ri}-0`) as TextNode).y;
    const gapless = rowY(2) - rowY(1);
    // Row 6 is the first after the 5-row group → extra gap.
    expect(rowY(6) - rowY(5)).toBeGreaterThan(gapless + 1);
    expect(rowY(3) - rowY(2)).toBeCloseTo(gapless);
  });

  it("renders a bold totals row with a rule above it", () => {
    const s = buildTableScene(cells, 400, { totalRow: true });
    expect(s.nodes.some((n) => n.name === "rule-total")).toBe(true);
    const total = s.nodes.find((n) => n.name === "cell-text-7-0") as TextNode;
    expect(total.bold).toBe(true);
  });

  it("parses in-cell effect tokens into glyphs and colors", () => {
    const s = buildTableScene(
      [
        ["", "A"],
        ["Full", "[hb:1] done"],
        ["Half", "[hb:50%] going"],
        ["Up", "[up] +14%"],
        ["Risk", "[bad] slipping"],
      ],
      300,
    );
    expect(s.nodes.some((n) => n.name === "cell-hb-1-1" && n.kind === "ellipse")).toBe(true);
    expect(s.nodes.some((n) => n.name === "cell-hb-2-1" && n.kind === "wedge")).toBe(true);
    const up = s.nodes.find((n) => n.name === "cell-trend-3-1");
    expect(up?.kind === "arrowhead" && up.angle).toBe(-90);
    const risk = s.nodes.find((n) => n.name === "cell-text-4-1") as TextNode;
    expect(risk.color).toBe("#d03b3b");
    expect(risk.text).toBe("slipping");
    // Token stripped from the visible text everywhere.
    expect((s.nodes.find((n) => n.name === "cell-text-3-1") as TextNode).text).toBe("+14%");
  });
});

describe("trend statistics", () => {
  it("computes R² = 1 and p = 0 for a perfect line", () => {
    const pts = [1, 2, 3, 4, 5].map((x) => ({ x, y: 2 * x + 1 }));
    expect(trendStats(pts)).toEqual({ r2: 1, p: 0 });
  });

  it("computes a plausible p for noisy data and null for tiny samples", () => {
    const noisy = [
      { x: 1, y: 2 }, { x: 2, y: 4.2 }, { x: 3, y: 5.4 }, { x: 4, y: 8.9 },
      { x: 5, y: 9.6 }, { x: 6, y: 12.4 }, { x: 7, y: 13.1 },
    ];
    const s = trendStats(noisy)!;
    expect(s.r2).toBeGreaterThan(0.95);
    expect(s.p).not.toBeNull();
    expect(s.p!).toBeLessThan(0.001);
    // Uncorrelated data → high p.
    const flat = trendStats([{ x: 1, y: 5 }, { x: 2, y: 1 }, { x: 3, y: 6 }, { x: 4, y: 2 }, { x: 5, y: 5 }])!;
    expect(flat.p!).toBeGreaterThan(0.3);
    // Two points: fit exists, no significance (df = 0).
    expect(trendStats([{ x: 1, y: 1 }, { x: 2, y: 3 }])!.p).toBeNull();
  });

  it("formats p against the conventional cuts", () => {
    expect(formatP(0.0004)).toBe("< 0.001");
    expect(formatP(0.004)).toBe("< 0.01");
    expect(formatP(0.03)).toBe("< 0.05");
    expect(formatP(0.31)).toBe("= 0.31");
  });

  it("labels the scatter trend line with R² and p", () => {
    const cfg: ChartConfig = {
      kind: "scatter",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A", "B", "C", "D", "E"],
        series: [
          { name: "X", values: [1, 2, 3, 4, 5] },
          { name: "Y", values: [2, 4.2, 5.4, 8.9, 9.6] },
          { name: "Trend", values: [1, null, null, null, null] },
        ],
      },
    };
    const label = buildChart(cfg).nodes.find((n) => n.name === "trend-stats") as TextNode;
    expect(label).toBeDefined();
    expect(label.text).toMatch(/^R² = 0\.9\d, p [<=]/);
  });
});

describe("pattern fills", () => {
  const cfg: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "Base", values: [10, 12] },
        { name: "Plan", values: [4, 5], pattern: "diagonal" },
      ],
    },
  };

  it("emits one SVG pattern def per (pattern, color) and references it", () => {
    const svg = sceneToSvg(buildChart(cfg));
    expect(svg.match(/<pattern /g)).toHaveLength(1);
    expect(svg).toContain('fill="url(#p-diagonal-');
    // The base series stays a plain solid fill.
    expect(svg).toContain('data-name="seg-0-0"');
  });

  it("keeps the solid fill on the scene node for non-SVG renderers", () => {
    const seg = buildChart(cfg).nodes.find((n) => n.name === "seg-1-0");
    expect(seg?.kind === "rect" && seg.fill.startsWith("#")).toBe(true);
    expect(seg?.kind === "rect" && seg.pattern).toBe("diagonal");
  });

  it("supports all four tile types", () => {
    const all: ChartConfig = {
      ...cfg,
      data: {
        categories: ["A"],
        series: (["diagonal", "crosshatch", "dots", "horizontal"] as const).map((pattern, i) => ({
          name: pattern,
          values: [5 + i],
          pattern,
        })),
      },
    };
    expect(sceneToSvg(buildChart(all)).match(/<pattern /g)).toHaveLength(4);
  });
});
