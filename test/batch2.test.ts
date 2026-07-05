import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { layoutWaterfall } from "../src/core/layout/waterfall";
import { layoutPie } from "../src/core/layout/pie";
import { valueScale } from "../src/core/layout/frame";
import { formatNumber, segmentLabel } from "../src/core/format";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode, WedgeNode } from "../src/core/scene";

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}
const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

describe("stacked waterfall", () => {
  const c = cfg({
    kind: "waterfall",
    data: {
      categories: ["Base", "Growth", "End"],
      series: [
        { name: "EU", values: [50, 8, 0] },
        { name: "US", values: [30, -6, 0] },
      ],
    },
    waterfall: { totalIndices: [2] },
  });
  const { nodes, anchors } = layoutWaterfall(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("stacks per-series contributions and moves the running level by the column sum", () => {
    expect(anchors.columnValue).toEqual([80, 82, 82]); // 50+30, +8-6, total
    const segs = byName(nodes, "bar-1-s") as RectNode[];
    expect(segs).toHaveLength(2);
  });

  it("colors segments by series in stacked mode", () => {
    const [eu] = byName(nodes, "bar-1-s0") as RectNode[];
    expect(eu.fill).toBe(DEFAULT_STYLE.palette[0]);
  });
});

describe("pie & doughnut", () => {
  const c = cfg({
    kind: "pie",
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [50, 30, 20] }] },
  });

  it("slices sum to 360° in data order", () => {
    const { nodes } = layoutPie(c, DEFAULT_STYLE, DEFAULT_DECOR);
    const wedges = nodes.filter((n): n is WedgeNode => n.kind === "wedge");
    expect(wedges).toHaveLength(3);
    expect(wedges[0].endAngle - wedges[0].startAngle).toBeCloseTo(180, 5);
    expect(wedges[2].endAngle).toBeCloseTo(360, 5);
  });

  it("doughnut adds a hole with the total", () => {
    const scene = buildChart({ ...c, kind: "doughnut" });
    expect(scene.nodes.find((n) => n.name === "hole")).toBeTruthy();
    const label = scene.nodes.find((n) => n.name === "hole-label") as TextNode;
    expect(label.text).toBe("100");
  });
});

describe("combo chart", () => {
  it("draws marked line series over the columns on a shared scale", () => {
    const scene = buildChart(
      cfg({
        kind: "combo",
        data: {
          categories: ["A", "B"],
          series: [
            { name: "Cols", values: [10, 20] },
            { name: "Line", values: [40, 50], type: "line" },
          ],
        },
      }),
    );
    expect(byName(scene.nodes, "combo-line-").length).toBeGreaterThan(0);
    expect(byName(scene.nodes, "seg-0-")).toHaveLength(2);
    // Line values above stack totals still fit: markers stay inside the plot.
    const markers = byName(scene.nodes, "combo-marker-") as RectNode[];
    for (const m of markers) expect(m.y).toBeGreaterThan(0);
  });
});

describe("label content & locale", () => {
  it("builds multi-part labels", () => {
    const label = segmentLabel(["series", "value"], {
      value: 12,
      fraction: 0.5,
      series: "SMB",
      category: "2024",
      fmt: { decimals: 0 },
    });
    expect(label).toBe("SMB 12");
  });
  it("applies segment label content in stacked charts", () => {
    const scene = buildChart(
      cfg({
        data: { categories: ["A"], series: [{ name: "S1", values: [60] }, { name: "S2", values: [40] }] },
        decorations: { labelContent: ["value", "percent"] },
      }),
    );
    const label = byName(scene.nodes, "label-0-0")[0] as TextNode;
    expect(label.text).toBe("60 60%");
  });
  it("formats with a locale", () => {
    expect(formatNumber(1234.5, { decimals: 1, locale: "de-DE" })).toBe("1.234,5");
  });
});

describe("log scale", () => {
  it("uses decade ticks and log positions", () => {
    const s = valueScale({ x: 0, y: 0, w: 100, h: 100 }, 4, 1900, undefined, undefined, true);
    expect(s.ticks[0]).toBe(1);
    expect(s.ticks[s.ticks.length - 1]).toBe(10000);
    // 100 sits halfway between 1 and 10000 in log space.
    expect(s.toY(100)).toBeCloseTo(50, 1);
  });
});

describe("decoration corners", () => {
  it("computes CAGR on a single series", () => {
    const scene = buildChart(
      cfg({
        data: {
          categories: ["Y1", "Y2"],
          series: [
            { name: "A", values: [100, 121] },
            { name: "B", values: [500, 400] },
          ],
        },
        decorations: { cagr: { from: 0, to: 1, series: 0 } },
      }),
    );
    const label = scene.nodes.find((n) => n.name === "cagr-label") as TextNode;
    expect(label.text).toContain("+21.0%");
  });

  it("anchors difference arrows at a value line", () => {
    const scene = buildChart(
      cfg({
        data: { categories: ["A"], series: [{ name: "S", values: [150] }] },
        decorations: {
          valueLines: [{ mode: "value", value: 100 }],
          difference: { from: 0, to: 0, fromValueLine: 0, percent: false },
        },
      }),
    );
    const label = scene.nodes.find((n) => n.name === "diff-label") as TextNode;
    expect(label.text).toBe("+50");
  });
});
