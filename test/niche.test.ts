import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { layoutGantt } from "../src/core/layout/gantt";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}
const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

describe("category sorting", () => {
  it("sorts categories by total, descending", () => {
    const scene = buildChart(
      cfg({
        categorySort: "descending",
        data: {
          categories: ["Small", "Big", "Mid"],
          series: [{ name: "S", values: [10, 100, 50] }],
        },
        decorations: { categoryAxis: true },
      }),
    );
    const cats = byName(scene.nodes, "category-") as TextNode[];
    expect(cats.map((c) => c.text)).toEqual(["Big", "Mid", "Small"]);
  });

  it("leaves order-sensitive kinds untouched", () => {
    const scene = buildChart(
      cfg({
        kind: "waterfall",
        categorySort: "descending",
        data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 100] }] },
      }),
    );
    const cats = byName(scene.nodes, "category-") as TextNode[];
    expect(cats.map((c) => c.text)).toEqual(["A", "B"]);
  });
});

describe("combo secondary axis", () => {
  it("scales line series independently and adds right-hand ticks", () => {
    const scene = buildChart(
      cfg({
        kind: "combo",
        secondaryAxis: true,
        data: {
          categories: ["A", "B"],
          series: [
            { name: "Revenue", values: [500, 800] },
            { name: "Margin %", values: [30, 45], type: "line" },
          ],
        },
      }),
    );
    expect(byName(scene.nodes, "secondary-axis").length).toBeGreaterThanOrEqual(3);
    // With its own 0..50-ish scale, the 45% marker sits high in the plot.
    const markers = byName(scene.nodes, "combo-marker-0-1");
    expect(markers).toHaveLength(1);
  });
});

describe("gantt holidays & brackets", () => {
  const day = (iso: string) => Math.round(Date.parse(iso) / 86400000);
  const c = cfg({
    kind: "gantt",
    data: {
      categories: ["Build"],
      series: [
        { name: "Start", values: [day("2026-01-05")] },
        { name: "End", values: [day("2026-01-30")] },
        { name: "Holiday", values: [day("2026-01-15")] },
        { name: "Bracket Sprint 1", values: [day("2026-01-05"), day("2026-01-19")] },
      ],
      dates: true,
    },
  });
  const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("shades holidays", () => {
    expect(byName(nodes, "holiday-")).toHaveLength(1);
  });
  it("draws labelled bracket annotations", () => {
    expect(nodes.find((n) => n.name === "bracket-0")).toBeTruthy();
    const label = nodes.find((n) => n.name === "bracket-label-0") as TextNode;
    expect(label.text).toBe("Sprint 1");
  });
});

describe("label content everywhere", () => {
  it("line point labels honor labelContent", () => {
    const scene = buildChart(
      cfg({
        kind: "line",
        data: { categories: ["A"], series: [{ name: "GM", values: [42] }] },
        decorations: { segmentLabels: true, labelContent: ["series", "value"] },
      }),
    );
    const label = byName(scene.nodes, "label-0-0")[0] as TextNode;
    expect(label.text).toBe("GM 42");
  });
  it("scatter labels can include coordinates", () => {
    const scene = buildChart(
      cfg({
        kind: "scatter",
        data: {
          categories: ["P"],
          series: [
            { name: "X", values: [10] },
            { name: "Y", values: [20] },
          ],
        },
        decorations: { labelContent: ["category", "value"] },
      }),
    );
    const label = byName(scene.nodes, "label-0")[0] as TextNode;
    expect(label.text).toBe("P (10, 20)");
  });
});
