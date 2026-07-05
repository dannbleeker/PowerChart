import { describe, expect, it } from "vitest";
import { buildCheckbox, buildHarveyBall, buildProcessFlow, buildTableScene } from "../src/core/elements";
import { layoutGantt } from "../src/core/layout/gantt";
import { layoutPie } from "../src/core/layout/pie";
import { buildChart } from "../src/core/chart";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import type { ChartConfig } from "../src/core/types";
import type { TextNode, WedgeNode } from "../src/core/scene";

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}
const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

describe("elements", () => {
  it("harvey ball wedge matches the fraction", () => {
    const scene = buildHarveyBall(0.25);
    const wedge = scene.nodes.find((n): n is WedgeNode => n.kind === "wedge")!;
    expect(wedge.endAngle).toBeCloseTo(90, 5);
    expect(buildHarveyBall(1).nodes.filter((n) => n.kind === "ellipse")).toHaveLength(2);
    expect(buildHarveyBall(0).nodes.some((n) => n.kind === "wedge")).toBe(false);
  });

  it("checkbox states use distinct glyphs", () => {
    const glyphs = (["yes", "no", "partial"] as const).map(
      (s) => (buildCheckbox(s).nodes.find((n) => n.kind === "text") as TextNode).text,
    );
    expect(new Set(glyphs).size).toBe(3);
  });

  it("process flow highlights the active chevron", () => {
    const scene = buildProcessFlow(["A", "B", "C"], 1, 300, 40);
    const chevrons = scene.nodes.filter((n) => n.kind === "chevron" && n.name?.startsWith("step-"));
    expect(chevrons).toHaveLength(3);
    const fills = chevrons.map((c) => (c as { fill: string }).fill);
    expect(fills[1]).not.toBe(fills[0]);
    expect((chevrons[0] as { flatLeft?: boolean }).flatLeft).toBe(true);
  });

  it("table sizes columns to content and styles the header", () => {
    const scene = buildTableScene([["", "Long header col", "B"], ["Row", "1", "2"]], 400);
    const cells = byName(scene.nodes, "cell-0-");
    expect(cells).toHaveLength(3);
    const w1 = (scene.nodes.find((n) => n.name === "cell-0-1") as { w: number }).w;
    const w2 = (scene.nodes.find((n) => n.name === "cell-0-2") as { w: number }).w;
    expect(w1).toBeGreaterThan(w2);
  });
});

describe("gantt depth", () => {
  const c = cfg({
    kind: "gantt",
    data: {
      categories: ["Design | Anna", "Build | Ben"],
      series: [
        { name: "Start", values: [0, 5] },
        { name: "End", values: [5, 10] },
        { name: "After", values: [null, 1] },
        { name: "Today", values: [7, null] },
      ],
    },
  });
  const { nodes } = layoutGantt(c, DEFAULT_STYLE, DEFAULT_DECOR);

  it("splits owners into a responsible column", () => {
    const owners = byName(nodes, "owner-") as TextNode[];
    expect(owners.map((o) => o.text)).toEqual(["Anna", "Ben"]);
    const cat = byName(nodes, "category-0")[0] as TextNode;
    expect(cat.text).toBe("Design");
  });

  it("draws dependency elbows with arrowheads", () => {
    expect(nodes.find((n) => n.name === "dep-v-1")).toBeTruthy();
    expect(nodes.find((n) => n.name === "dep-head-1")).toBeTruthy();
  });

  it("draws the today line from the Today row", () => {
    expect(nodes.find((n) => n.name === "today-line")).toBeTruthy();
  });
});

describe("smarter labels", () => {
  it("adds leader lines to outside pie labels", () => {
    const c = cfg({
      kind: "pie",
      data: { categories: ["Big", "Tiny"], series: [{ name: "S", values: [95, 5] }] },
    });
    const { nodes } = layoutPie(c, DEFAULT_STYLE, DEFAULT_DECOR);
    expect(nodes.find((n) => n.name === "leader-1")).toBeTruthy(); // tiny slice → outside
    expect(nodes.find((n) => n.name === "leader-0")).toBeFalsy(); // big slice → inside
  });

  it("applies manual label offsets by node name", () => {
    const base = cfg({
      data: { categories: ["A"], series: [{ name: "S", values: [50] }] },
    });
    const before = buildChart(base).nodes.find((n) => n.name === "label-0-0") as TextNode;
    const after = buildChart({ ...base, labelOffsets: { "label-0-0": { dx: 5, dy: -10 } } }).nodes.find(
      (n) => n.name === "label-0-0",
    ) as TextNode;
    expect(after.x - before.x).toBeCloseTo(5, 5);
    expect(after.y - before.y).toBeCloseTo(-10, 5);
  });
});
