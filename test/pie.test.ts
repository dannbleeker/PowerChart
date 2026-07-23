import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { layoutPie } from "../src/core/layout/pie";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "../src/core/style";
import { sceneToSvg } from "../src/render/svg";
import type { EllipseNode, LineNode, RectNode, TextNode, WedgeNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Pie / doughnut — breakouts, all-zero totals, narrow-frame radius, full-circle slice. */

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

/** Backlog batch F: pie-of-pie breakout, small multiples. */
describe("pie breakout (bar-of-pie)", () => {
  const cfg: ChartConfig = {
    kind: "pie",
    ...DEFAULT_SIZE,
    data: {
      categories: ["EMEA", "Americas", "APAC", "Nordics", "Benelux", "DACH"],
      series: [{ name: "Revenue", values: [80, 100, 60, 20, 15, 25] }],
    },
    pie: { breakout: [3, 4, 5] },
  };
  const s = buildChart(cfg);

  it("collapses breakout categories into one muted Other slice facing the bar", () => {
    expect(s.nodes.some((n) => n.name === "slice-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "slice-3")).toBe(false);
    const other = s.nodes.find((n) => n.name === "slice-other") as WedgeNode;
    expect(other.fill).toBe("#898781");
    // Other (60 of 300 = 72°) is centered at 3 o'clock: spans 90° ± 36°.
    const mid = ((other.startAngle + other.endAngle) / 2) % 360;
    expect(mid).toBeCloseTo(90, 1);
  });

  it("details the breakout in a stacked bar with connectors and grand-total %", () => {
    const segs = [3, 4, 5].map((c) => s.nodes.find((n) => n.name === `breakout-seg-${c}`) as RectNode);
    expect(segs.every(Boolean)).toBe(true);
    // Stacked contiguously, heights ∝ values (20/15/25 of 60).
    expect(segs[0].y + segs[0].h).toBeCloseTo(segs[1].y, 5);
    expect(segs[1].y + segs[1].h).toBeCloseTo(segs[2].y, 5);
    expect(segs[2].h / segs[0].h).toBeCloseTo(25 / 20, 5);
    // Bar sits right of the pie.
    const other = s.nodes.find((n) => n.name === "slice-other") as WedgeNode;
    expect(segs[0].x).toBeGreaterThan(other.cx);
    // Labels carry the share of the GRAND total (20/300 ≈ 7%).
    expect((s.nodes.find((n) => n.name === "breakout-label-3") as TextNode).text).toContain("7%");
    const conns = s.nodes.filter((n): n is LineNode => !!n.name?.startsWith("breakout-conn"));
    expect(conns).toHaveLength(2);
    // Connectors join the bar's top and bottom.
    const ends = conns.map((c) => c.y2).sort((a, b) => a - b);
    expect(ends[0]).toBeCloseTo(segs[0].y, 5);
    expect(ends[1]).toBeCloseTo(segs[2].y + segs[2].h, 5);
  });

  it("plain pies and doughnuts are unaffected", () => {
    const plain = buildChart({ ...cfg, pie: {} });
    expect(plain.nodes.some((n) => n.name === "slice-3")).toBe(true);
    expect(plain.nodes.some((n) => n.name === "slice-other")).toBe(false);
    const dough = buildChart({ ...cfg, kind: "doughnut" });
    expect(dough.nodes.some((n) => n.name === "slice-other")).toBe(false);
  });
});

describe("variable-radius pie", () => {
  const cfg: ChartConfig = {
    kind: "pie",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "Share", values: [50, 30, 20] }, // angle
        { name: "Radius", values: [20, 80, 50] }, // radius metric
      ],
    },
    pie: { variableRadius: true },
  };
  const s = buildChart(cfg);
  const slice = (c: number) => s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === `slice-${c}`)!;

  it("keeps angle ∝ the first series but sets radius from the Radius row", () => {
    // Angle still encodes Share: A (50) sweeps the widest.
    const spanA = slice(0).endAngle - slice(0).startAngle;
    const spanB = slice(1).endAngle - slice(1).startAngle;
    expect(spanA).toBeGreaterThan(spanB);
    // Radius encodes the Radius row: B (80) longest, A (20) shortest — even
    // though A has the biggest share.
    expect(slice(1).r).toBeGreaterThan(slice(2).r);
    expect(slice(2).r).toBeGreaterThan(slice(0).r);
  });

  it("is inert without a Radius row or the flag", () => {
    const plain = buildChart({
      ...cfg,
      pie: {},
      data: { categories: ["A", "B"], series: [{ name: "Share", values: [50, 50] }] },
    });
    const a = plain.nodes.find((n): n is WedgeNode => n.name === "slice-0")!;
    const b = plain.nodes.find((n): n is WedgeNode => n.name === "slice-1")!;
    expect(a.r).toBe(b.r); // uniform radius
  });
});

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

describe("pie radius never goes negative on a narrow frame", () => {
  it("floors the radius so wedges and the hole keep valid geometry", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 80, // width*0.5 - fs*7 = -30 before the floor
      height: 320,
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [3, 4, 5] }] },
    };
    const nodes = buildChart(cfg).nodes;
    for (const n of nodes) {
      if (n.kind === "wedge") expect((n as WedgeNode).r).toBeGreaterThanOrEqual(0);
      if (n.kind === "ellipse") {
        expect((n as EllipseNode).rx).toBeGreaterThanOrEqual(0);
        expect((n as EllipseNode).ry).toBeGreaterThanOrEqual(0);
      }
    }
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

describe("pie / doughnut all-zero total", () => {
  it("shows the true total (0), not the divisor fallback of 1", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 400,
      height: 300,
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [0, 0, 0] }] },
    };
    const nodes = buildChart(cfg).nodes;
    const texts = nodes.filter((n): n is TextNode => n.kind === "text");
    expect(nodes.some((n) => n.name === "hole")).toBe(true);
    // The centre shows the honest total (0, at the data's 2-decimal precision),
    // never the divisor fallback the old `|| 1` displayed as "1".
    expect(texts.some((t) => /^0(\.0+)?$/.test(t.text))).toBe(true);
    expect(texts.some((t) => /^1(\.0+)?$/.test(t.text))).toBe(false);
  });

  it("still renders a normal doughnut total unchanged", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 400,
      height: 300,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [30, 70] }] },
    };
    const texts = buildChart(cfg).nodes.filter((n): n is TextNode => n.kind === "text");
    expect(texts.some((t) => t.text === "100")).toBe(true);
  });
});
