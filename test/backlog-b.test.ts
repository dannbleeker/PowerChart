import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, LineNode, RectNode, TextNode } from "../src/core/scene";

/** Backlog batch B: funnel kind + lollipop/dot/range bar styles. */

describe("funnel", () => {
  const s = buildChart(sampleConfig("funnel"));

  it("draws centered bands with width proportional to value", () => {
    const bands = [0, 1, 4].map((c) => s.nodes.find((n) => n.name === `stage-${c}`) as RectNode);
    // Widths proportional: 720/1200, 120/1200.
    expect(bands[1].w / bands[0].w).toBeCloseTo(720 / 1200, 2);
    expect(bands[2].w / bands[0].w).toBeCloseTo(120 / 1200, 2);
    // Centered: all bands share the same center x.
    const cx = (r: RectNode) => r.x + r.w / 2;
    expect(cx(bands[1])).toBeCloseTo(cx(bands[0]), 5);
    expect(cx(bands[2])).toBeCloseTo(cx(bands[0]), 5);
  });

  it("labels conversion vs the previous stage between bands", () => {
    const conv = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("conversion-"));
    expect(conv).toHaveLength(4);
    expect(conv[0].text).toContain("60.0%"); // 720/1200
    // Stage names on the left, values on/beside the bands.
    expect(s.nodes.some((n) => n.name === "category-0")).toBe(true);
    expect((s.nodes.find((n) => n.name === "stage-value-4") as TextNode).text).toBe("120");
  });

  it("narrow bands put their value beside the band, wide ones inside", () => {
    const wide = s.nodes.find((n) => n.name === "stage-value-0") as TextNode;
    const band0 = s.nodes.find((n) => n.name === "stage-0") as RectNode;
    expect(wide.x).toBeGreaterThanOrEqual(band0.x); // inside
    const tiny = buildChart({
      ...sampleConfig("funnel"),
      data: { categories: ["All", "Won"], series: [{ name: "Deals", values: [10000, 12] }] },
    });
    const narrow = tiny.nodes.find((n) => n.name === "stage-value-1") as TextNode;
    const band = tiny.nodes.find((n) => n.name === "stage-1") as RectNode;
    expect(narrow.x).toBeGreaterThan(band.x + band.w); // beside
  });
});

describe("bar styles on clustered", () => {
  const base: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "2024", values: [40, 55, 48] },
        { name: "2025", values: [52, 60, 45] },
      ],
    },
  };
  const styled = (barStyle: "lollipop" | "dot" | "range") =>
    buildChart({ ...base, decorations: { barStyle, segmentLabels: true } });

  it("lollipop: stem from the baseline + dot at the value", () => {
    const s = styled("lollipop");
    const stems = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("stem-"));
    const dots = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("seg-"));
    expect(stems).toHaveLength(6);
    expect(dots).toHaveLength(6);
    // Stem ends at the dot; no rectangles drawn for the data.
    expect(stems[0].y2).toBeCloseTo(dots[0].cy, 5);
    expect(s.nodes.some((n) => n.kind === "rect" && n.name?.startsWith("seg-"))).toBe(false);
  });

  it("dot: dots only, no stems", () => {
    const s = styled("dot");
    expect(s.nodes.some((n) => n.name?.startsWith("stem-"))).toBe(false);
    expect(s.nodes.filter((n) => n.kind === "ellipse" && n.name?.startsWith("seg-"))).toHaveLength(6);
  });

  it("range: two series' dots joined by a connector on one shared x", () => {
    const s = styled("range");
    const ranges = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("range-"));
    expect(ranges).toHaveLength(3);
    const d0 = s.nodes.find((n) => n.name === "seg-0-0") as EllipseNode;
    const d1 = s.nodes.find((n) => n.name === "seg-1-0") as EllipseNode;
    expect(d0.cx).toBeCloseTo(d1.cx, 5); // dumbbell: same x per category
    expect(ranges[0].y1).toBeCloseTo(d0.cy, 5);
    expect(ranges[0].y2).toBeCloseTo(d1.cy, 5);
  });

  it("stays plain bars by default and on stacked charts", () => {
    expect(buildChart(base).nodes.some((n) => n.kind === "rect" && n.name === "seg-0-0")).toBe(true);
    const stacked = buildChart({
      ...base,
      kind: "stacked",
      decorations: { barStyle: "lollipop", segmentLabels: true },
    });
    expect(stacked.nodes.some((n) => n.name?.startsWith("stem-"))).toBe(false);
  });
});
