import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { TextNode, WedgeNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Semi-circle gauge. */

describe("semi-circle gauge", () => {
  const base: ChartConfig = {
    kind: "doughnut",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "S", values: [40, 30, 20, 10] }] },
    decorations: { segmentLabels: false },
  };

  it("fills a 180° arc with a doughnut hole and a centre total", () => {
    const s = buildChart({ ...base, pie: { semi: true } });
    const wedges = s.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    expect(wedges).toHaveLength(4);
    const span = wedges.reduce((a, w) => a + (w.endAngle - w.startAngle), 0);
    expect(span).toBeCloseTo(180, 1); // half circle
    expect(wedges.every((w) => w.innerR > 0)).toBe(true); // it's a doughnut
    expect(s.nodes.some((n) => n.name === "gauge-total")).toBe(true);
  });

  it("a plain doughnut is still a full 360° ring", () => {
    const s = buildChart(base);
    const wedges = s.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    const span = wedges.reduce((a, w) => a + (w.endAngle - w.startAngle), 0);
    expect(span).toBeCloseTo(360, 1);
  });
});

/**
 * The gauge reserved `fs * 3` a side for its outer labels — 30pt at the default
 * font, against an "Others 12%" that is 58 — so the labels of the SHIPPED
 * showcase gauge ran 32pt past the right edge of the slide.
 */
describe("a semi-circle gauge reserves room for the labels it draws", () => {
  const gauge = (categories: string[], values: number[], width = 480): ChartConfig =>
    ({
      kind: "doughnut",
      width,
      height: 300,
      pie: { semi: true },
      decorations: { segmentLabels: true },
      data: { categories, series: [{ name: "Share", values }] },
    }) as unknown as ChartConfig;

  it("keeps every outer label on the canvas", () => {
    const s = buildChart(gauge(["Us", "Rival A", "Rival B", "Others"], [42, 28, 18, 12]));
    const labels = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("label-"));
    expect(labels.length, "no labels drawn, so this proves nothing").toBe(4);
    for (const l of labels) {
      expect(l.x, `${l.text} off the left`).toBeGreaterThanOrEqual(-0.01);
      expect(l.x + l.w, `${l.text} off the right`).toBeLessThanOrEqual(480.01);
    }
  });

  it("reserves nothing extra when it is not drawing labels at all", () => {
    // The negative control: the margin is measured from the labels, so a gauge
    // with none must keep the original `fs * 3` reservation exactly.
    const bare = buildChart({
      ...gauge(["Us", "Rival A", "Rival B", "Others"], [42, 28, 18, 12]),
      decorations: { segmentLabels: false },
    } as ChartConfig);
    const arc = bare.nodes.find((n) => n.name === "slice-0") as WedgeNode;
    expect(arc.r).toBeCloseTo(480 / 2 - 30, 5);
  });

  it("grows the margin only as far as the widest label needs", () => {
    const wide = buildChart(gauge(["A category with a very long name indeed"], [100]));
    const narrow = buildChart(gauge(["A"], [100]));
    const rOf = (s: ReturnType<typeof buildChart>) => (s.nodes.find((n) => n.name === "slice-0") as WedgeNode).r;
    expect(rOf(wide)).toBeLessThan(rOf(narrow));
  });
});
