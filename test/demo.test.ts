import { describe, expect, it } from "vitest";
import { demoItems } from "../src/core/demo";
import { CHART_KINDS } from "../src/core/samples";
import { sceneToSvg } from "../src/render/svg";
import { estimateOfficeShapes } from "../src/core/scene";

describe("demo deck", () => {
  const items = demoItems();

  it("covers every chart kind plus feature and element slides", () => {
    // One slide per kind, plus the extra feature/element highlights.
    expect(items.length).toBeGreaterThan(CHART_KINDS.length);
    const titles = new Set(items.map((i) => i.title));
    expect(titles.has("Small multiples")).toBe(true);
    expect(titles.has("Agenda")).toBe(true);
    expect(titles.has("Forecast split")).toBe(true);
  });

  it("builds a non-empty, renderable scene for every slide", () => {
    for (const item of items) {
      expect(item.scene.nodes.length, `${item.title} has nodes`).toBeGreaterThan(0);
      // Every scene must render to SVG without producing NaN geometry.
      const svg = sceneToSvg(item.scene);
      expect(svg, `${item.title} renders clean`).not.toMatch(/NaN/);
    }
  });

  it("opens with a title slide and a contents/manifest table", () => {
    expect(items[0].title).toBe("Title");
    expect(items[1].title).toBe("Contents");
    // Title slide names the deck.
    expect(items[0].scene.nodes.some((n) => n.kind === "text" && /chart gallery/i.test(n.text))).toBe(true);
    // The contents table must itself stay UNDER the ~90 web shape budget, or the
    // manifest would be the first thing skipped. It lists ~35 charts in two pairs.
    expect(items[1].scene.nodes.length).toBeLessThan(90);
    const idxText = items[1].scene.nodes.filter((n) => n.kind === "text").map((n) => n.text);
    expect(idxText).toContain("Shapes");
    expect(idxText.some((t) => /Doughnut/.test(t))).toBe(true);
    // Neither structural slide is a re-editable chart.
    expect(items[0].configJson).toBeUndefined();
    expect(items[1].configJson).toBeUndefined();
  });

  it("stamps the running build onto the title slide so a test PDF is self-identifying", () => {
    const stamped = demoItems("abc1234 · 2026-07-17 20:00Z");
    const titleTexts = stamped[0].scene.nodes.filter((n) => n.kind === "text").map((n) => n.text);
    expect(titleTexts.some((t) => t.includes("abc1234 · 2026-07-17 20:00Z"))).toBe(true);
    // Default (no build passed) still renders, with a placeholder.
    expect(demoItems()[0].scene.nodes.some((n) => n.kind === "text" && /Build local build/.test(n.text))).toBe(true);
  });

  it("estimates the EXPANDED office shape count so wedge/polygon charts are budgeted honestly", () => {
    const scene = (t: string) => items.find((i) => i.title === t)!.scene;
    // The bug the self-check exposed: node count under-counts the render. A wedge
    // fans out and a polygon draws one line per edge, so these explode.
    expect(estimateOfficeShapes(scene("Violin"))).toBeGreaterThan(200); // ~250, was 10 nodes
    expect(estimateOfficeShapes(scene("Violin"))).toBeGreaterThan(scene("Violin").nodes.length * 5);
    expect(estimateOfficeShapes(scene("Sunburst"))).toBeGreaterThan(90); // now over budget → skipped
    expect(estimateOfficeShapes(scene("Pie"))).toBeGreaterThan(scene("Pie").nodes.length); // wedge fan expands
    // A plain bar chart is one shape per node — no expansion, no over-count.
    expect(estimateOfficeShapes(scene("Stacked"))).toBe(scene("Stacked").nodes.length);
  });

  it("tags real charts with re-editable config JSON, and leaves elements untagged", () => {
    const charts = items.filter((i) => i.configJson);
    // Every chart kind is a tagged, re-editable config.
    expect(charts.length).toBeGreaterThanOrEqual(CHART_KINDS.length);
    for (const c of charts) expect(() => JSON.parse(c.configJson!)).not.toThrow();
    // The static element slides (agenda, KPI, …) are not chart configs.
    expect(items.some((i) => i.title === "Agenda" && !i.configJson)).toBe(true);
  });
});
