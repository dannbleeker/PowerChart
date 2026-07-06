import { describe, expect, it } from "vitest";
import { demoItems } from "../src/core/demo";
import { CHART_KINDS } from "../src/core/samples";
import { sceneToSvg } from "../src/render/svg";

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

  it("tags real charts with re-editable config JSON, and leaves elements untagged", () => {
    const charts = items.filter((i) => i.configJson);
    // Every chart kind is a tagged, re-editable config.
    expect(charts.length).toBeGreaterThanOrEqual(CHART_KINDS.length);
    for (const c of charts) expect(() => JSON.parse(c.configJson!)).not.toThrow();
    // The static element slides (agenda, KPI, …) are not chart configs.
    expect(items.some((i) => i.title === "Agenda" && !i.configJson)).toBe(true);
  });
});
