import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { buildHarveyBall, buildCheckbox, buildProcessFlow, buildTableScene } from "../src/core/elements";
import { buildAgendaScene } from "../src/core/agenda";

/**
 * Visual regression via SVG snapshots: any unintended layout change to a
 * sample chart fails CI with a readable diff. Refresh intentionally with
 * `vitest -u` after reviewing the gallery.
 */
describe("SVG snapshots", () => {
  for (const { kind } of CHART_KINDS) {
    it(`sample ${kind}`, () => {
      expect(sceneToSvg(buildChart(sampleConfig(kind)))).toMatchSnapshot();
    });
  }

  it("harvey ball 75%", () => {
    expect(sceneToSvg(buildHarveyBall(0.75))).toMatchSnapshot();
  });
  it("checkbox states", () => {
    expect(sceneToSvg(buildCheckbox("yes"))).toMatchSnapshot();
    expect(sceneToSvg(buildCheckbox("partial"))).toMatchSnapshot();
  });
  it("process flow", () => {
    expect(sceneToSvg(buildProcessFlow(["Scope", "Build", "Launch"], 1))).toMatchSnapshot();
  });
  it("table", () => {
    expect(sceneToSvg(buildTableScene([["", "2024", "2025"], ["Revenue", "78", "91"]]))).toMatchSnapshot();
  });
  it("agenda", () => {
    expect(sceneToSvg(buildAgendaScene(["Intro", "Main"], { highlight: 0 }))).toMatchSnapshot();
  });
});
