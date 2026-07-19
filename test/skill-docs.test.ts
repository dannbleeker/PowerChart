import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { CHART_KINDS } from "../src/core/samples";

/**
 * Keeps the Agent Skill honest: whenever the feature set grows, the skill's
 * reference must grow with it — otherwise CI fails and the zip can't ship
 * stale. (The skill zip itself is rebuilt from these sources by CI on every
 * push, so passing docs ⇒ current skill.)
 */
describe("skill documentation coverage", () => {
  const reference = readFileSync("skill/reference.md", "utf8");
  const skillMd = readFileSync("skill/SKILL.md", "utf8");

  for (const { kind } of CHART_KINDS) {
    it(`reference.md documents chart kind "${kind}"`, () => {
      expect(reference).toContain(`"${kind}"`);
    });
  }

  it("SKILL.md names every chart kind at least once", () => {
    for (const { kind } of CHART_KINDS) {
      expect(skillMd.includes(kind)).toBe(true);
    }
  });

  it("reference.md documents the combo marker series", () => {
    // The bare word "marker" is already all over this file (forecast markers,
    // milestone markers, the boxplot mean marker), so assert the quoted
    // Series.type value — the thing that actually has to be documented.
    expect(reference).toContain('"marker"');
  });

  it("reference.md documents the special datasheet rows", () => {
    for (const row of [
      "100%=",
      "X extent",
      "Start",
      "After",
      "Today",
      "Holiday",
      "Bracket",
      "Trend",
      "X line",
      "Color",
      "Min",
      "Q1",
      "Median",
      "Q3",
      "Max",
      "Mean",
      "Outlier",
      "Error",
      "Target",
      "% Complete",
      "Baseline start",
      "Band low",
      "Band high",
      "Column <label>",
    ]) {
      expect(reference).toContain(row);
    }
  });

  it("reference.md documents the decoration keys", () => {
    for (const key of [
      "cagr",
      "difference",
      "valueLines",
      "labelContent",
      "segmentOrder",
      "axisBreak",
      "categorySort",
      "secondaryAxis",
      "labelOffsets",
      "connectors",
      "callouts",
      "bands",
      "hundredPercentNote",
      "footnote",
      "explode",
      "colors",
      "pattern",
      "scenario",
      "datamarks",
      "tickMode",
      "gridShape",
      "fillOpacity",
      "whiskers",
      "quartileMethod",
      "iqrMultiplier",
      "jitter",
      "notch",
      "map",
      "negativeColor",
      "forecastFrom",
      "quadrants",
      "columns",
      "barStyle",
      "fillBetween",
      "totals",
      "grandTotal",
      "variance",
      "slope",
      "breakout",
      "multiples",
      "stepped",
      "gapWidth",
      "overlap",
      "trajectory",
      "smooth",
      "spacerIndices",
      "summaryBars",
      "radarBand",
      "otherBucket",
      "calendar",
      "bridgeGaps",
      "perSpoke",
      "lineAxes",
      "tilemap",
      "glyph",
      "semi",
      "pareto",
      "bump",
      "criticalPath",
      "meanSd",
      "sparkline",
      "workdays",
      "marginals",
      "spreadLimit",
      "detailGroups",
      "lanes",
      "markers",
      "symbols",
      "variableRadius",
      "bars",
      "stacked",
      "sizeEncode",
      "cluster",
    ]) {
      expect(reference).toContain(key);
    }
  });
});
