import { describe, expect, it } from "vitest";
import { readFileSync, existsSync } from "node:fs";
import { CHART_KINDS } from "../src/core/samples";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

/**
 * The committed showcase deck must demonstrate every chart kind and every
 * signature feature. When the feature set grows, extend
 * scripts/build-showcase.mjs and run `npm run showcase` — this test (and the
 * CI staleness gate on showcase.pptx) makes forgetting impossible.
 */
describe("showcase deck coverage", () => {
  it("showcase files exist (run `npm run showcase`)", () => {
    expect(existsSync("examples/showcase.json")).toBe(true);
    expect(existsSync("examples/showcase.pptx")).toBe(true);
  });

  const configs = JSON.parse(readFileSync("examples/showcase.json", "utf8")) as ChartConfig[];
  const json = JSON.stringify(configs);

  for (const { kind } of CHART_KINDS) {
    it(`has a slide for chart kind "${kind}"`, () => {
      expect(configs.some((c) => c.kind === kind)).toBe(true);
    });
  }

  const FEATURES: [string, (c: ChartConfig) => boolean][] = [
    ["totals", (c) => !!c.decorations?.totals],
    ["cagr arrow", (c) => !!c.decorations?.cagr],
    ["difference arrow", (c) => !!c.decorations?.difference],
    ["level difference (series)", (c) => c.decorations?.difference?.series != null],
    ["value-line-anchored difference", (c) => c.decorations?.difference?.fromValueLine != null],
    ["value lines", (c) => !!c.decorations?.valueLines?.length],
    ["label content", (c) => !!c.decorations?.labelContent],
    ["horizontal orientation", (c) => !!c.horizontal],
    ["segment order", (c) => !!c.segmentOrder],
    ["category sort", (c) => !!c.categorySort],
    ["pinned scale", (c) => c.scale?.max != null || c.scale?.min != null],
    ["axis break", (c) => !!c.axisBreak],
    ["log scale", (c) => !!c.logScale],
    ["axis title", (c) => !!c.valueAxisTitle],
    ["secondary axis", (c) => !!c.secondaryAxis],
    ["number locale", (c) => !!c.numberFormat?.locale],
    ["100%= row", (c) => !!c.data.hundredPercent],
    ["X extent row (Mekko units)", (c) => !!c.data.xExtent],
    ["stack groups", (c) => c.data.series.some((s) => (s.stack ?? 0) > 0)],
    ["waterfall totals (e)", (c) => !!c.waterfall?.totalIndices?.length],
    ["stacked waterfall", (c) => c.kind === "waterfall" && c.data.series.length > 1],
    ["label offsets", (c) => !!c.labelOffsets],
    ["calendar dates", (c) => !!c.data.dates],
    ["connector lines", (c) => !!c.decorations?.connectors],
    ["per-cell highlight colors", (c) => c.data.series.some((s) => s.colors?.some(Boolean))],
    ["callouts", (c) => !!c.decorations?.callouts?.length],
    ["background bands", (c) => !!c.decorations?.bands?.length],
    ["scatter background bands", (c) => (c.kind === "scatter" || c.kind === "bubble") && !!c.decorations?.bands?.length],
    ["exploding pie slice", (c) => !!c.pie?.explode?.length],
    ["footnote / source line", (c) => !!c.footnote],
    ["100% = note", (c) => !!c.decorations?.hundredPercentNote],
    ["pattern fills", (c) => c.data.series.some((s) => !!s.pattern)],
    ["trend statistics (Trend row present)", (c) => c.data.series.some((s) => /^trend$/i.test(s.name.trim()))],
  ];
  for (const [name, test] of FEATURES) {
    it(`demonstrates ${name}`, () => {
      expect(configs.some(test)).toBe(true);
    });
  }

  const GANTT_ROWS = ["After", "Today", "Holiday", "Bracket", "X line", "Y line", "Trend", "Group"];
  for (const row of GANTT_ROWS) {
    it(`uses the "${row}" datasheet row`, () => {
      expect(json).toContain(`"${row}`);
    });
  }

  it("every showcase config builds", () => {
    for (const c of configs) {
      expect(buildChart({ ...c, width: c.width ?? 480, height: c.height ?? 300 }).nodes.length).toBeGreaterThan(3);
    }
  });
});
