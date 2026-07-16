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
    ["boxplot raw samples", (c) => c.kind === "boxplot" && !c.data.series.some((s) => /^min$/i.test(s.name))],
    ["boxplot summary rows", (c) => c.kind === "boxplot" && c.data.series.some((s) => /^median$/i.test(s.name))],
    ["horizontal boxplot", (c) => c.kind === "boxplot" && !!c.horizontal],
    ["error bars", (c) => c.data.series.some((s) => /^error/i.test(s.name.trim()))],
    ["bullet target row", (c) => c.data.series.some((s) => /^target$/i.test(s.name.trim()))],
    ["combo column mode", (c) => !!c.combo?.columns],
    ["forecast styling", (c) => c.decorations?.forecastFrom != null],
    ["scatter quadrants", (c) => !!c.decorations?.quadrants],
    ["bar style (lollipop/dot/range)", (c) => !!c.decorations?.barStyle && c.decorations.barStyle !== "bar"],
    ["gantt progress", (c) => c.data.series.some((s) => /complete/i.test(s.name))],
    ["gantt baselines", (c) => c.data.series.some((s) => /^baseline\s*start$/i.test(s.name.trim()))],
    ["grouped boxplots", (c) => c.kind === "boxplot" && c.data.series.some((s) => s.name.includes("|"))],
    ["waterfall gap-to-target", (c) => c.kind === "waterfall" && c.data.series.some((s) => /^target$/i.test(s.name.trim()))],
    ["confidence band rows", (c) => c.data.series.some((s) => /^band\s*low$/i.test(s.name.trim()))],
    ["heatmap marginal totals", (c) => !!c.heatmap?.totals],
    ["slope chart", (c) => !!c.decorations?.slope],
    ["multi-part waffle", (c) => c.kind === "waffle" && c.data.categories.length > 1],
    ["bar-of-pie breakout", (c) => !!c.pie?.breakout?.length],
    ["small multiples", (c) => !!c.multiples],
    ["datamark axis", (c) => c.decorations?.valueAxis === "datamarks"],
    ["diverging heatmap (negative values)", (c) => c.kind === "heatmap" && c.data.series.some((s) => s.values.some((v) => (v ?? 0) < 0))],
    ["explicit map layout", (c) => !!c.map],
    ["auto-detected map layout", (c) => c.kind === "tilemap" && !c.map],
    ["trend statistics (Trend row present)", (c) => c.data.series.some((s) => /^trend$/i.test(s.name.trim()))],
    ["stepped line", (c) => c.kind === "line" && !!c.decorations?.stepped],
    ["stepped area", (c) => c.kind === "area" && !!c.decorations?.stepped],
    ["butterfly value ticks", (c) => c.kind === "butterfly" && !!c.decorations?.valueAxis],
    ["clustered overlap", (c) => c.overlap != null],
    ["column gap width", (c) => c.gapWidth != null],
    ["area with negative values", (c) => c.kind === "area" && c.data.series.some((s) => s.values.some((v) => (v ?? 0) < 0))],
    ["scatter/bubble trajectory", (c) => (c.kind === "scatter" || c.kind === "bubble") && !!c.decorations?.trajectory],
    ["boxplot jitter", (c) => c.kind === "boxplot" && !!c.boxplot?.jitter],
    ["scatter continuous color row", (c) => (c.kind === "scatter" || c.kind === "bubble") && c.data.series.some((s) => /^colou?r$/i.test(s.name.trim()))],
    ["smoothed line", (c) => c.kind === "line" && !!c.decorations?.smooth],
    ["waterfall grouping spacers", (c) => !!c.waterfall?.spacerIndices?.length],
    ["gantt summary bars", (c) => c.kind === "gantt" && !!c.decorations?.summaryBars],
    ["notched boxplots", (c) => c.kind === "boxplot" && !!c.boxplot?.notch],
    ["radar peer-range band", (c) => c.kind === "radar" && !!c.decorations?.radarBand],
    ["other bucket", (c) => !!c.otherBucket],
    ["calendar heatmap", (c) => c.kind === "heatmap" && !!c.heatmap?.calendar],
    ["butterfly stacked flanks", (c) => c.kind === "butterfly" && c.butterfly?.split != null],
    ["radar per-spoke scales", (c) => c.kind === "radar" && !!c.radar?.perSpoke],
    ["line bridge gaps", (c) => c.kind === "line" && !!c.decorations?.bridgeGaps],
    ["transparent floating segment", (c) => c.data.series.some((s) => s.color === "transparent")],
    ["combo waterfall base", (c) => c.combo?.columns === "waterfall"],
    ["combo mekko base", (c) => c.combo?.columns === "mekko"],
    ["combo independent line axes", (c) => c.combo?.lineAxes === "independent"],
    ["hex tilemap", (c) => c.kind === "tilemap" && c.tilemap?.shape === "hex"],
    ["tilemap mini-glyphs", (c) => c.kind === "tilemap" && c.tilemap?.glyph === "bars"],
    ["stacked100 with negatives", (c) => c.kind === "stacked100" && c.data.series.some((s) => s.values.some((v) => (v ?? 0) < 0))],
    ["semi-circle gauge", (c) => c.kind === "doughnut" && !!c.pie?.semi],
    ["pareto chart", (c) => !!c.pareto],
    ["bump chart", (c) => c.kind === "line" && !!c.decorations?.bump],
    ["horizontal profile (line/area)", (c) => (c.kind === "line" || c.kind === "area") && !!c.horizontal],
    ["two-level treemap", (c) => c.kind === "treemap" && c.data.categories.some((cat) => cat.includes("|"))],
    ["critical-path gantt", (c) => c.kind === "gantt" && !!c.decorations?.criticalPath],
    ["mean±SD boxplot", (c) => c.kind === "boxplot" && !!c.boxplot?.meanSd],
    ["sparklines", (c) => !!c.decorations?.sparkline],
    ["radial bar chart", (c) => c.kind === "radar" && !!c.radar?.bars],
    ["stacked radar", (c) => c.kind === "radar" && !!c.radar?.stacked],
    ["variable-radius pie", (c) => c.kind === "pie" && !!c.pie?.variableRadius],
    ["cell-size heatmap", (c) => c.kind === "heatmap" && !!c.heatmap?.sizeEncode],
    ["clustered heatmap", (c) => c.kind === "heatmap" && !!c.heatmap?.cluster],
    ["combo area base", (c) => c.combo?.columns === "area"],
    ["gantt gutter columns", (c) => c.kind === "gantt" && !!c.data.series.some((s) => /^column\b/i.test(s.name))],
    ["gantt working-day timeline", (c) => c.kind === "gantt" && c.gantt?.workdays != null],
    ["scatter marginal histograms", (c) => c.decorations?.marginals != null],
    ["bubble overlap relief", (c) => c.scatter?.spread != null],
  ];
  for (const [name, test] of FEATURES) {
    it(`demonstrates ${name}`, () => {
      expect(configs.some(test)).toBe(true);
    });
  }

  const GANTT_ROWS = ["After", "Today", "Holiday", "Bracket", "Column", "X line", "Y line", "Trend", "Group"];
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
