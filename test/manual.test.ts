import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { CHART_KINDS } from "../src/core/samples";

/**
 * Keeps the user manual honest, the same way skill-docs.test.ts keeps the
 * Agent Skill honest: when a chart kind, datasheet row, pane control, or
 * decoration ships, docs/MANUAL.md must mention it — otherwise CI fails.
 * Extend the manual (not this test's exception lists) when adding features.
 */
describe("user manual coverage", () => {
  const manual = readFileSync("docs/MANUAL.md", "utf8");
  const lower = manual.toLowerCase();

  for (const { kind, label } of CHART_KINDS) {
    it(`documents chart kind "${kind}"`, () => {
      // Kinds appear by their gallery label (the user-facing name).
      expect(lower).toContain(label.toLowerCase());
    });
  }

  it("documents every special datasheet row", () => {
    for (const row of [
      "100%=", "X extent", "X line", "Y line", "Trend",
      "Error+", "Error-",
      "Min", "Q1", "Median", "Q3", "Max", "Mean", "Outlier",
      "Start", "End", "Milestone", "After", "Today", "Holiday", "Bracket",
    ]) {
      expect(manual, `datasheet row ${row}`).toContain(row);
    }
  });

  it("documents the decorations and chart options", () => {
    for (const feature of [
      "Segment labels", "Series labels", "totals", "Value axis", "Gridlines",
      "Connector lines", "100% = note", "Datamark axis",
      "CAGR", "Difference arrow", "Value line", "Segment order",
      "Axis break", "label content", "Footnote", "Explode slices",
      "Same scale", "callouts", "background bands", "pattern fills", "slope",
      "breakout", "Small multiples",
    ]) {
      expect(manual, `option ${feature}`).toContain(feature);
    }
  });

  it("documents the pane workflows", () => {
    for (const flow of [
      "Edit selected chart", "Update chart", "Auto-update", "Transpose",
      "Use deck theme", "Export style", "Save as template",
      "Export current", "Import", "Insert batch",
      "Horizontal (bar)", "formulas".toLowerCase(),
    ]) {
      expect(lower, `workflow ${flow}`).toContain(flow.toLowerCase());
    }
  });

  it("documents the elements and their in-cell tokens", () => {
    for (const el of ["Harvey ball", "Checkbox", "Process flow", "KPI tile", "Total row", "[hb:", "[up]", "[good]", "Agenda"]) {
      expect(manual, `element ${el}`).toContain(el);
    }
  });

  it("documents the Excel companion and the Claude skill", () => {
    expect(manual).toContain("manifest-excel.xml");
    expect(manual).toContain("powerchart-charts.zip");
  });
});
