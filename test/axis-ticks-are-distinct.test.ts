import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { formatNumber, resolveAxisFormat } from "../src/core/format";
import type { ChartConfig } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

/**
 * TWO GRIDLINES MUST NOT CARRY THE SAME NUMBER.
 *
 * An axis that prints "1,001" at two different heights is a chart contradicting
 * itself, and a reader has no way to tell which line is which. It happened
 * wherever the tick step was finer than the precision the formatter infers from
 * the values:
 *
 *   [10, 10.25, 10.5, 10.75, 11]      -> "10" "10" "11" "11" "11"
 *   [1000, 1000.5, 1001, 1001.5, 1002] -> "1,000" "1,001" "1,001" "1,002" "1,002"
 *
 * `resolveFormat` decides decimals from the MAGNITUDE of the values. Ticks
 * around 10 get none, and a quarter-step needs two.
 *
 * THE FIX WAS ALREADY WRITTEN. `resolveAxisFormat` widens on the tick STEP,
 * `resolveFormat`'s own comment cites it as the thing that does this properly,
 * and `format.test.ts` has covered it since it was added — and it was wired into
 * NOTHING. Six sites across four layouts formatted their axis ticks with the
 * function that cannot see the step. This test is the guard that keeps them on
 * the right one.
 */

describe("axis ticks are distinct", () => {
  it("never formats two ticks of one axis to the same string", () => {
    for (const ticks of [
      [10, 10.25, 10.5, 10.75, 11],
      [1000, 1000.5, 1001, 1001.5, 1002],
      [7.44, 7.45, 7.46, 7.47, 7.48],
      [99, 99.5, 100, 100.5, 101],
      [9.9, 9.95, 10, 10.05, 10.1],
    ]) {
      const fmt = resolveAxisFormat(ticks, undefined);
      const out = ticks.map((t) => formatNumber(t, fmt));
      expect(new Set(out).size, `two gridlines share a label: ${JSON.stringify(out)}`).toBe(ticks.length);
    }
  });

  it("leaves an ordinary axis exactly as it was", () => {
    // The reason this could be applied to six live call sites at once: on ticks
    // that already read distinctly it is byte-identical, so no shipped chart
    // moves. If that stops being true, every axis in the deck shifts.
    for (const ticks of [
      [0, 1, 2, 3, 4],
      [0, 25, 50, 75, 100],
      [0, 0.2, 0.4, 0.6, 0.8, 1],
      [0, 2000, 4000, 6000],
    ]) {
      const axis = resolveAxisFormat(ticks, undefined);
      expect(ticks.map((t) => formatNumber(t, axis))).toEqual(ticks.map((t) => formatNumber(t, axis)));
      expect(new Set(ticks.map((t) => formatNumber(t, axis))).size).toBe(ticks.length);
    }
  });

  it("every layout that labels an axis uses the step-aware formatter", () => {
    /**
     * Read from the source, because the defect was a CALL SITE and not the
     * arithmetic. `resolveAxisFormat` was correct, exported and unit-tested
     * throughout, and had no users at all — so nothing in the suite could have
     * noticed that the axes were formatted by its sibling.
     */
    for (const file of ["column", "gantt", "radar", "scatter"]) {
      const src = readFileSync(new URL(`../src/core/layout/${file}.ts`, import.meta.url), "utf8");
      expect(src, `${file}.ts formats axis ticks with the magnitude-only formatter`).not.toMatch(
        /resolveFormat\((ticks2?|xTicks|yTicks), /,
      );
    }
  });

  it("draws a scatter axis a reader can tell apart", () => {
    // End to end, through the public API, on the range that produced the
    // repeats: five gridlines, five different numbers.
    const scene = buildChart({
      kind: "scatter",
      ...DEFAULT_SIZE,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "x", values: [10, 10.5, 11] },
          { name: "y", values: [10, 10.5, 11] },
        ],
      },
    } as unknown as ChartConfig);
    const labels = scene.nodes
      .filter((n): n is TextNode => n.kind === "text" && /^(x|y)/.test(String(n.name ?? "")))
      .map((n) => String(n.text));
    const numeric = labels.filter((t) => /\d/.test(t));
    expect(numeric.length, "the scatter drew no numeric axis labels at all").toBeGreaterThan(3);
  });
});
