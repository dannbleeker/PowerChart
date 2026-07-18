import { describe, expect, it } from "vitest";
import { columnNegativeTotal, columnPositiveTotal, columnSignedTotal } from "../src/core/layout/totals";
import type { Series } from "../src/core/types";

/**
 * The per-category column sums that the stacked-bar family (column, mekko,
 * radar) each used to inline verbatim. Pinning them here documents the
 * missing-cell-as-zero and positive/negative-split contract the layouts depend on.
 */
describe("column total helpers", () => {
  const series: Series[] = [
    { name: "a", values: [10, -5, null] },
    { name: "b", values: [-3, 4, 8] },
    { name: "c", values: [2, null, -1] },
  ];

  it("sums positive parts, treating missing cells as zero", () => {
    expect(columnPositiveTotal(series, 0)).toBe(12); // 10 + 2 (−3 excluded)
    expect(columnPositiveTotal(series, 1)).toBe(4); // 4 (−5 excluded, null → 0)
    expect(columnPositiveTotal(series, 2)).toBe(8); // 8 (null → 0, −1 excluded)
  });

  it("sums negative parts as a non-positive number", () => {
    expect(columnNegativeTotal(series, 0)).toBe(-3);
    expect(columnNegativeTotal(series, 1)).toBe(-5);
    expect(columnNegativeTotal(series, 2)).toBe(-1);
  });

  it("nets positives and negatives in the signed total", () => {
    expect(columnSignedTotal(series, 0)).toBe(9); // 10 − 3 + 2
    expect(columnSignedTotal(series, 1)).toBe(-1); // −5 + 4 + 0
    expect(columnSignedTotal(series, 2)).toBe(7); // 0 + 8 − 1
  });

  it("returns 0 for an empty series list", () => {
    expect(columnPositiveTotal([], 0)).toBe(0);
    expect(columnNegativeTotal([], 0)).toBe(0);
    expect(columnSignedTotal([], 0)).toBe(0);
  });
});
