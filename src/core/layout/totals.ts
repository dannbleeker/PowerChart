import type { Series } from "../types";

/**
 * Per-category column sums shared by the stacked-bar family (column, mekko,
 * radar, butterfly). Missing cells count as 0. Factored out of the four layouts
 * that each inlined the identical reduce — one definition to keep them in step.
 */

/** Sum of the positive parts of category column `c` (the stack's upward reach). */
export const columnPositiveTotal = (series: Series[], c: number): number =>
  series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0);

/** Sum of the negative parts of category column `c` (the stack's downward reach). */
export const columnNegativeTotal = (series: Series[], c: number): number =>
  series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0);

/** Signed sum of category column `c` — positives and negatives net out. */
export const columnSignedTotal = (series: Series[], c: number): number =>
  series.reduce((a, s) => a + (s.values[c] ?? 0), 0);
