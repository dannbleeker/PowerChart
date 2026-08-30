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

/**
 * Whether category `c` has ANY measured cell — the question the sums above
 * cannot answer.
 *
 * "Missing cells count as 0" is right for every sum here: a stack's reach is
 * the sum of what it has, and an absent segment is the right drawing. It stops
 * being right the moment a sum becomes a LABEL. A category whose every series
 * is blank sums to 0 and printed `total-3 = "0"`, which tells the reader the
 * business measured nothing when the truth is that nobody filled the cells in
 * — and a reader cannot tell those apart from the slide. Stacked and combo both
 * did it, on the most ordinary input there is: one empty column.
 *
 * So the arithmetic above is unchanged and the label asks this first. Only the
 * assertion is withheld, never the drawing. `test/blank-is-not-zero.test.ts`
 * sweeps every kind the engine ships for exactly this.
 */
export const columnHasData = (series: Series[], c: number): boolean =>
  series.some((s) => {
    const v = s.values[c];
    return typeof v === "number" && Number.isFinite(v);
  });
