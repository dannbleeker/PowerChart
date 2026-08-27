/**
 * How long an insert onto a given slide is likely to take.
 *
 * SSF Charts is fast exactly when it CREATES a slide and slow exactly when it
 * ADDS to one. A chart on a new slide costs about 0.75s; the same chart onto a
 * slide that already holds content costs about 24s — roughly 32x, measured over
 * 2,917 timed batches in 169 rounds. The cause is the host's, not ours: adding
 * to an existing slide has to draw shape by shape, where a new slide can be
 * handed over as a generated file.
 *
 * The user cannot see any of that. This is the number that lets the pane say so.
 *
 * THE CURVE IS THE ARCHIVE'S, NOT A MODEL. These four points are the measured
 * medians, per ten-shape batch, bucketed by how many shapes the target slide
 * already held. Interpolating between them is the most honest estimate
 * available — a fitted formula would look more precise and be less true, since
 * the curve visibly saturates and no fit this project has data for explains why.
 *
 *     shapes already on the slide    median batch
 *       0                              3886ms
 *       1-20                           5490ms
 *      21-50                          13995ms
 *      51-100                         18074ms
 *
 * It is an ESTIMATE and the pane must present it as one. Real spread inside a
 * bucket is wide, the archive's own noise floor is 14% IQR between rounds, and
 * a host having a bad minute can double any of this.
 */

/** Measured medians: [shapes already present, ms for one ten-shape batch]. */
const CURVE: readonly (readonly [number, number])[] = [
  [0, 3886],
  [10, 5490],
  [35, 13995],
  [75, 18074],
];

/** Shapes per batch, matching the renderer's own batching. */
const SHAPES_PER_BATCH = 10;

/**
 * Milliseconds for one batch onto a slide already holding `present` shapes.
 *
 * Linear between measured points, flat outside them. Flat rather than
 * extrapolated on the right: past 100 shapes this project has no readings, and
 * a straight line through the last two would keep climbing on no evidence.
 */
export function batchMs(present: number): number {
  const n = Math.max(0, present);
  const first = CURVE[0];
  const last = CURVE[CURVE.length - 1];
  if (n <= first[0]) return first[1];
  if (n >= last[0]) return last[1];
  for (let i = 1; i < CURVE.length; i++) {
    const [x0, y0] = CURVE[i - 1];
    const [x1, y1] = CURVE[i];
    if (n <= x1) return y0 + ((y1 - y0) * (n - x0)) / (x1 - x0);
  }
  return last[1];
}

/**
 * Estimated milliseconds to draw `shapes` onto a slide holding `present`.
 *
 * Uses the STARTING occupancy for every batch, which understates: the slide
 * fills as we draw, so later batches are dearer than the first. Understating is
 * the right direction for a number used to warn — a warning that fires late is
 * a nuisance, one that fires early on a fast insert is noise, and this project
 * would rather be quiet than cry wolf.
 */
export function estimateInsertMs(shapes: number, present: number): number {
  if (shapes <= 0) return 0;
  return Math.ceil(shapes / SHAPES_PER_BATCH) * batchMs(present);
}

/**
 * Above this, the pane offers to put the chart on its own slide instead.
 *
 * Fourteen seconds, set by the owner. It sits between the 1-20 bucket (~5.5s
 * for a ten-shape chart) and the 21-50 one (~14s), so an ordinary chart onto a
 * lightly-used slide stays silent and one onto a loaded slide speaks up.
 */
export const SLOW_INSERT_MS = 14_000;

/** Whether this insert is slow enough to be worth offering an alternative. */
export function isSlowInsert(shapes: number, present: number): boolean {
  return estimateInsertMs(shapes, present) > SLOW_INSERT_MS;
}

/** "about 20 seconds" / "about a minute" — for a sentence, not a readout. */
export function describeMs(ms: number): string {
  const s = Math.round(ms / 1000);
  // ROUNDED FIRST, THEN COMPARED. Rounding 58s to the nearest five gives 60,
  // and "about 60 seconds" is a phrase no one says. The check has to happen
  // after the rounding that can produce it, not before.
  const rounded = Math.max(1, Math.round(s / 5) * 5);
  if (rounded < 60) return `about ${rounded} seconds`;
  const m = Math.round(s / 30) / 2;
  return m <= 1 ? "about a minute" : `about ${m} minutes`;
}
