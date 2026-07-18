/**
 * Fold-based min / max over an array.
 *
 * `Math.max(seed, ...arr)` passes every element as a separate call argument,
 * which throws `RangeError: Maximum call stack size exceeded` once the array
 * exceeds V8's argument limit (~10⁵ elements). That is reachable with a valid,
 * within-cap data grid: `normalizeData` allows up to MAX_CATEGORIES × MAX_SERIES
 * ≈ 10⁶ cells, and the extent/format passes flatten the whole grid before taking
 * a min/max — so a large-but-legal chart crashed instead of rendering.
 *
 * A reduce with the SAME two-argument `Math.max`/`Math.min` is byte-identical to
 * the spread (both are order-independent and share `NaN`/`±0` semantics) and
 * can't overflow the argument list. Use these for any array whose length scales
 * with the data (cells or samples); a fixed handful of scalars can stay inline.
 */
export const maxOf = (arr: readonly number[], seed = -Infinity): number => arr.reduce((m, v) => Math.max(m, v), seed);

export const minOf = (arr: readonly number[], seed = Infinity): number => arr.reduce((m, v) => Math.min(m, v), seed);
