import type { NumberFormat } from "./types";

export const DEFAULT_FORMAT: NumberFormat = { decimals: "auto" };

/** Format a value the way think-cell's default label format does: compact, thousands-separated. */
export function formatNumber(v: number, fmt: Partial<NumberFormat> = {}): string {
  const f = { ...DEFAULT_FORMAT, ...fmt };
  const abs = Math.abs(v);
  const decimals =
    f.decimals === "auto" ? (abs !== 0 && abs < 1 ? 2 : abs < 10 ? 1 : 0) : f.decimals;
  let s = v.toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
  if (f.forceSign && v > 0) s = "+" + s;
  if (f.suffix) s += f.suffix;
  return s;
}

/**
 * Resolve "auto" decimals once per chart from the data's magnitude, so all
 * labels in one chart share the same precision (as think-cell does).
 */
export function resolveFormat(values: number[], fmt: Partial<NumberFormat> = {}): NumberFormat {
  if (fmt.decimals != null && fmt.decimals !== "auto") {
    return { ...DEFAULT_FORMAT, ...fmt, decimals: fmt.decimals };
  }
  const maxAbs = Math.max(0, ...values.filter((v) => Number.isFinite(v)).map(Math.abs));
  const decimals = maxAbs >= 10 ? 0 : maxAbs >= 1 ? 1 : 2;
  return { ...DEFAULT_FORMAT, ...fmt, decimals };
}

export function formatPercent(v: number, decimals = 0, forceSign = false): string {
  const s = (v * 100).toFixed(decimals) + "%";
  return forceSign && v > 0 ? "+" + s : s;
}

/**
 * "Nice" axis ticks covering [min, max] with roughly `count` steps.
 * Returns the tick values including the padded ends.
 */
export function niceTicks(min: number, max: number, count = 5): number[] {
  if (min === max) {
    if (min === 0) return [0, 1];
    min = Math.min(0, min);
    max = Math.max(0, max);
    if (min === max) max = min + 1;
  }
  const span = max - min;
  const rawStep = span / Math.max(1, count);
  const mag = Math.pow(10, Math.floor(Math.log10(rawStep)));
  const norm = rawStep / mag;
  const step = (norm >= 5 ? 10 : norm >= 2 ? 5 : norm >= 1 ? 2 : 1) * mag;
  const lo = Math.floor(min / step) * step;
  const hi = Math.ceil(max / step) * step;
  const ticks: number[] = [];
  // Guard against FP drift producing an extra/short tick.
  for (let i = 0; i <= Math.round((hi - lo) / step); i++) {
    ticks.push(+(lo + i * step).toPrecision(12));
  }
  return ticks;
}

/** Compound annual growth rate between two values over `periods` steps. */
export function cagr(from: number, to: number, periods: number): number | null {
  if (periods <= 0 || from <= 0 || to <= 0) return null;
  return Math.pow(to / from, 1 / periods) - 1;
}
