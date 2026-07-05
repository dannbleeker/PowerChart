import type { NumberFormat } from "./types";

export const DEFAULT_FORMAT: NumberFormat = { decimals: "auto" };

/** Format a value the way think-cell's default label format does: compact, thousands-separated. */
export function formatNumber(v: number, fmt: Partial<NumberFormat> = {}): string {
  const f = { ...DEFAULT_FORMAT, ...fmt };
  const abs = Math.abs(v);
  const decimals =
    f.decimals === "auto" ? (abs !== 0 && abs < 1 ? 2 : abs < 10 ? 1 : 0) : f.decimals;
  let s = v.toLocaleString(f.locale ?? "en-US", {
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

/**
 * Segment label text from think-cell's label-content dropdown: any ordered
 * combination of value, percent (of column), series and category names.
 */
export function segmentLabel(
  parts: ("value" | "percent" | "series" | "category")[],
  ctx: { value: number; fraction: number | null; series: string; category: string; fmt: Partial<NumberFormat> },
): string {
  return parts
    .map((p) => {
      switch (p) {
        case "value":
          return formatNumber(ctx.value, ctx.fmt);
        case "percent":
          return ctx.fraction == null ? null : formatPercent(ctx.fraction);
        case "series":
          return ctx.series;
        case "category":
          return ctx.category;
      }
    })
    .filter(Boolean)
    .join(p2sep(parts));
}

/** Multi-part labels read best on one line for two parts, else spaced. */
function p2sep(parts: string[]): string {
  return parts.length > 1 ? " " : "";
}

const DAY_MS = 86400000;

/**
 * Parse a calendar-date cell ("2026-01-15", "15.01.2026", "Jan 2026", …)
 * into days since the Unix epoch. Returns null for non-dates.
 */
export function parseDateToken(raw: string): number | null {
  const t = raw.trim();
  if (!t) return null;
  const dmy = t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!dmy && /^[-+]?[\d,.]+$/.test(t)) return null; // plain numbers are not dates
  const ms = dmy
    ? Date.UTC(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]))
    : Date.parse(/^\d{4}-\d{2}(-\d{2})?$/.test(t) ? t : `${t} UTC`);
  if (!Number.isFinite(ms)) return null;
  return Math.round(ms / DAY_MS);
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/** Short label for an epoch-day value: "5 Jan" or "Jan 26" on month starts. */
export function formatDay(days: number, withYear = false): string {
  const d = new Date(days * DAY_MS);
  const m = MONTHS[d.getUTCMonth()];
  if (withYear) return `${m} ${String(d.getUTCFullYear()).slice(2)}`;
  return d.getUTCDate() === 1 ? m : `${d.getUTCDate()} ${m}`;
}

/** Epoch-day values of every Monday covering [minDay, maxDay]. */
export function weekStarts(minDay: number, maxDay: number): number[] {
  // Day 0 (1970-01-01) was a Thursday; Monday ≡ 4 (mod 7).
  const first = minDay + ((4 - (minDay % 7) + 7) % 7);
  const out: number[] = [];
  for (let d = first; d <= maxDay && out.length < 120; d += 7) out.push(d);
  return out;
}

/** Epoch-day values of every month start covering [minDay, maxDay]. */
export function monthStarts(minDay: number, maxDay: number): number[] {
  const start = new Date(minDay * DAY_MS);
  let y = start.getUTCFullYear();
  let m = start.getUTCMonth();
  const out: number[] = [];
  for (let guard = 0; guard < 240; guard++) {
    const day = Date.UTC(y, m, 1) / DAY_MS;
    if (day > maxDay) break;
    if (day >= minDay) out.push(day);
    m++;
    if (m === 12) {
      m = 0;
      y++;
    }
  }
  return out;
}

/** Compound annual growth rate between two values over `periods` steps. */
export function cagr(from: number, to: number, periods: number): number | null {
  if (periods <= 0 || from <= 0 || to <= 0) return null;
  return Math.pow(to / from, 1 / periods) - 1;
}
