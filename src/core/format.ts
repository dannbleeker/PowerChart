import type { NumberFormat } from "./types";

export const DEFAULT_FORMAT: NumberFormat = { decimals: "auto" };

/**
 * Intl.NumberFormat instances are expensive to construct but immutable and
 * reusable, and a chart formats hundreds of labels sharing a handful of
 * (locale, decimals) pairs. `Number.prototype.toLocaleString(locale, opts)` is
 * specified to construct a fresh NumberFormat on every call, so caching by
 * (locale, decimals) and reusing `.format()` is byte-identical output at a
 * fraction of the cost.
 */
const NUMBER_FORMATTERS = new Map<string, Intl.NumberFormat>();
function numberFormatter(locale: string, decimals: number): Intl.NumberFormat {
  const key = `${locale} ${decimals}`;
  let nf = NUMBER_FORMATTERS.get(key);
  if (!nf) {
    nf = new Intl.NumberFormat(locale, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
    NUMBER_FORMATTERS.set(key, nf);
  }
  return nf;
}

/** Format a value the way think-cell's default label format does: compact, thousands-separated. */
export function formatNumber(v: number, fmt: Partial<NumberFormat> = {}): string {
  const f = { ...DEFAULT_FORMAT, ...fmt };
  const abs = Math.abs(v);
  const decimals =
    f.decimals === "auto" ? (abs !== 0 && abs < 1 ? 2 : abs < 10 ? 1 : 0) : f.decimals;
  // A small negative that rounds toward zero would print as "-0". Normalise the
  // VALUE, not the formatted string: Intl renders the sign as U+2212 in some
  // locales, prefixes an invisible directional mark in RTL ones, and uses
  // non-Latin digits in others — no pattern match on the output survives all
  // three, and -0 leaked through in every locale that does any of them.
  if (Number(v.toFixed(decimals)) === 0) v = 0;
  let s = numberFormatter(f.locale ?? "en-US", decimals).format(v);
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
  // toFixed is not locale-aware, so the "-0" it can produce is always ASCII.
  let n = (v * 100).toFixed(decimals);
  if (/^-0(\.0+)?$/.test(n)) n = n.slice(1);
  const s = n + "%";
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
  // Numeric ranges ("3-5", "10–20") are category labels, not dates — Date.parse
  // would otherwise misread them as partial ISO dates.
  if (/^\d{1,3}\s*[-–]\s*\d{1,3}$/.test(t)) return null;
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

/**
 * OLS trend statistics: R² and the two-tailed p-value of the slope
 * (Student's t via the regularized incomplete beta function). Good charts
 * always state fit and significance next to a trend line.
 */
export function trendStats(pts: { x: number; y: number }[]): { r2: number; p: number | null } | null {
  const n = pts.length;
  if (n < 2) return null;
  const mx = pts.reduce((s, p) => s + p.x, 0) / n;
  const my = pts.reduce((s, p) => s + p.y, 0) / n;
  const sxx = pts.reduce((s, p) => s + (p.x - mx) ** 2, 0);
  const syy = pts.reduce((s, p) => s + (p.y - my) ** 2, 0);
  const sxy = pts.reduce((s, p) => s + (p.x - mx) * (p.y - my), 0);
  if (sxx <= 0 || syy <= 0) return null;
  const r2 = (sxy * sxy) / (sxx * syy);
  const df = n - 2;
  if (df < 1) return { r2, p: null };
  if (r2 >= 1) return { r2: 1, p: 0 };
  const t2 = (r2 * df) / (1 - r2);
  // Two-tailed p for Student's t: p = I_{df/(df+t²)}(df/2, 1/2).
  return { r2, p: betaI(df / 2, 0.5, df / (df + t2)) };
}

/** Human p-value: "< 0.001", "< 0.01", "< 0.05", or "= 0.31". */
export function formatP(p: number): string {
  for (const cut of [0.001, 0.01, 0.05]) if (p < cut) return `< ${cut}`;
  return `= ${p.toFixed(2)}`;
}

/** Regularized incomplete beta I_x(a, b) via continued fraction (NR-style). */
function betaI(a: number, b: number, x: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const lnBeta =
    lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x);
  const front = Math.exp(lnBeta);
  const symmetric = x >= (a + 1) / (a + b + 2);
  const [aa, bb, xx] = symmetric ? [b, a, 1 - x] : [a, b, x];
  // Lentz's continued fraction.
  let c = 1;
  let d = 1 - ((aa + bb) * xx) / (aa + 1);
  if (Math.abs(d) < 1e-30) d = 1e-30;
  d = 1 / d;
  let f = d;
  for (let m = 1; m <= 200; m++) {
    let num = (m * (bb - m) * xx) / ((aa + 2 * m - 1) * (aa + 2 * m));
    d = 1 + num * d;
    if (Math.abs(d) < 1e-30) d = 1e-30;
    c = 1 + num / c;
    if (Math.abs(c) < 1e-30) c = 1e-30;
    d = 1 / d;
    f *= d * c;
    num = (-(aa + m) * (aa + bb + m) * xx) / ((aa + 2 * m) * (aa + 2 * m + 1));
    d = 1 + num * d;
    if (Math.abs(d) < 1e-30) d = 1e-30;
    c = 1 + num / c;
    if (Math.abs(c) < 1e-30) c = 1e-30;
    d = 1 / d;
    const delta = d * c;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-10) break;
  }
  // `f` is Numerical Recipes' betacf continued fraction, so I_x(a,b) = front·f/a.
  // (An earlier `front·(f-1)/a` mixed in a different formulation's offset and made
  // every p-value wrong — e.g. p=0.014 where the true value is 0.205.)
  const result = (front * f) / aa;
  return symmetric ? 1 - result : result;
}

/** Lanczos log-gamma. */
function lnGamma(z: number): number {
  const g = [
    676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
    12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
  ];
  if (z < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * z)) - lnGamma(1 - z);
  z -= 1;
  let a = 0.99999999999980993;
  for (let i = 0; i < g.length; i++) a += g[i] / (z + i + 1);
  const t = z + g.length - 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
}
