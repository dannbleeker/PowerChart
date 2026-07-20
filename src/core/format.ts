import type { NumberFormat } from "./types";
import { maxOf } from "./agg";

export const DEFAULT_FORMAT: NumberFormat = { decimals: "auto" };

/**
 * Coerce a decimals count into the range `toFixed`/`Intl.NumberFormat` accept
 * (0–100). An authored or hand-edited `numberFormat.decimals` of -1 or 500
 * would otherwise throw a RangeError out of `toFixed`, aborting the whole
 * render — the same class of bad input the locale try/catch already repairs.
 */
const safeDecimals = (d: number): number => (Number.isFinite(d) ? Math.min(100, Math.max(0, Math.trunc(d))) : 0);

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
    // A malformed BCP-47 tag (an authored config, a hand-edited shape tag)
    // makes the Intl constructor throw a RangeError — fall back to en-US
    // rather than let one bad locale abort the whole render.
    try {
      nf = new Intl.NumberFormat(locale, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
    } catch {
      nf = new Intl.NumberFormat("en-US", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
    }
    NUMBER_FORMATTERS.set(key, nf);
  }
  return nf;
}

/** Format a value the way think-cell's default label format does: compact, thousands-separated. */
export function formatNumber(v: number, fmt: Partial<NumberFormat> = {}): string {
  // A non-finite value (a divide-by-zero, an empty average, a stray Infinity in
  // authored data) would otherwise print the literal "NaN"/"Infinity" as chart
  // text — suppress the label instead of drawing a broken one.
  if (!Number.isFinite(v)) return "";
  const f = { ...DEFAULT_FORMAT, ...fmt };
  const abs = Math.abs(v);
  // Exact zero takes no fractional digits — "0", not "0.0". (The old test
  // `abs !== 0 && abs < 1` fell through to the `< 10 → 1 decimal` branch for 0.)
  const decimals = f.decimals === "auto" ? (abs === 0 ? 0 : abs < 1 ? 2 : abs < 10 ? 1 : 0) : safeDecimals(f.decimals);
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
  const maxAbs = maxOf(values.filter((v) => Number.isFinite(v)).map(Math.abs), 0);
  const decimals = maxAbs >= 10 ? 0 : maxAbs >= 1 ? 1 : 2;
  return { ...DEFAULT_FORMAT, ...fmt, decimals };
}

export function formatPercent(v: number, decimals = 0, forceSign = false): string {
  if (!Number.isFinite(v)) return "";
  decimals = safeDecimals(decimals);
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
  // A non-finite bound (from a NaN in the data extent) or an inverted range
  // (a reversed manual scale.min>max) makes `Math.log10(span)` NaN and turns
  // every tick into NaN — invisible geometry for the whole value axis. Repair
  // to a usable range instead of propagating the poison.
  if (!Number.isFinite(min) || !Number.isFinite(max)) return [0, 1];
  if (min > max) [min, max] = [max, min];
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

/**
 * Least-squares polynomial fit of degree `degree` (2–4 for scatter's higher-order
 * trend lines). Solves the normal equations of the Vandermonde system in a
 * centered AND unit-scaled variable u = (x − mx) / sx, and returns an evaluator
 * plus the fit's R² and the degree actually used — clamped to points − 2, so at
 * least one residual degree of freedom remains and the fit never interpolates
 * noise (a degree n−1 fit through n points has a meaningless R² of 1).
 */
export function polyTrend(
  pts: { x: number; y: number }[],
  degree: number,
): { at: (x: number) => number; r2: number; degree: number } | null {
  const n = pts.length;
  if (n < 2) return null;
  const d = Math.max(1, Math.min(Math.floor(degree), n - 2));
  const mx = pts.reduce((s, p) => s + p.x, 0) / n;
  const my = pts.reduce((s, p) => s + p.y, 0) / n;
  // Unit-scale the centered abscissa. Centering fixes LARGE-x conditioning, but
  // the power sums S[k] scale as span^k, so a SMALL x-span drives S[2d] below the
  // solver's 1e-12 pivot floor and the whole trendline silently vanished. In u the
  // sums are O(1) regardless of span; `at` divides by the same sx, so the fit is
  // identical, only better conditioned.
  const sx = pts.reduce((s, p) => Math.max(s, Math.abs(p.x - mx)), 0) || 1;
  const m = d + 1;
  // Power sums in u = (x − mx) / sx: S[k] = Σ uᵏ (k=0..2d), T[k] = Σ uᵏ·y (k=0..d).
  const S = new Array(2 * d + 1).fill(0);
  const T = new Array(m).fill(0);
  for (const p of pts) {
    const u = (p.x - mx) / sx;
    let up = 1;
    for (let k = 0; k <= 2 * d; k++) {
      S[k] += up;
      if (k < m) T[k] += up * p.y;
      up *= u;
    }
  }
  const A = Array.from({ length: m }, (_, i) => S.slice(i, i + m));
  const c = gaussSolve(A, T.slice());
  if (!c) return null;
  const at = (x: number) => {
    const u = (x - mx) / sx;
    let up = 1;
    let y = 0;
    for (let k = 0; k < m; k++) {
      y += c[k] * up;
      up *= u;
    }
    return y;
  };
  let ssRes = 0;
  let ssTot = 0;
  for (const p of pts) {
    ssRes += (p.y - at(p.x)) ** 2;
    ssTot += (p.y - my) ** 2;
  }
  const r2 = ssTot > 0 ? Math.max(0, 1 - ssRes / ssTot) : 1;
  return { at, r2, degree: d };
}

/** Gaussian elimination with partial pivoting; null if the matrix is singular. */
function gaussSolve(A: number[][], b: number[]): number[] | null {
  const n = b.length;
  for (let col = 0; col < n; col++) {
    let piv = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(A[r][col]) > Math.abs(A[piv][col])) piv = r;
    if (Math.abs(A[piv][col]) < 1e-12) return null;
    [A[col], A[piv]] = [A[piv], A[col]];
    [b[col], b[piv]] = [b[piv], b[col]];
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const f = A[r][col] / A[col][col];
      for (let k = col; k < n; k++) A[r][k] -= f * A[col][k];
      b[r] -= f * b[col];
    }
  }
  return b.map((v, i) => v / A[i][i]);
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
  const lnBeta = lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x);
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
    676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905,
    -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
  ];
  if (z < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * z)) - lnGamma(1 - z);
  z -= 1;
  let a = 0.99999999999980993;
  for (let i = 0; i < g.length; i++) a += g[i] / (z + i + 1);
  const t = z + g.length - 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
}

/**
 * Count `values` into `k` equal bins tiling [lo, hi]. The caller supplies k
 * rather than a rule, because the only k worth using here is one derived from
 * the axis's own tick grid — see layoutScatter's marginals, where k is a
 * multiple of the tick-interval count so every tick is a bin edge. A rule
 * keyed off the sample size alone (Sturges, Freedman-Diaconis) produces edges
 * that land between the ticks, which is exactly what a marginal histogram
 * must not do: it is read against the axis beside it.
 *
 * A value on an interior edge counts to the upper bin; hi counts to the last.
 * Values outside [lo, hi] are ignored — the caller's domain is the axis, and
 * the axis already covers the data.
 */
export function histogramBins(values: number[], lo: number, hi: number, k: number): number[] {
  const counts = new Array(Math.max(1, k)).fill(0);
  const span = hi - lo;
  if (!(span > 0)) return counts;
  for (const v of values) {
    if (v < lo || v > hi) continue;
    const i = Math.min(counts.length - 1, Math.floor(((v - lo) / span) * counts.length));
    counts[i]++;
  }
  return counts;
}
