import type { ChartConfig, ChartData, ChartKind, ChartStyle, Decorations } from "./types";
import type { Scene } from "./scene";
import { maxOf, minOf } from "./agg";
import { DEFAULT_DECOR, DEFAULT_STYLE, seriesColor } from "./style";
import { layoutColumns, layoutCombo } from "./layout/column";
import { layoutWaterfall, waterfallExtent } from "./layout/waterfall";
import { layoutMekko } from "./layout/mekko";
import { layoutLine } from "./layout/line";
import { layoutButterfly } from "./layout/butterfly";
import { layoutScatter, spreadCap } from "./layout/scatter";
import { layoutGantt } from "./layout/gantt";
import { layoutPie } from "./layout/pie";
import { boxplotExtent, layoutBoxplot } from "./layout/boxplot";
import { layoutRadar } from "./layout/radar";
import { layoutHeatmap } from "./layout/heatmap";
import { layoutTilemap } from "./layout/tilemap";
import { layoutCascade } from "./layout/cascade";
import { layoutFunnel } from "./layout/funnel";
import { layoutWaffle } from "./layout/waffle";
import { layoutTreemap } from "./layout/treemap";
import { layoutSunburst } from "./layout/sunburst";
import { layoutViolin } from "./layout/violin";
import { layoutCandlestick } from "./layout/candlestick";
import { titleHeight, titleNode } from "./layout/frame";
import { bandNodes, decorationNodes } from "./decor";
import { resolveLabelCollisions } from "./collide";
import { formatNumber, niceTicks, resolveFormat } from "./format";
import type { SceneNode } from "./scene";
import type { LayoutResult } from "./layout/column";

export const DEFAULT_SIZE = { width: 480, height: 300 };

/** A positive, finite dimension, or the fallback for a zero/negative/NaN one. */
function clampDim(v: number | undefined, fallback: number): number {
  return typeof v === "number" && Number.isFinite(v) && v > 0 ? v : fallback;
}

/** Coerce a data block to trustworthy arrays: every series padded to the
 *  category count, non-finite cells nulled, per-category arrays aligned. */
// Hard ceilings on grid size. buildChart is fed arbitrary authored JSON (the
// skill) and configs round-tripped from shape tags, so a config claiming
// millions of categories × thousands of series would otherwise make normalizeData
// allocate ~10⁹ cells and OOM the headless renderer. These caps sit far above any
// real chart (a multi-year daily calendar heatmap is ~1–4k days, business charts
// rarely exceed a few dozen series), so truncating past them is a no-op for valid
// input and a bound on abusive input.
const MAX_CATEGORIES = 4096;
const MAX_SERIES = 256;

function normalizeData(data: ChartData): ChartData {
  const categories = (Array.isArray(data?.categories) ? data.categories : []).slice(0, MAX_CATEGORIES);
  const n = categories.length;
  const cell = (v: number | null | undefined): number | null => (v == null ? null : Number.isFinite(v) ? v : null);
  const series = (Array.isArray(data?.series) ? data.series : []).slice(0, MAX_SERIES).map((s) => {
    const raw = Array.isArray(s?.values) ? s.values : [];
    const values = Array.from({ length: n }, (_, c) => cell(raw[c]));
    if (s?.colors) {
      return { ...s, values, colors: Array.from({ length: n }, (_, c) => s.colors![c] ?? null) };
    }
    return { ...s, values };
  });
  const pad = <T>(arr: (T | null)[] | undefined): (T | null)[] | undefined =>
    arr ? Array.from({ length: n }, (_, c) => arr[c] ?? null) : arr;
  return { ...data, categories, series, hundredPercent: pad(data.hundredPercent), xExtent: pad(data.xExtent) };
}

/** Drop label nudges carrying a non-finite delta (they corrupt a node's coords). */
function cleanOffsets(offs: Record<string, { dx: number; dy: number }>): Record<string, { dx: number; dy: number }> {
  const out: Record<string, { dx: number; dy: number }> = {};
  for (const [k, o] of Object.entries(offs)) {
    if (o && Number.isFinite(o.dx) && Number.isFinite(o.dy)) out[k] = o;
  }
  return out;
}

/**
 * Coerce a possibly-malformed ChartConfig into one the layout engine can trust.
 *
 * buildChart is now fed by the Claude skill (arbitrary authored JSON) and by
 * configs round-tripped out of slide shape tags, so it can no longer assume its
 * input is well-formed. This REPAIRS rather than throws — a chart with a clamped
 * size beats a stack trace in a headless render — and every branch is a no-op
 * unless the field is actually out of range, so valid configs render identically.
 */
export function normalizeConfig(cfg: ChartConfig): ChartConfig {
  const width = clampDim(cfg.width, DEFAULT_SIZE.width);
  const height = clampDim(cfg.height, DEFAULT_SIZE.height);

  // Order a reversed manual scale and drop non-finite ends — a NaN or inverted
  // bound reaches niceTicks and blanks the whole value axis.
  let scale = cfg.scale;
  if (scale) {
    let min = typeof scale.min === "number" && Number.isFinite(scale.min) ? scale.min : undefined;
    let max = typeof scale.max === "number" && Number.isFinite(scale.max) ? scale.max : undefined;
    if (min != null && max != null && min > max) [min, max] = [max, min];
    scale = { ...(min != null ? { min } : {}), ...(max != null ? { max } : {}) };
  }

  return {
    ...cfg,
    width,
    height,
    scale,
    data: normalizeData(cfg.data),
    labelOffsets: cfg.labelOffsets ? cleanOffsets(cfg.labelOffsets) : cfg.labelOffsets,
  };
}

const SORTABLE: ChartKind[] = ["stacked", "clustered", "stacked100", "mekko", "pie", "doughnut", "butterfly"];

/** Reorder categories (and every per-category array) by column total. */
function sortCategories(cfg: ChartConfig): ChartConfig {
  if (!cfg.categorySort || !SORTABLE.includes(cfg.kind)) return cfg;
  const { data } = cfg;
  // Sort runs before error/target rows are extracted; those carried rows are
  // not real stack contributors, so exclude them from the ranking totals.
  const ranked = data.series.filter((s) => !CARRIED_ROW.test(s.name.trim()));
  const totals = data.categories.map((_, c) => ranked.reduce((a, s) => a + (s.values[c] ?? 0), 0));
  const sign = cfg.categorySort === "ascending" ? 1 : -1;
  const order = data.categories.map((_, c) => c).sort((a, b) => sign * (totals[a] - totals[b]));
  const pick = <T>(arr: T[] | undefined) => (arr ? order.map((c) => arr[c]) : undefined);
  return {
    ...cfg,
    data: {
      ...data,
      categories: order.map((c) => data.categories[c]),
      // colors is per-category too: leaving it in pre-sort order pins a
      // highlight to a screen position instead of to its data point.
      series: data.series.map((s) => ({
        ...s,
        values: order.map((c) => s.values[c]),
        ...(s.colors ? { colors: pick(s.colors) } : {}),
      })),
      hundredPercent: pick(data.hundredPercent),
      xExtent: pick(data.xExtent),
    },
  };
}

/** Column-family kinds whose series can be collapsed into an "Other" bucket. */
const OTHER_KINDS: ChartKind[] = ["stacked", "clustered", "stacked100"];

/**
 * Collapse the long tail of series into one "Other" segment: keep the
 * (max − 1) largest by absolute total and sum the rest into "Other", so the
 * result has at most `max` series. No-op when already within budget.
 */
function collapseOther(cfg: ChartConfig): ChartConfig {
  const ob = cfg.otherBucket;
  if (!ob || !OTHER_KINDS.includes(cfg.kind)) return cfg;
  const max = Math.max(2, Math.floor(ob.max ?? 5));
  const series = cfg.data.series;
  if (series.length <= max) return cfg;
  const totals = series.map((s) => s.values.reduce((a: number, v) => a + Math.abs(v ?? 0), 0));
  const rank = series.map((_, i) => i).sort((a, b) => totals[b] - totals[a]);
  const keep = new Set(rank.slice(0, max - 1));
  const kept = series.filter((_, i) => keep.has(i)); // original order preserved
  const otherVals = cfg.data.categories.map((_, c) =>
    series.reduce((a, s, i) => (keep.has(i) ? a : a + (s.values[c] ?? 0)), 0),
  );
  return {
    ...cfg,
    data: { ...cfg.data, series: [...kept, { name: "Other", values: otherVals }] },
  };
}

/**
 * Pareto helper: sort categories by the first (non-line) series descending and
 * overlay a computed cumulative-% line on a secondary axis — the classic 80/20
 * view. Rewrites the config into a combo.
 */
function applyPareto(cfg: ChartConfig): ChartConfig {
  if (!cfg.pareto) return cfg;
  // The bars are the first series that is neither overlay kind. Testing only
  // against "line" would rank the chart by a marker series when one is written
  // first — sorting the columns by their own benchmark.
  const bar = cfg.data.series.find((s) => s.type !== "line" && s.type !== "marker");
  if (!bar) return cfg;
  const order = cfg.data.categories.map((_, c) => c).sort((a, b) => (bar.values[b] ?? 0) - (bar.values[a] ?? 0));
  const pick = <T>(arr: T[] | undefined) => (arr ? order.map((c) => arr[c]) : undefined);
  const barVals = order.map((c) => bar.values[c] ?? 0);
  const total = barVals.reduce((a, v) => a + Math.max(0, v), 0) || 1;
  let run = 0;
  const cum = barVals.map((v) => {
    run += Math.max(0, v);
    return Math.round((run / total) * 1000) / 10;
  });
  return {
    ...cfg,
    kind: "combo",
    secondaryAxis: true,
    data: {
      ...cfg.data,
      categories: order.map((c) => cfg.data.categories[c]),
      series: [
        // colors follows the permutation, like values — Pareto's whole point is
        // that the ranked order changes, and a highlight has to travel with it.
        { ...bar, values: barVals, ...(bar.colors ? { colors: pick(bar.colors) } : {}) },
        { name: "Cumulative %", type: "line", values: cum },
      ],
      hundredPercent: pick(cfg.data.hundredPercent),
    },
  };
}

/** "After" / "Dep" row: a 1-based predecessor index, so it moves with the rows. */
const AFTER_ROW = /^(after|dep(endency)?)$/i;

/**
 * Gantt lanes: regroup task rows under one synthesized header per owner.
 *
 * A stable partition, not a sort — inside a lane the rows keep the order they
 * were written in, because that order is the plan. (categorySort refuses to
 * touch a Gantt for exactly that reason; this reorders only on request.) Tasks
 * with no owner keep to themselves at the end.
 *
 * The subtle part is `After`: its values are 1-based ROW INDICES, so moving
 * rows without renumbering them would silently point every dependency arrow at
 * the wrong task.
 */
function applyGanttLanes(cfg: ChartConfig): ChartConfig {
  if (cfg.kind !== "gantt" || cfg.gantt?.lanes !== "owner") return cfg;
  const { data } = cfg;
  const part = (c: string, i: number) => (c.split("|")[i] ?? "").trim();
  const ownerOf = (c: string) => part(c, 1);
  const lanes: string[] = [];
  for (const c of data.categories) {
    const o = ownerOf(c);
    if (o && !lanes.includes(o)) lanes.push(o);
  }
  if (!lanes.length) return cfg;

  /** New row order: -1 marks a synthesized header. */
  const perm: number[] = [];
  const labels: string[] = [];
  const push = (i: number) => {
    const c = data.categories[i];
    const act = part(c, 0).replace(/^>+\s*/, "");
    const remark = part(c, 2);
    // Drop the owner from the row — the lane header carries it now.
    labels.push(remark ? `> ${act} |  | ${remark}` : `> ${act}`);
    perm.push(i);
  };
  for (const lane of lanes) {
    labels.push(lane);
    perm.push(-1);
    data.categories.forEach((c, i) => {
      if (ownerOf(c) === lane) push(i);
    });
  }
  // Unassigned rows keep their own order, after the lanes.
  data.categories.forEach((c, i) => {
    if (!ownerOf(c)) push(i);
  });

  const oldToNew = new Map<number, number>();
  perm.forEach((old, next) => {
    if (old >= 0) oldToNew.set(old, next);
  });
  return {
    ...cfg,
    data: {
      ...data,
      categories: labels,
      series: data.series.map((s) => ({
        ...s,
        values: perm.map((old, next) => {
          if (old < 0) return null; // synthesized header: no bar
          const v = s.values[old];
          if (v == null || !AFTER_ROW.test(s.name.trim())) return v;
          // Renumber the dependency onto its row's new position.
          const target = oldToNew.get(Math.round(v) - 1);
          return target == null || target === next ? null : target + 1;
        }),
      })),
    },
  };
}

/** Datasheet rows carrying error-bar deltas: Error (±), Error+ / Error−. */
const ERROR_ROW = /^error\s*([+\-−])?$/i;
/** Bullet-chart target row: a bold tick across each column at the value. */
const TARGET_ROW = /^target$/i;
const ERROR_KINDS: ChartKind[] = ["stacked", "clustered", "line", "area"];
/** Kinds that accept a Target row (error bars stay column/line-family). */
const TARGET_KINDS: ChartKind[] = [...ERROR_KINDS, "waterfall"];

/**
 * Pull Error rows out of the data (so they don't render as segments) and
 * return per-category plus/minus deltas. Bars anchor at the column total
 * (single-series charts: the value) or the first line series.
 */
function extractErrorRows(cfg: ChartConfig): {
  cfg: ChartConfig;
  errors: { plus: (number | null)[]; minus: (number | null)[] } | null;
  targets: (number | null)[] | null;
} {
  if (!TARGET_KINDS.includes(cfg.kind)) return { cfg, errors: null, targets: null };
  const rows = ERROR_KINDS.includes(cfg.kind) ? cfg.data.series.filter((s) => ERROR_ROW.test(s.name.trim())) : [];
  const targetRow = cfg.data.series.find((s) => TARGET_ROW.test(s.name.trim()));
  if (!rows.length && !targetRow) return { cfg, errors: null, targets: null };
  const pick = (sign: "+" | "-") =>
    cfg.data.categories.map((_, c) => {
      for (const r of rows) {
        const m = r.name.trim().match(ERROR_ROW)!;
        const dir = m[1] === "−" ? "-" : (m[1] ?? "both");
        if ((dir === "both" || dir === sign) && r.values[c] != null) return Math.abs(r.values[c]!);
      }
      return null;
    });
  return {
    cfg: {
      ...cfg,
      data: {
        ...cfg.data,
        series: cfg.data.series.filter((s) => !ERROR_ROW.test(s.name.trim()) && !TARGET_ROW.test(s.name.trim())),
      },
    },
    errors: rows.length ? { plus: pick("+"), minus: pick("-") } : null,
    targets: targetRow ? cfg.data.categories.map((_, c) => targetRow.values[c] ?? null) : null,
  };
}

/** Kinds whose multi-series data splits cleanly into per-series panels. */
const MULTIPLES_KINDS: ChartKind[] = ["stacked", "clustered", "line", "area", "waterfall", "radar"];
/** Special rows carried into every small-multiples panel (not split). */
const CARRIED_ROW = /^(error\s*[+\-−]?|target|band\s*(low|high))$/i;

/** Shift a scene's nodes by (dx, dy) and prefix their names. */
function translateNodes(nodes: SceneNode[], dx: number, dy: number, prefix: string): SceneNode[] {
  return nodes.map((node) => {
    const n = { ...node, name: node.name ? `${prefix}${node.name}` : undefined };
    switch (n.kind) {
      case "rect":
      case "text":
      case "chevron":
      case "arrowhead":
        n.x += dx;
        n.y += dy;
        break;
      case "line":
        n.x1 += dx;
        n.y1 += dy;
        n.x2 += dx;
        n.y2 += dy;
        break;
      case "ellipse":
      case "wedge":
      case "symbol":
        n.cx += dx;
        n.cy += dy;
        break;
      case "polygon":
        n.points = n.points.map((p) => ({ x: p.x + dx, y: p.y + dy }));
        break;
      default: {
        // A new SceneNode kind that reaches here would be silently left at the
        // origin panel's coordinates — a misplaced shape, with nothing failing.
        // This turns that into a compile error at the point of adding the kind.
        const unreached: never = n;
        return unreached;
      }
    }
    return n;
  });
}

/**
 * Small multiples: one single-series panel per data series, titled by the
 * series name, laid out in a grid and pinned to one shared value scale so
 * panels compare honestly. Null when the config doesn't call for it.
 */
function buildMultiples(rawCfg: ChartConfig): Scene | null {
  const multiples = rawCfg.multiples;
  if (!multiples || !MULTIPLES_KINDS.includes(rawCfg.kind)) return null;
  // Sort categories ONCE, ranked across the full data, so every panel shares one
  // x-axis. (Letting each panel sort itself — it inherits categorySort — ranked
  // each panel by its own single series and gave the panels contradictory axes.)
  const cfg = sortCategories(rawCfg);
  const carried = cfg.data.series.filter((s) => CARRIED_ROW.test(s.name.trim()));
  const dataSeries = cfg.data.series.filter((s) => !CARRIED_ROW.test(s.name.trim()));
  if (dataSeries.length < 2) return null;

  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const fs = style.fontSize;
  const n = dataSeries.length;
  const cols = Math.max(1, Math.min(n, multiples.columns ?? (n <= 3 ? n : Math.ceil(Math.sqrt(n)))));
  const rows = Math.ceil(n / cols);
  const gap = 10;
  const titleH = titleHeight(cfg, style);
  const footH = cfg.footnote ? fs * 1.3 : 0;
  const panelW = (cfg.width - gap * (cols - 1)) / cols;
  const panelH = (cfg.height - titleH - footH - gap * (rows - 1)) / rows;

  const panelCfg = (s: (typeof dataSeries)[number], si: number): ChartConfig => ({
    ...cfg,
    multiples: undefined,
    categorySort: undefined, // already applied above, on the full data
    title: s.name,
    footnote: undefined,
    width: panelW,
    height: panelH,
    data: { ...cfg.data, series: [{ ...s, color: seriesColor(style, si, s.color) }, ...carried] },
    decorations: { ...cfg.decorations, seriesLabels: false },
  });

  // Per-panel extent that also covers the carried Error whiskers (base ± delta)
  // and Target rows — `valueExtent` alone treats an error row as a standalone
  // series, so a whisker or high target could render past the panel edge.
  const panelExtent = (s: (typeof dataSeries)[number], si: number): { min: number; max: number } | null => {
    const pcfg = panelCfg(s, si);
    const { cfg: baseCfg, errors, targets } = extractErrorRows(pcfg);
    const base = valueExtent(baseCfg);
    if (!base) return null;
    let { min, max } = base;
    const vals = baseCfg.data.series[0]?.values ?? [];
    if (errors) {
      vals.forEach((v, c) => {
        if (v == null) return;
        if (errors.plus[c] != null) max = Math.max(max, v + errors.plus[c]!);
        if (errors.minus[c] != null) min = Math.min(min, v - errors.minus[c]!);
      });
    }
    if (targets) {
      for (const t of targets)
        if (t != null) {
          max = Math.max(max, t);
          min = Math.min(min, t);
        }
    }
    return { min, max };
  };

  // Shared scale across panels; radar (no valueExtent) pins 0..global max.
  let scale = cfg.scale;
  if (scale?.min == null || scale?.max == null) {
    const exts = dataSeries.map((s, si) => panelExtent(s, si));
    const global = exts.every((e) => e != null)
      ? { min: Math.min(...exts.map((e) => e!.min)), max: Math.max(...exts.map((e) => e!.max)) }
      : {
          min: 0,
          max: maxOf(
            dataSeries.flatMap((s) => s.values.filter((v): v is number => v != null)),
            1,
          ),
        };
    scale = { min: scale?.min ?? global.min, max: scale?.max ?? global.max };
  }

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  dataSeries.forEach((s, si) => {
    const panel = buildChart({ ...panelCfg(s, si), scale });
    const dx = (si % cols) * (panelW + gap);
    const dy = titleH + Math.floor(si / cols) * (panelH + gap);
    nodes.push(...translateNodes(panel.nodes, dx, dy, `p${si}-`));
  });
  if (cfg.footnote) {
    nodes.push({
      kind: "text",
      x: 2,
      y: cfg.height - fs * 1.15,
      w: cfg.width - 4,
      h: fs * 1.1,
      text: cfg.footnote,
      fontSize: fs * 0.85,
      color: style.mutedText,
      align: "left",
      valign: "bottom",
      name: "footnote",
    });
  }
  return { width: cfg.width, height: cfg.height, nodes };
}

/** Human-readable chart-kind names for the accessible description. */
const KIND_LABEL: Record<ChartKind, string> = {
  stacked: "stacked column chart",
  clustered: "clustered column chart",
  stacked100: "100% stacked column chart",
  waterfall: "waterfall chart",
  mekko: "Mekko chart",
  line: "line chart",
  area: "area chart",
  butterfly: "butterfly chart",
  scatter: "scatter plot",
  bubble: "bubble chart",
  gantt: "Gantt chart",
  combo: "combination chart",
  pie: "pie chart",
  doughnut: "doughnut chart",
  boxplot: "box-and-whisker plot",
  radar: "radar chart",
  heatmap: "heatmap",
  tilemap: "tile-grid map",
  cascade: "cascade chart",
  funnel: "funnel chart",
  waffle: "waffle chart",
  treemap: "treemap",
  sunburst: "sunburst chart",
  violin: "violin plot",
  candlestick: "candlestick chart",
};

/**
 * A one-line text alternative for the chart — the accessible description a
 * screen reader reads after the title. Names the kind, the series and the
 * categories (scatter/bubble rows and points read as series/categories, which
 * is what the data model calls them). Deliberately concise; it is a summary,
 * not a data table.
 */
export function describeChart(cfg: ChartConfig): string {
  const label = KIND_LABEL[cfg.kind] ?? "chart";
  const cats = (cfg.data?.categories ?? []).filter(Boolean);
  const series = cfg.data?.series ?? [];
  const seriesNames = series.map((s) => s.name).filter(Boolean);
  const list = (xs: string[], n = 4) =>
    xs.length <= n ? xs.join(", ") : `${xs.slice(0, n).join(", ")} and ${xs.length - n} more`;
  const bits: string[] = [`${label}${cfg.horizontal ? " (horizontal)" : ""}`];
  if (seriesNames.length)
    bits.push(`${seriesNames.length} data ${seriesNames.length === 1 ? "series" : "series"}: ${list(seriesNames)}`);
  if (cats.length) bits.push(`${cats.length} ${cats.length === 1 ? "category" : "categories"}: ${list(cats)}`);
  return bits.join(". ") + ".";
}

/** Build a renderer-agnostic scene from a chart config. Pure and synchronous. */
export function buildChart(rawCfg: ChartConfig): Scene {
  const cfg0 = normalizeConfig(rawCfg);
  const a11y = { title: cfg0.title, desc: describeChart(cfg0) };
  const multiples = buildMultiples(cfg0);
  if (multiples) return { ...multiples, ...a11y };
  const extracted = extractErrorRows(sortCategories(applyPareto(applyGanttLanes(cfg0))));
  let cfg = collapseOther(extracted.cfg);
  const errors = extracted.errors;
  const targets = extracted.targets;
  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const decor: Decorations = { ...DEFAULT_DECOR, ...cfg.decorations };

  // Widen the auto scale so error bars and target ticks stay inside the plot.
  // valueExtent already folds the whiskers and ticks in — it has to, since Same
  // Scale reads it — so this only has to round the result to nice ticks.
  if ((errors || targets) && !cfg.horizontal && cfg.scale?.max == null) {
    const ext = drawnExtent(cfg, errors, targets);
    if (ext) {
      const ticks = niceTicks(ext.min, ext.max, 5);
      cfg = { ...cfg, scale: { ...cfg.scale, min: cfg.scale?.min ?? ticks[0], max: ticks[ticks.length - 1] } };
    }
  }

  let result: LayoutResult;
  switch (cfg.kind) {
    case "waterfall":
      result = layoutWaterfall(cfg, style, decor);
      break;
    case "mekko":
      result = layoutMekko(cfg, style, decor);
      break;
    case "line":
    case "area":
      result = layoutLine(cfg, style, decor);
      break;
    case "butterfly":
      result = layoutButterfly(cfg, style, decor);
      break;
    case "scatter":
    case "bubble":
      result = layoutScatter(cfg, style, decor);
      break;
    case "gantt":
      result = layoutGantt(cfg, style, decor);
      break;
    case "combo":
      result = layoutCombo(cfg, style, decor);
      break;
    case "pie":
    case "doughnut":
      result = layoutPie(cfg, style, decor);
      break;
    case "boxplot":
      result = layoutBoxplot(cfg, style, decor);
      break;
    case "radar":
      result = layoutRadar(cfg, style, decor);
      break;
    case "heatmap":
      result = layoutHeatmap(cfg, style, decor);
      break;
    case "tilemap":
      result = layoutTilemap(cfg, style, decor);
      break;
    case "cascade":
      result = layoutCascade(cfg, style, decor);
      break;
    case "funnel":
      result = layoutFunnel(cfg, style, decor);
      break;
    case "treemap":
      result = layoutTreemap(cfg, style, decor);
      break;
    case "sunburst":
      result = layoutSunburst(cfg, style, decor);
      break;
    case "violin":
      result = layoutViolin(cfg, style, decor);
      break;
    case "candlestick":
      result = layoutCandlestick(cfg, style, decor);
      break;
    case "waffle":
      result = layoutWaffle(cfg, style, decor);
      break;
    default:
      result = layoutColumns(cfg, style, decor);
  }

  // Decorations assume a vertical value axis; skip them for bar orientation
  // and butterfly charts.
  const skipDecor =
    cfg.horizontal ||
    [
      "butterfly",
      "scatter",
      "bubble",
      "gantt",
      "pie",
      "doughnut",
      "radar",
      "heatmap",
      "tilemap",
      "cascade",
      "funnel",
      "waffle",
    ].includes(cfg.kind);
  // Background bands go BEFORE the layout's nodes so they render behind the
  // data (scatter/bubble draw their own, in value units).
  const bands = !skipDecor && decor.bands?.length ? bandNodes(cfg, style, decor, result.anchors) : [];
  const nodes = skipDecor
    ? [...bands, ...result.nodes]
    : [...bands, ...result.nodes, ...decorationNodes(cfg, style, decor, result.anchors)];

  // Error bars from Error / Error+ / Error− rows: a whisker with caps at the
  // column total (or line point), on the shared value scale.
  if (errors && !skipDecor && result.anchors.valueToY) {
    const a = result.anchors;
    cfg.data.categories.forEach((_, c) => {
      const plus = errors.plus[c];
      const minus = errors.minus[c];
      if (plus == null && minus == null) return;
      const base = a.columnValue[c];
      const x = a.categoryX[c];
      const capW = Math.min(a.categoryWidth[c] * 0.35, 10);
      const yHi = a.valueToY!(base + (plus ?? 0));
      const yLo = a.valueToY!(base - (minus ?? 0));
      nodes.push({
        kind: "line",
        x1: x,
        y1: yHi,
        x2: x,
        y2: yLo,
        stroke: style.axis,
        strokeWidth: 1,
        name: `error-${c}`,
      });
      if (plus != null)
        nodes.push({
          kind: "line",
          x1: x - capW / 2,
          y1: yHi,
          x2: x + capW / 2,
          y2: yHi,
          stroke: style.axis,
          strokeWidth: 1,
          name: `error-cap-hi-${c}`,
        });
      if (minus != null)
        nodes.push({
          kind: "line",
          x1: x - capW / 2,
          y1: yLo,
          x2: x + capW / 2,
          y2: yLo,
          stroke: style.axis,
          strokeWidth: 1,
          name: `error-cap-lo-${c}`,
        });
    });
  }

  // Bullet-chart target ticks: a bold marker across each column at the
  // target value (combine with decorations.bands for the range zones).
  if (targets && !skipDecor && result.anchors.valueToY) {
    const a = result.anchors;
    cfg.data.categories.forEach((_, c) => {
      const t = targets[c];
      if (t == null) return;
      const x = a.categoryX[c];
      const half = Math.min(a.categoryWidth[c] * 0.62, 26);
      const y = a.valueToY!(t);
      nodes.push({
        kind: "line",
        x1: x - half,
        y1: y,
        x2: x + half,
        y2: y,
        stroke: style.text,
        strokeWidth: 2.25,
        name: `target-${c}`,
      });
      // Budget-vs-actual bridge: on a waterfall, a hatched gap segment shows
      // the distance from the achieved level to the target.
      if (cfg.kind === "waterfall") {
        const actual = a.columnValue[c];
        const gap = t - actual;
        if (Math.abs(gap) > 1e-9) {
          const yA = a.valueToY!(actual);
          const w = a.categoryWidth[c];
          nodes.push({
            kind: "rect",
            x: x - w / 2,
            y: Math.min(y, yA),
            w,
            h: Math.abs(yA - y),
            fill: style.neutral,
            pattern: "diagonal",
            stroke: style.mutedText,
            strokeWidth: 0.75,
            name: `target-gap-${c}`,
          });
          const fs = style.fontSize;
          nodes.push({
            kind: "text",
            x: x - w / 2 - 4,
            y: Math.min(y, yA) - fs * 1.5,
            w: w + 8,
            h: fs * 1.4,
            text: `Gap ${formatNumber(gap, { ...resolveFormat([t, actual], cfg.numberFormat), forceSign: true })}`,
            fontSize: fs * 0.95,
            bold: true,
            color: style.text,
            align: "center",
            valign: "bottom",
            name: `target-gap-label-${c}`,
          });
        }
      }
    });
  }

  // Footnote line: source citation and/or the "100% = N" note, bottom-left.
  const footParts: string[] = [];
  if (decor.hundredPercentNote) {
    const total = hundredPercentTotal(cfg);
    if (total != null) footParts.push(`100% = ${formatNumber(total, resolveFormat([total], cfg.numberFormat))}`);
  }
  if (cfg.footnote) footParts.push(cfg.footnote);
  // Overlap relief is an approximation, so the chart discloses it — quoting
  // the cap the layout actually enforces, not a restatement of it.
  const sc = spreadCap(cfg);
  if (sc) {
    footParts.push(
      `${sc.axis.toUpperCase()} positions approximate: markers spread by up to ±${formatNumber(sc.limit, resolveFormat([sc.limit], cfg.numberFormat))} to reduce overlap`,
    );
  }
  if (footParts.length) {
    const fs = style.fontSize;
    nodes.push({
      kind: "text",
      x: 2,
      y: cfg.height - fs * 1.15,
      w: cfg.width - 4,
      h: fs * 1.1,
      text: footParts.join("   ·   "),
      fontSize: fs * 0.85,
      color: style.mutedText,
      align: "left",
      valign: "bottom",
      name: "footnote",
    } satisfies SceneNode);
  }

  // Global de-collision for outside labels (vertical cartesian charts).
  if (!skipDecor) resolveLabelCollisions(nodes);

  // Manual label nudges (think-cell's label dragging, config-driven).
  if (cfg.labelOffsets) {
    for (const n of nodes) {
      const off = n.name && cfg.labelOffsets[n.name];
      if (off && n.kind === "text") {
        n.x += off.dx;
        n.y += off.dy;
      }
    }
  }
  return { width: cfg.width, height: cfg.height, nodes, ...a11y };
}

/**
 * The denominator behind a "100% = N" note: the series total for pies, the
 * uniform per-category denominator for 100% charts (null when categories
 * have different denominators — the note would be a lie then).
 */
function hundredPercentTotal(cfg: ChartConfig): number | null {
  const { data, kind } = cfg;
  if (kind === "pie" || kind === "doughnut") {
    const total = data.categories.reduce((a, _, c) => a + Math.max(0, data.series[0]?.values[c] ?? 0), 0);
    return total > 0 ? total : null;
  }
  if (kind === "stacked100") {
    const denominators = data.categories.map((_, c) => {
      const d = data.hundredPercent?.[c];
      return d != null && d > 0 ? d : data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0);
    });
    if (!denominators.length || denominators[0] <= 0) return null;
    return denominators.every((d) => Math.abs(d - denominators[0]) < 1e-9) ? denominators[0] : null;
  }
  return null;
}

/** The data's own range, ignoring the anatomy drawn on top of it. */
function dataExtent(cfg: ChartConfig): { min: number; max: number } | null {
  const { data, kind } = cfg;
  const cats = data.categories.map((_, c) => c);
  const vals = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  if (!vals.length) return null;
  switch (kind) {
    case "stacked": {
      const pos = cats.map((c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
      const neg = cats.map((c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
      return { min: minOf(neg, 0), max: maxOf(pos, 0) };
    }
    case "clustered":
    case "line":
      return { min: minOf(vals, 0), max: maxOf(vals, 0) };
    case "boxplot":
      return boxplotExtent(cfg);
    case "area": {
      // Areas stack under the baseline too (layout/line.ts dips negatives below
      // zero), so the extent must mirror "stacked" — a hard 0 floor here clipped
      // every negative area on the shared-scale paths (small multiples, Same Scale).
      const pos = cats.map((c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
      const neg = cats.map((c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
      return { min: minOf(neg, 0), max: maxOf(pos, 0) };
    }
    case "waterfall": {
      // One chain, shared with the layout. This used to be a second walk that
      // added only series[0] and skipped spacerIndices, so a stacked bridge
      // reported the first series' total as the whole chart's — 63 instead of
      // 114 — and Same scale then clipped the bars off the shape.
      return waterfallExtent(cfg);
    }
    default:
      return null;
  }
}

/**
 * Value-axis extent of a chart (for think-cell's Same Scale): the range the
 * auto scale would cover. Null for charts without a value axis (100%, Mekko,
 * butterfly, scatter, gantt).
 *
 * This reports what the layout DRAWS, not what the datasheet holds. Error and
 * Target rows are chart anatomy rather than data — counting an Error row's
 * magnitude as a data point gave a range that was neither — so they are pulled
 * out first and then folded back in as the whiskers and ticks they become.
 * Same Scale writes this straight back as a hard `scale` override, which
 * suppresses the auto-widen below, so anything this under-reports renders off
 * the shape.
 */
export function valueExtent(cfg: ChartConfig): { min: number; max: number } | null {
  const { cfg: clean, errors, targets } = extractErrorRows(cfg);
  return drawnExtent(clean, errors, targets);
}

/**
 * valueExtent for a config whose carried rows are already extracted — the state
 * buildChart is in by the time it needs the scale. Re-extracting there would
 * find nothing left to extract and silently widen by zero.
 */
function drawnExtent(
  clean: ChartConfig,
  errors: { plus: (number | null)[]; minus: (number | null)[] } | null,
  targets: (number | null)[] | null,
): { min: number; max: number } | null {
  const base = dataExtent(clean);
  return base ? widenForAnatomy(base, clean, errors, targets) : null;
}

/**
 * Grow an extent to cover the decorations drawn beyond the data: error whiskers,
 * target ticks and threshold lines. Deliberately conservative — whiskers are
 * anchored per category, but taking the largest delta against the extreme value
 * can only over-reach, and over-reaching costs a little white space where
 * under-reaching pushes ink off the shape.
 */
function widenForAnatomy(
  base: { min: number; max: number },
  cfg: ChartConfig,
  errors: { plus: (number | null)[]; minus: (number | null)[] } | null,
  targets: (number | null)[] | null,
): { min: number; max: number } {
  const nums = (xs: (number | null)[] | null | undefined) => (xs ?? []).filter((v): v is number => v != null);
  const maxPlus = Math.max(0, ...nums(errors?.plus));
  const maxMinus = Math.max(0, ...nums(errors?.minus));
  const targetVals = nums(targets);
  // Same normalization decor.ts applies to the legacy single valueLine. Only an
  // explicit threshold can sit outside the data — a "mean" line never can.
  const decor = cfg.decorations;
  const lineVals = (decor?.valueLines ?? (decor?.valueLine ? [decor.valueLine] : []))
    .filter((l): l is { mode: "value"; value: number } => l.mode === "value")
    .map((l) => l.value);
  return {
    min: Math.min(base.min - maxMinus, 0, ...targetVals, ...lineVals),
    max: Math.max(base.max + maxPlus, 0, ...targetVals, ...lineVals),
  };
}

export { layoutColumns, layoutWaterfall, layoutMekko, layoutLine };
