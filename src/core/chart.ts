import type { ChartConfig, ChartKind, ChartStyle, Decorations } from "./types";
import type { Scene } from "./scene";
import { DEFAULT_DECOR, DEFAULT_STYLE, seriesColor } from "./style";
import { layoutColumns, layoutCombo } from "./layout/column";
import { layoutWaterfall } from "./layout/waterfall";
import { layoutMekko } from "./layout/mekko";
import { layoutLine } from "./layout/line";
import { layoutButterfly } from "./layout/butterfly";
import { layoutScatter } from "./layout/scatter";
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
import { bandNodes, decorationNodes } from "./decor";
import { resolveLabelCollisions } from "./collide";
import { formatNumber, niceTicks, resolveFormat } from "./format";
import type { SceneNode } from "./scene";
import type { LayoutResult } from "./layout/column";

export const DEFAULT_SIZE = { width: 480, height: 300 };

const SORTABLE: ChartKind[] = ["stacked", "clustered", "stacked100", "mekko", "pie", "doughnut", "butterfly"];

/** Reorder categories (and every per-category array) by column total. */
function sortCategories(cfg: ChartConfig): ChartConfig {
  if (!cfg.categorySort || !SORTABLE.includes(cfg.kind)) return cfg;
  const { data } = cfg;
  // Sort runs before error/target rows are extracted; those carried rows are
  // not real stack contributors, so exclude them from the ranking totals.
  const ranked = data.series.filter((s) => !CARRIED_ROW.test(s.name.trim()));
  const totals = data.categories.map((_, c) =>
    ranked.reduce((a, s) => a + (s.values[c] ?? 0), 0),
  );
  const sign = cfg.categorySort === "ascending" ? 1 : -1;
  const order = data.categories.map((_, c) => c).sort((a, b) => sign * (totals[a] - totals[b]));
  const pick = <T,>(arr: T[] | undefined) => (arr ? order.map((c) => arr[c]) : undefined);
  return {
    ...cfg,
    data: {
      ...data,
      categories: order.map((c) => data.categories[c]),
      series: data.series.map((s) => ({ ...s, values: order.map((c) => s.values[c]) })),
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
  const bar = cfg.data.series.find((s) => s.type !== "line");
  if (!bar) return cfg;
  const order = cfg.data.categories.map((_, c) => c).sort((a, b) => (bar.values[b] ?? 0) - (bar.values[a] ?? 0));
  const pick = <T,>(arr: T[] | undefined) => (arr ? order.map((c) => arr[c]) : undefined);
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
      series: [{ ...bar, values: barVals }, { name: "Cumulative %", type: "line", values: cum }],
      hundredPercent: pick(cfg.data.hundredPercent),
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
        n.cx += dx;
        n.cy += dy;
        break;
      case "polygon":
        n.points = n.points.map((p) => ({ x: p.x + dx, y: p.y + dy }));
        break;
    }
    return n;
  });
}

/**
 * Small multiples: one single-series panel per data series, titled by the
 * series name, laid out in a grid and pinned to one shared value scale so
 * panels compare honestly. Null when the config doesn't call for it.
 */
function buildMultiples(cfg: ChartConfig): Scene | null {
  if (!cfg.multiples || !MULTIPLES_KINDS.includes(cfg.kind)) return null;
  const carried = cfg.data.series.filter((s) => CARRIED_ROW.test(s.name.trim()));
  const dataSeries = cfg.data.series.filter((s) => !CARRIED_ROW.test(s.name.trim()));
  if (dataSeries.length < 2) return null;

  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const fs = style.fontSize;
  const n = dataSeries.length;
  const cols = Math.max(1, Math.min(n, cfg.multiples.columns ?? (n <= 3 ? n : Math.ceil(Math.sqrt(n)))));
  const rows = Math.ceil(n / cols);
  const gap = 10;
  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const footH = cfg.footnote ? fs * 1.3 : 0;
  const panelW = (cfg.width - gap * (cols - 1)) / cols;
  const panelH = (cfg.height - titleH - footH - gap * (rows - 1)) / rows;

  const panelCfg = (s: (typeof dataSeries)[number], si: number): ChartConfig => ({
    ...cfg,
    multiples: undefined,
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
      for (const t of targets) if (t != null) {
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
      : { min: 0, max: Math.max(1, ...dataSeries.flatMap((s) => s.values.filter((v): v is number => v != null))) };
    scale = { min: scale?.min ?? global.min, max: scale?.max ?? global.max };
  }

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }
  dataSeries.forEach((s, si) => {
    const panel = buildChart({ ...panelCfg(s, si), scale });
    const dx = (si % cols) * (panelW + gap);
    const dy = titleH + Math.floor(si / cols) * (panelH + gap);
    nodes.push(...translateNodes(panel.nodes, dx, dy, `p${si}-`));
  });
  if (cfg.footnote) {
    nodes.push({
      kind: "text", x: 2, y: cfg.height - fs * 1.15, w: cfg.width - 4, h: fs * 1.1,
      text: cfg.footnote, fontSize: fs * 0.85, color: style.mutedText,
      align: "left", valign: "bottom", name: "footnote",
    });
  }
  return { width: cfg.width, height: cfg.height, nodes };
}

/** Build a renderer-agnostic scene from a chart config. Pure and synchronous. */
export function buildChart(rawCfg: ChartConfig): Scene {
  const multiples = buildMultiples(rawCfg);
  if (multiples) return multiples;
  const extracted = extractErrorRows(sortCategories(applyPareto(rawCfg)));
  let cfg = collapseOther(extracted.cfg);
  const errors = extracted.errors;
  const targets = extracted.targets;
  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const decor: Decorations = { ...DEFAULT_DECOR, ...cfg.decorations };

  // Widen the auto scale so error bars and target ticks stay inside the plot.
  if ((errors || targets) && !cfg.horizontal && cfg.scale?.max == null) {
    const ext = valueExtent(cfg);
    if (ext) {
      const maxPlus = Math.max(0, ...(errors?.plus ?? []).filter((v): v is number => v != null));
      const maxMinus = Math.max(0, ...(errors?.minus ?? []).filter((v): v is number => v != null));
      const maxTarget = Math.max(0, ...(targets ?? []).filter((v): v is number => v != null));
      const ticks = niceTicks(Math.min(ext.min - maxMinus, 0), Math.max(ext.max + maxPlus, maxTarget), 5);
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
    ["butterfly", "scatter", "bubble", "gantt", "pie", "doughnut", "radar", "heatmap", "tilemap", "cascade", "funnel", "waffle"].includes(cfg.kind);
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
      nodes.push({ kind: "line", x1: x, y1: yHi, x2: x, y2: yLo, stroke: style.axis, strokeWidth: 1, name: `error-${c}` });
      if (plus != null) nodes.push({ kind: "line", x1: x - capW / 2, y1: yHi, x2: x + capW / 2, y2: yHi, stroke: style.axis, strokeWidth: 1, name: `error-cap-hi-${c}` });
      if (minus != null) nodes.push({ kind: "line", x1: x - capW / 2, y1: yLo, x2: x + capW / 2, y2: yLo, stroke: style.axis, strokeWidth: 1, name: `error-cap-lo-${c}` });
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
      nodes.push({ kind: "line", x1: x - half, y1: y, x2: x + half, y2: y, stroke: style.text, strokeWidth: 2.25, name: `target-${c}` });
      // Budget-vs-actual bridge: on a waterfall, a hatched gap segment shows
      // the distance from the achieved level to the target.
      if (cfg.kind === "waterfall") {
        const actual = a.columnValue[c];
        const gap = t - actual;
        if (Math.abs(gap) > 1e-9) {
          const yA = a.valueToY!(actual);
          const w = a.categoryWidth[c];
          nodes.push({
            kind: "rect", x: x - w / 2, y: Math.min(y, yA), w, h: Math.abs(yA - y),
            fill: style.neutral, pattern: "diagonal", stroke: style.mutedText, strokeWidth: 0.75,
            name: `target-gap-${c}`,
          });
          const fs = style.fontSize;
          nodes.push({
            kind: "text", x: x - w / 2 - 4, y: Math.min(y, yA) - fs * 1.5, w: w + 8, h: fs * 1.4,
            text: `Gap ${formatNumber(gap, { ...resolveFormat([t, actual], cfg.numberFormat), forceSign: true })}`,
            fontSize: fs * 0.95, bold: true, color: style.text,
            align: "center", valign: "bottom", name: `target-gap-label-${c}`,
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
  return { width: cfg.width, height: cfg.height, nodes };
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
      return d != null && d > 0
        ? d
        : data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0);
    });
    if (!denominators.length || denominators[0] <= 0) return null;
    return denominators.every((d) => Math.abs(d - denominators[0]) < 1e-9) ? denominators[0] : null;
  }
  return null;
}

/**
 * Value-axis extent of a chart's data (for think-cell's Same Scale): the
 * range the auto scale would cover. Null for charts without a value axis
 * (100%, Mekko, butterfly, scatter, gantt).
 */
export function valueExtent(cfg: ChartConfig): { min: number; max: number } | null {
  const { data, kind } = cfg;
  const cats = data.categories.map((_, c) => c);
  const vals = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  if (!vals.length) return null;
  switch (kind) {
    case "stacked": {
      const pos = cats.map((c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
      const neg = cats.map((c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
      return { min: Math.min(0, ...neg), max: Math.max(0, ...pos) };
    }
    case "clustered":
    case "line":
      return { min: Math.min(0, ...vals), max: Math.max(0, ...vals) };
    case "boxplot":
      return boxplotExtent(cfg);
    case "area": {
      const pos = cats.map((c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
      return { min: 0, max: Math.max(0, ...pos) };
    }
    case "waterfall": {
      const totals = new Set(cfg.waterfall?.totalIndices ?? []);
      let running = 0;
      let min = 0;
      let max = 0;
      cats.forEach((c) => {
        if (!totals.has(c)) running += data.series[0]?.values[c] ?? 0;
        min = Math.min(min, running);
        max = Math.max(max, running);
      });
      return { min, max };
    }
    default:
      return null;
  }
}

export { layoutColumns, layoutWaterfall, layoutMekko, layoutLine };
