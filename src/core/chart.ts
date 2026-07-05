import type { ChartConfig, ChartKind, ChartStyle, Decorations } from "./types";
import type { Scene } from "./scene";
import { DEFAULT_DECOR, DEFAULT_STYLE } from "./style";
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
  const totals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + (s.values[c] ?? 0), 0),
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

/** Datasheet rows carrying error-bar deltas: Error (±), Error+ / Error−. */
const ERROR_ROW = /^error\s*([+\-−])?$/i;
const ERROR_KINDS: ChartKind[] = ["stacked", "clustered", "line", "area"];

/**
 * Pull Error rows out of the data (so they don't render as segments) and
 * return per-category plus/minus deltas. Bars anchor at the column total
 * (single-series charts: the value) or the first line series.
 */
function extractErrorRows(cfg: ChartConfig): {
  cfg: ChartConfig;
  errors: { plus: (number | null)[]; minus: (number | null)[] } | null;
} {
  if (!ERROR_KINDS.includes(cfg.kind)) return { cfg, errors: null };
  const rows = cfg.data.series.filter((s) => ERROR_ROW.test(s.name.trim()));
  if (!rows.length) return { cfg, errors: null };
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
    cfg: { ...cfg, data: { ...cfg.data, series: cfg.data.series.filter((s) => !ERROR_ROW.test(s.name.trim())) } },
    errors: { plus: pick("+"), minus: pick("-") },
  };
}

/** Build a renderer-agnostic scene from a chart config. Pure and synchronous. */
export function buildChart(rawCfg: ChartConfig): Scene {
  const extracted = extractErrorRows(sortCategories(rawCfg));
  let cfg = extracted.cfg;
  const errors = extracted.errors;
  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const decor: Decorations = { ...DEFAULT_DECOR, ...cfg.decorations };

  // Widen the auto scale so error bars stay inside the plot.
  if (errors && !cfg.horizontal && cfg.scale?.max == null) {
    const ext = valueExtent(cfg);
    if (ext) {
      const maxPlus = Math.max(0, ...errors.plus.filter((v): v is number => v != null));
      const maxMinus = Math.max(0, ...errors.minus.filter((v): v is number => v != null));
      const ticks = niceTicks(Math.min(ext.min - maxMinus, 0), ext.max + maxPlus, 5);
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
    default:
      result = layoutColumns(cfg, style, decor);
  }

  // Decorations assume a vertical value axis; skip them for bar orientation
  // and butterfly charts.
  const skipDecor =
    cfg.horizontal ||
    ["butterfly", "scatter", "bubble", "gantt", "pie", "doughnut", "radar", "heatmap", "tilemap"].includes(cfg.kind);
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
