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
import { decorationNodes } from "./decor";
import { resolveLabelCollisions } from "./collide";
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

/** Build a renderer-agnostic scene from a chart config. Pure and synchronous. */
export function buildChart(rawCfg: ChartConfig): Scene {
  const cfg = sortCategories(rawCfg);
  const style: ChartStyle = { ...DEFAULT_STYLE, ...cfg.style };
  const decor: Decorations = { ...DEFAULT_DECOR, ...cfg.decorations };

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
    default:
      result = layoutColumns(cfg, style, decor);
  }

  // Decorations assume a vertical value axis; skip them for bar orientation
  // and butterfly charts.
  const skipDecor =
    cfg.horizontal ||
    ["butterfly", "scatter", "bubble", "gantt", "pie", "doughnut"].includes(cfg.kind);
  const nodes = skipDecor
    ? result.nodes
    : [...result.nodes, ...decorationNodes(cfg, style, decor, result.anchors)];

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
