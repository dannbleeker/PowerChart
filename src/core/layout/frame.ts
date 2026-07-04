import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { niceTicks, formatNumber, resolveFormat } from "../format";

export interface Frame {
  /** Plot rectangle in chart coordinates. */
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface CategorySlots {
  /** Center x of each slot. */
  centers: number[];
  /** Width of each slot (before the gap is removed). */
  slotWidth: number;
  /** Width of the column/cluster within a slot. */
  colWidth: number;
}

/** think-cell default: column width : gap ≈ 2 : 1 within each slot. */
export function categorySlots(frame: Frame, n: number, gapRatio = 1 / 3): CategorySlots {
  const slotWidth = frame.w / Math.max(1, n);
  const colWidth = slotWidth * (1 - gapRatio);
  const centers = Array.from({ length: n }, (_, i) => frame.x + slotWidth * (i + 0.5));
  return { centers, slotWidth, colWidth };
}

export interface ValueScale {
  min: number;
  max: number;
  ticks: number[];
  toY: (v: number) => number;
}

/** Linear scale mapping [min, max] onto the plot height, snapped to nice ticks. */
export function valueScale(frame: Frame, dataMin: number, dataMax: number): ValueScale {
  const lo = Math.min(0, dataMin);
  const hi = Math.max(0, dataMax);
  const ticks = niceTicks(lo, hi, 5);
  const min = ticks[0];
  const max = ticks[ticks.length - 1];
  const toY = (v: number) => frame.y + frame.h - ((v - min) / (max - min || 1)) * frame.h;
  return { min, max, ticks, toY };
}

export interface FrameReservations {
  titleH: number;
  totalsH: number;
  categoryAxisH: number;
  valueAxisW: number;
  seriesLabelsW: number;
}

/**
 * Compute the plot rectangle by reserving margins for the enabled decorations
 * (title, totals row, category labels, value axis, right-hand series labels).
 */
export function computeFrame(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  seriesNames: string[],
): { frame: Frame; res: FrameReservations } {
  const fs = style.fontSize;
  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const totalsH = decor.totals || decor.cagr || decor.difference ? fs * 1.5 + 4 : fs * 0.8;
  const categoryAxisH = decor.categoryAxis ? fs * 1.5 + 3 : 4;
  const valueAxisW = decor.valueAxis ? 34 : 2;
  const seriesLabelsW = decor.seriesLabels
    ? Math.min(
        cfg.width * 0.3,
        Math.max(0, ...seriesNames.map((s) => textWidth(s, fs))) + 14,
      )
    : 2;
  // Extra headroom when a difference arrow is drawn past the last column.
  const diffW = decor.difference ? 26 : 0;
  const frame: Frame = {
    x: valueAxisW,
    y: titleH + totalsH,
    w: cfg.width - valueAxisW - Math.max(seriesLabelsW, diffW + 2) - 2,
    h: cfg.height - titleH - totalsH - categoryAxisH,
  };
  return { frame, res: { titleH, totalsH, categoryAxisH, valueAxisW, seriesLabelsW } };
}

/**
 * Plot rectangle for horizontal (bar) orientation: category labels on the
 * left, value axis at the bottom, totals to the right of the bar ends,
 * series legend row at the top.
 */
export function computeFrameHorizontal(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
): Frame {
  const fs = style.fontSize;
  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const legendH = decor.seriesLabels && cfg.data.series.length > 1 ? fs * 1.6 + 4 : fs * 0.6;
  const catW = decor.categoryAxis
    ? Math.min(cfg.width * 0.3, Math.max(0, ...cfg.data.categories.map((c) => textWidth(c, fs))) + 8)
    : 2;
  const valueAxisH = decor.valueAxis ? fs * 1.6 : 4;
  const totalsW = decor.totals ? fs * 4 : fs * 0.8;
  return {
    x: catW,
    y: titleH + legendH,
    w: cfg.width - catW - totalsW - 2,
    h: cfg.height - titleH - legendH - valueAxisH,
  };
}

/** Title, category labels, value axis, gridlines — shared chrome for all cartesian charts. */
export function chromeNodes(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  frame: Frame,
  centers: number[],
  scale?: ValueScale,
): SceneNode[] {
  const nodes: SceneNode[] = [];
  const fs = style.fontSize;
  if (cfg.title) {
    nodes.push({
      kind: "text",
      x: 0,
      y: 0,
      w: cfg.width,
      h: fs * 1.6,
      text: cfg.title,
      fontSize: fs * 1.2,
      bold: true,
      color: style.text,
      align: "left",
      valign: "top",
      name: "title",
    });
  }
  if (decor.gridlines && scale) {
    for (const t of scale.ticks) {
      if (t === 0) continue;
      const y = scale.toY(t);
      nodes.push({ kind: "line", x1: frame.x, y1: y, x2: frame.x + frame.w, y2: y, stroke: style.gridline, strokeWidth: 0.75, name: "gridline" });
    }
  }
  if (decor.valueAxis && scale) {
    const axisFmt = resolveFormat(scale.ticks, cfg.numberFormat);
    for (const t of scale.ticks) {
      const y = scale.toY(t);
      nodes.push({
        kind: "text",
        x: 0,
        y: y - fs * 0.7,
        w: frame.x - 4,
        h: fs * 1.4,
        text: formatNumber(t, axisFmt),
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: "value-axis",
      });
    }
  }
  if (decor.categoryAxis) {
    const slotW = centers.length > 1 ? centers[1] - centers[0] : frame.w;
    cfg.data.categories.forEach((cat, i) => {
      nodes.push({
        kind: "text",
        x: centers[i] - slotW / 2,
        y: frame.y + frame.h + 3,
        w: slotW,
        h: fs * 1.4,
        text: cat,
        fontSize: fs,
        color: style.text,
        align: "center",
        valign: "top",
        name: `category-${i}`,
      });
    });
  }
  return nodes;
}

/** Baseline (zero line) — drawn on top of columns, think-cell style. */
export function baselineNode(frame: Frame, y: number, style: ChartStyle): SceneNode {
  return { kind: "line", x1: frame.x, y1: y, x2: frame.x + frame.w, y2: y, stroke: style.axis, strokeWidth: 1, name: "baseline" };
}
