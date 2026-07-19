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
  /** Present when an axis break compresses part of the range: the band's y extent. */
  breakBand?: { yLow: number; yHigh: number };
}

/**
 * Linear scale mapping [min, max] onto the plot height, snapped to nice ticks.
 * `override` pins either end manually (think-cell's axis-handle dragging).
 */
export function valueScale(
  frame: Frame,
  dataMin: number,
  dataMax: number,
  override?: { min?: number; max?: number },
  axisBreak?: { from: number; to: number },
  logScale?: boolean,
): ValueScale {
  // Logarithmic axis: decade ticks; requires positive data (falls back otherwise).
  if (logScale && dataMax > 0) {
    const minPos = Math.max(dataMin > 0 ? dataMin : dataMax / 1000, 1e-12);
    const lo10 = Math.floor(Math.log10(override?.min && override.min > 0 ? override.min : minPos));
    // Never let the top decade fall below the bottom one: a manual scale.min set
    // above the data would give an empty tick list → undefined min → NaN toY for
    // the whole axis. Clamp to at least one decade.
    const hi10 = Math.max(lo10, Math.ceil(Math.log10(override?.max && override.max > 0 ? override.max : dataMax)));
    const ticks: number[] = [];
    for (let e = lo10; e <= hi10; e++) ticks.push(Math.pow(10, e));
    const min = ticks[0];
    const max = ticks[ticks.length - 1];
    const span = Math.log10(max) - Math.log10(min) || 1;
    const toY = (v: number) => frame.y + frame.h - ((Math.log10(Math.max(v, min)) - Math.log10(min)) / span) * frame.h;
    return { min, max, ticks, toY };
  }
  const lo = override?.min ?? Math.min(0, dataMin);
  const hi = override?.max ?? Math.max(0, dataMax);
  let ticks = niceTicks(lo, hi, 5).filter(
    (t) => (override?.min == null || t >= override.min - 1e-9) && (override?.max == null || t <= override.max + 1e-9),
  );
  const min = override?.min ?? ticks[0];
  const max = override?.max ?? ticks[ticks.length - 1];
  let toY = (v: number) => frame.y + frame.h - ((v - min) / (max - min || 1)) * frame.h;
  let breakBand: ValueScale["breakBand"];

  // think-cell axis break: the [from, to] range is compressed into a small
  // fixed band so out-of-scale columns fit the plot.
  if (axisBreak && axisBreak.from > min && axisBreak.to < max && axisBreak.to > axisBreak.from) {
    const { from, to } = axisBreak;
    const gapFrac = 0.06;
    const below = from - min;
    const above = max - to;
    const span = below + above || 1;
    const belowFrac = (below / span) * (1 - gapFrac);
    const aboveFrac = (above / span) * (1 - gapFrac);
    const frac = (v: number) =>
      v <= from
        ? (below ? (v - min) / below : 0) * belowFrac
        : v >= to
          ? belowFrac + gapFrac + (above ? (v - to) / above : 0) * aboveFrac
          : belowFrac + ((v - from) / (to - from)) * gapFrac;
    toY = (v: number) => frame.y + frame.h - frac(v) * frame.h;
    ticks = ticks.filter((t) => t <= from + 1e-9 || t >= to - 1e-9);
    breakBand = { yLow: toY(from), yHigh: toY(to) };
  }
  return { min, max, ticks, toY, breakBand };
}

/** Slanted-band break marker drawn across the plot (over the columns). */
export function breakMarkerNodes(frame: Frame, scale: ValueScale, style: ChartStyle): SceneNode[] {
  if (!scale.breakBand) return [];
  const { yLow, yHigh } = scale.breakBand;
  const skew = 2.5;
  return [
    {
      kind: "rect",
      x: frame.x - 2,
      y: yHigh,
      w: frame.w + 4,
      h: yLow - yHigh,
      fill: style.background,
      name: "axis-break",
    },
    {
      kind: "line",
      x1: frame.x - 4,
      y1: yLow + skew,
      x2: frame.x + frame.w + 4,
      y2: yLow - skew,
      stroke: style.mutedText,
      strokeWidth: 1,
      name: "axis-break-lo",
    },
    {
      kind: "line",
      x1: frame.x - 4,
      y1: yHigh + skew,
      x2: frame.x + frame.w + 4,
      y2: yHigh - skew,
      stroke: style.mutedText,
      strokeWidth: 1,
      name: "axis-break-hi",
    },
  ];
}

export interface FrameReservations {
  titleH: number;
  totalsH: number;
  categoryAxisH: number;
  valueAxisW: number;
  seriesLabelsW: number;
}

/** Height reserved above the plot for the chart title (0 when untitled). */
export function titleHeight(cfg: ChartConfig, style: ChartStyle): number {
  return cfg.title ? style.fontSize * 1.6 + 6 : 0;
}

/**
 * The chart-title text node, or null when the chart has no title. Every layout
 * that draws its own title emitted this exact node inline; sharing it keeps the
 * title's size/weight/placement from drifting between chart kinds.
 */
export function titleNode(cfg: ChartConfig, style: ChartStyle): SceneNode | null {
  if (!cfg.title) return null;
  const fs = style.fontSize;
  return {
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
  };
}

/** Height reserved at the bottom for the footnote / "100% =" line. */
export function footnoteH(cfg: ChartConfig, style: ChartStyle, decor: Decorations): number {
  // scatter.spread prints its cap on the footnote line, so it needs the row
  // reserved even when the author wrote no footnote of their own.
  return cfg.footnote || decor.hundredPercentNote || cfg.scatter?.spread ? style.fontSize * 1.3 : 0;
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
  const titleH = titleHeight(cfg, style);
  const totalsH = decor.totals || decor.cagr || decor.difference || decor.grandTotal ? fs * 1.5 + 4 : fs * 0.8;
  const categoryAxisH = decor.categoryAxis ? fs * 1.5 + 3 : 4;
  const valueAxisW = decor.valueAxis ? 34 : 2;
  const seriesLabelsW = decor.seriesLabels
    ? Math.min(cfg.width * 0.3, Math.max(0, ...seriesNames.map((s) => textWidth(s, fs))) + 14)
    : 2;
  // Extra headroom when a difference arrow is drawn past the last column.
  const diffW = decor.difference ? 26 : 0;
  const frame: Frame = {
    x: valueAxisW,
    y: titleH + totalsH,
    w: cfg.width - valueAxisW - Math.max(seriesLabelsW, diffW + 2) - 2,
    h: cfg.height - titleH - totalsH - categoryAxisH - footnoteH(cfg, style, decor),
  };
  return { frame, res: { titleH, totalsH, categoryAxisH, valueAxisW, seriesLabelsW } };
}

/**
 * Plot rectangle for horizontal (bar) orientation: category labels on the
 * left, value axis at the bottom, totals to the right of the bar ends,
 * series legend row at the top.
 */
export function computeFrameHorizontal(cfg: ChartConfig, style: ChartStyle, decor: Decorations): Frame {
  const fs = style.fontSize;
  const titleH = titleHeight(cfg, style);
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
    h: cfg.height - titleH - legendH - valueAxisH - footnoteH(cfg, style, decor),
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
  const t = titleNode(cfg, style);
  if (t) nodes.push(t);
  if (decor.gridlines && scale) {
    for (const t of scale.ticks) {
      if (t === 0) continue;
      const y = scale.toY(t);
      nodes.push({
        kind: "line",
        x1: frame.x,
        y1: y,
        x2: frame.x + frame.w,
        y2: y,
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: "gridline",
      });
    }
  }
  if (decor.valueAxis && scale) {
    // "datamarks": Tufte-style tick dashes + labels, no axis line, no
    // gridlines; tickMode "data" places them at the data extremes instead
    // of nice round values (the range frame).
    const marks = decor.valueAxis === "datamarks";
    const ticks = marks && decor.tickMode === "data" ? [...new Set([scale.min, scale.max])] : scale.ticks;
    const axisFmt = resolveFormat(ticks, cfg.numberFormat);
    for (const t of ticks) {
      const y = scale.toY(t);
      if (marks) {
        nodes.push({
          kind: "line",
          x1: frame.x - 4,
          y1: y,
          x2: frame.x,
          y2: y,
          stroke: style.axis,
          strokeWidth: 1,
          name: "datamark",
        });
      }
      nodes.push({
        kind: "text",
        x: 0,
        y: y - fs * 0.7,
        w: marks ? frame.x - 6 : frame.x - 4,
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
  if (cfg.valueAxisTitle) {
    nodes.push({
      kind: "text",
      x: 0,
      y: Math.max(0, frame.y - fs * 1.5),
      w: Math.max(frame.x - 4, textWidth(cfg.valueAxisTitle, fs)),
      h: fs * 1.4,
      text: cfg.valueAxisTitle,
      fontSize: fs * 0.95,
      color: style.mutedText,
      align: "left",
      valign: "bottom",
      name: "value-axis-title",
    });
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
  return {
    kind: "line",
    x1: frame.x,
    y1: y,
    x2: frame.x + frame.w,
    y2: y,
    stroke: style.axis,
    strokeWidth: 1,
    name: "baseline",
  };
}
