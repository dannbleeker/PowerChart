import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { niceTicks, formatNumber, resolveAxisFormat } from "../format";

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
/**
 * The floor a LOG axis should be given.
 *
 * Every caller derives `dataMin` with a zero seed (`minOf(all, 0)`) because a
 * linear value axis baselines at zero — which means valueScale's `dataMin > 0`
 * branch is unreachable from real code and EVERY log axis fell back to
 * `dataMax / 1000`: three decades below the top, whatever the data does. Values
 * of 200–300 got the axis 0.1…1,000 and drew inside the top 4% of the plot,
 * with two decades of gridlines under them that no datum can reach.
 *
 * Return the smallest POSITIVE value instead, or `dataMin` when there is none
 * (valueScale then declines the log branch anyway, since dataMax ≤ 0).
 */
export const logFloor = (values: readonly number[], dataMin: number): number => {
  const lo = values.reduce((m, v) => (v > 0 && v < m ? v : m), Infinity);
  return Number.isFinite(lo) ? lo : dataMin;
};

export function valueScale(
  frame: Frame,
  dataMin: number,
  dataMax: number,
  override?: { min?: number; max?: number },
  axisBreak?: { from: number; to: number },
  logScale?: boolean,
  // Column charts baseline at zero, so their axis must always include it. But a
  // distribution chart (boxplot/violin/candlestick) of e.g. scores 40–95 forcing
  // 0 into the range squashes the data into the top of the plot — those callers
  // pass zeroFloor:false to keep the domain data-driven. cfg.scale still overrides.
  zeroFloor = true,
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
  const autoLo = zeroFloor ? Math.min(0, dataMin) : dataMin;
  const autoHi = zeroFloor ? Math.max(0, dataMax) : dataMax;
  const lo = override?.min ?? autoLo;
  const hi = override?.max ?? autoHi;
  let ticks = niceTicks(lo, hi, 5).filter(
    (t) => (override?.min == null || t >= override.min - 1e-9) && (override?.max == null || t <= override.max + 1e-9),
  );
  let min = override?.min ?? ticks[0];
  let max = override?.max ?? ticks[ticks.length - 1];
  // A manual scale the data cannot live in — scale.min at or above the auto max
  // (the filter then leaves a single tick, so max collapses onto min), or an
  // outright degenerate {min:0,max:0} — makes the `max - min || 1` fallback below
  // map ONE data unit to ONE point: a 300pt chart emitted its columns ~83 canvas
  // heights down the slide. normalizeConfig already repairs an inverted or
  // non-finite scale; repair an unusable one the same way, by falling back to the
  // range the data itself asks for.
  if (!(max > min)) {
    ticks = niceTicks(autoLo, autoHi, 5);
    min = ticks[0];
    max = ticks[ticks.length - 1];
  }
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
  /** Height of the IBCS variance tier below the plot (0 when off). */
  varianceH: number;
}

/**
 * Height reserved below the plot for the IBCS variance tier. Only the vertical
 * column family (stacked/clustered/100%, drawn by layoutColumns) actually paints
 * it, so gate the reservation on that — otherwise a line/area/waterfall/boxplot/
 * violin/candlestick chart with `decorations.variance` lost 4.5×fontSize of plot
 * to a strip that drew nothing.
 */
export function varianceBandHeight(cfg: ChartConfig, decor: Decorations, style: ChartStyle): number {
  const drawsTier =
    !cfg.horizontal && (cfg.kind === "stacked" || cfg.kind === "clustered" || cfg.kind === "stacked100");
  return decor.variance && drawsTier ? style.fontSize * 4.5 : 0;
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

/** One legend entry's placement: its top-left x and 0-based wrap row. */
export interface LegendSlot {
  x: number;
  row: number;
}

/**
 * The wrap walk shared by `legendRow` (which places the chips) and
 * `legendRowCount` (which reserves the vertical space they need). Extracting it
 * means the drawer and the reserver can never disagree on how many rows a legend
 * occupies. Chips march left-to-right; one that would cross `maxX` starts a new
 * row (never the first entry of a row, so a single over-wide label still draws).
 */
export function legendWrapWalk(labels: string[], fs: number, x0: number, maxX: number): LegendSlot[] {
  const chip = fs * 0.7;
  const slots: LegendSlot[] = [];
  let x = x0;
  let row = 0;
  for (const label of labels) {
    const wLabel = textWidth(label, fs);
    if (x > x0 && x + chip + 3 + wLabel > maxX) {
      x = x0;
      row++;
    }
    slots.push({ x, row });
    x += chip + 3 + wLabel + 12;
  }
  return slots;
}

/** Number of rows the wrapping legend occupies (0 when there are no labels). */
export function legendRowCount(labels: string[], fs: number, x0: number, maxX: number): number {
  const slots = legendWrapWalk(labels, fs, x0, maxX);
  return slots.length ? slots[slots.length - 1].row + 1 : 0;
}

/**
 * Labels for the default series legend (the one `horizontalChrome` and the
 * horizontal mekko draw): a scenario-qualified series name, and only when there
 * is more than one series to tell apart. The height-reserving frame code and the
 * chip-drawing `legendRow` both derive their labels here so they legend the same
 * set and their wrap walks land on the same row count.
 */
export function seriesLegendLabels(cfg: ChartConfig): string[] {
  if (cfg.data.series.length <= 1) return [];
  return cfg.data.series.map((s) => (s.scenario ? `${s.name} (${s.scenario})` : s.name));
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
  const varianceH = varianceBandHeight(cfg, decor, style);
  const frame: Frame = {
    x: valueAxisW,
    y: titleH + totalsH,
    w: cfg.width - valueAxisW - Math.max(seriesLabelsW, diffW + 2) - 2,
    h: cfg.height - titleH - totalsH - varianceH - categoryAxisH - footnoteH(cfg, style, decor),
  };
  return { frame, res: { titleH, totalsH, categoryAxisH, valueAxisW, seriesLabelsW, varianceH } };
}

/**
 * Plot rectangle for horizontal (bar) orientation: category labels on the
 * left, value axis at the bottom, totals to the right of the bar ends,
 * series legend row at the top.
 */
export function computeFrameHorizontal(cfg: ChartConfig, style: ChartStyle, decor: Decorations): Frame {
  const fs = style.fontSize;
  const titleH = titleHeight(cfg, style);
  const catW = decor.categoryAxis
    ? Math.min(cfg.width * 0.3, Math.max(0, ...cfg.data.categories.map((c) => textWidth(c, fs))) + 8)
    : 2;
  // Reserve one row per wrapped legend row, not a fixed single row: many/long
  // series names on a narrow chart make legendRow wrap (same walk), and each
  // extra row must push the plot down or the legend draws on top of the bars.
  // The walk runs at the legend's own x0 (frame.x === catW) and maxX
  // (cfg.width - 4), so this reservation and legendRow agree on the row count. A
  // one-row legend keeps the old fs*1.6+4 exactly, so snapshots stay identical.
  const legendLabels = decor.seriesLabels ? seriesLegendLabels(cfg) : [];
  const legendRows = legendLabels.length ? legendRowCount(legendLabels, fs, catW, cfg.width - 4) : 0;
  const legendH = legendRows > 0 ? legendRows * (fs * 1.6) + 4 : fs * 0.6;
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
    // Tick labels are read against each other, so their precision comes from the
    // tick step, not the tick magnitude (resolveAxisFormat) — otherwise a narrow
    // axis prints the same label at several heights.
    const axisFmt = resolveAxisFormat(ticks, cfg.numberFormat);
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
    // Sit the category axis below the IBCS variance tier's reserved band (0 off).
    const catY = frame.y + frame.h + varianceBandHeight(cfg, decor, style) + 3;
    cfg.data.categories.forEach((cat, i) => {
      nodes.push({
        kind: "text",
        x: centers[i] - slotW / 2,
        y: catY,
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
