import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors } from "./types";
import { textWidth, type SceneNode } from "./scene";
import { cagr, formatNumber, formatPercent } from "./format";

/**
 * think-cell's signature annotations, computed from layout anchors so they
 * work identically across chart types.
 */
export function decorationNodes(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  a: LayoutAnchors,
): SceneNode[] {
  const nodes: SceneNode[] = [];
  const fs = style.fontSize;

  // Every decoration here anchors to the category geometry (categoryX / columnTop
  // / seriesLevels). A chart that drew NO segments — a single-0 pie or sunburst,
  // say — has empty anchor arrays, so any anchored read comes back undefined and
  // the decoration would emit NaN geometry. There is nothing to annotate.
  if (!a.categoryX.length) return nodes;

  // --- CAGR arrow: diagonal arrow between two column tops with "+x.x% p.a." ---
  if (decor.cagr) {
    const { from, to } = clampPair(decor.cagr, a);
    // Per-series CAGR when requested (think-cell computes on totals by default).
    const si = decor.cagr.series;
    const sVals = si != null ? cfg.data.series[si]?.values : undefined;
    const vFrom = sVals ? (sVals[from] ?? 0) : a.columnValue[from];
    const vTo = sVals ? (sVals[to] ?? 0) : a.columnValue[to];
    const rate = cagr(vFrom, vTo, to - from);
    // Clear the column totals row and difference arrows when shown.
    const lift = fs * 1.6 + (decor.totals ? fs * 1.5 : 0) + (decor.difference ? fs * 1.2 : 0);
    const x1 = a.categoryX[from];
    const y1 = a.columnTop[from] - lift;
    const x2 = a.categoryX[to];
    const y2 = a.columnTop[to] - lift;
    const angle = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
    nodes.push(
      { kind: "line", x1, y1, x2, y2, stroke: style.text, strokeWidth: 1.25, name: "cagr-line" },
      { kind: "arrowhead", x: x2, y: y2, angle, size: 5, fill: style.text, name: "cagr-head" },
      {
        kind: "text",
        x: (x1 + x2) / 2 - 45,
        y: Math.min(y1, y2) - fs * 1.6,
        w: 90,
        h: fs * 1.4,
        text: rate == null ? "CAGR n/a" : `${formatPercent(rate, 1, true)} p.a.`,
        fontSize: fs,
        bold: true,
        color: style.text,
        align: "center",
        valign: "bottom",
        name: "cagr-label",
      },
    );
  }

  // --- Difference arrow: dashed level line + vertical arrow ---
  // Total (column totals) by default; a level difference arrow when `series`
  // names a series — it compares the cumulative stack level at that series.
  if (decor.difference) {
    const { from, to } = clampPair(decor.difference, a);
    const si = decor.difference.series;
    const useLevel =
      si != null && a.seriesLevels != null && a.valueToY != null && si >= 0 && si < (a.seriesLevels[0]?.length ?? 0);
    let vFrom = useLevel ? a.seriesLevels![from][si!] : a.columnValue[from];
    let yFrom = useLevel ? a.valueToY!(vFrom) : a.columnTop[from];
    const vTo = useLevel ? a.seriesLevels![to][si!] : a.columnValue[to];
    const yTo = useLevel ? a.valueToY!(vTo) : a.columnTop[to];
    // Anchor the arrow's start at a value line instead of a column.
    const vlIdx = decor.difference.fromValueLine;
    const vls = decor.valueLines ?? (decor.valueLine ? [decor.valueLine] : []);
    if (vlIdx != null && vls[vlIdx] && a.valueToY) {
      const vl = vls[vlIdx];
      vFrom = vl.mode === "mean" ? a.columnValue.reduce((s, v) => s + v, 0) / (a.columnValue.length || 1) : vl.value;
      yFrom = a.valueToY(vFrom);
    }
    const x = a.categoryX[to] + a.categoryWidth[to] / 2 + 10;
    nodes.push(
      {
        kind: "line",
        x1: a.categoryX[from] + a.categoryWidth[from] / 2 + 2,
        y1: yFrom,
        x2: x + 4,
        y2: yFrom,
        stroke: style.mutedText,
        strokeWidth: 0.75,
        dash: [2, 2],
        name: "diff-level",
      },
      { kind: "line", x1: x, y1: yFrom, x2: x, y2: yTo, stroke: style.text, strokeWidth: 1.25, name: "diff-line" },
      { kind: "arrowhead", x, y: yTo, angle: yTo < yFrom ? -90 : 90, size: 5, fill: style.text, name: "diff-head" },
    );
    const usePct = decor.difference.percent ?? true;
    const label =
      usePct && vFrom !== 0
        ? formatPercent(vTo / vFrom - 1, 0, true)
        : formatNumber(vTo - vFrom, { ...cfg.numberFormat, forceSign: true });
    nodes.push({
      kind: "text",
      x: x + 3,
      y: (yFrom + yTo) / 2 - fs * 0.75,
      w: Math.max(30, cfg.width - x - 3),
      h: fs * 1.5,
      text: label,
      fontSize: fs,
      bold: true,
      color: style.text,
      align: "left",
      valign: "middle",
      name: "diff-label",
    });
  }

  // --- Value lines: dashed horizontals at fixed values or the mean of totals ---
  const valueLines = decor.valueLines ?? (decor.valueLine ? [decor.valueLine] : []);
  if (valueLines.length && a.valueToY) {
    valueLines.forEach((vl, i) => {
      const value =
        vl.mode === "mean" ? a.columnValue.reduce((s, v) => s + v, 0) / (a.columnValue.length || 1) : vl.value;
      const y = a.valueToY!(value);
      nodes.push(
        {
          kind: "line",
          x1: a.plot.x,
          y1: y,
          x2: a.plot.x + a.plot.w,
          y2: y,
          stroke: style.mutedText,
          strokeWidth: 1,
          dash: [3, 2],
          name: `value-line-${i}`,
        },
        {
          kind: "text",
          x: a.plot.x + 2,
          y: y - fs * 1.5,
          w: 100,
          h: fs * 1.4,
          text: (vl.mode === "mean" ? "Ø " : "") + formatNumber(value, cfg.numberFormat),
          fontSize: fs * 0.95,
          color: style.mutedText,
          align: "left",
          valign: "bottom",
          name: `value-line-label-${i}`,
        },
      );
    });
  }

  // --- Speech-bubble callouts: a comment anchored to a column or level ---
  decor.callouts?.forEach((co, i) => {
    const c = Math.max(0, Math.min(a.categoryX.length - 1, co.category));
    const ax = a.categoryX[c];
    const useLevel =
      co.series != null &&
      a.seriesLevels != null &&
      a.valueToY != null &&
      co.series >= 0 &&
      co.series < (a.seriesLevels[c]?.length ?? 0);
    const ay = useLevel ? a.valueToY!(a.seriesLevels![c][co.series!]) : a.columnTop[c];
    const w = textWidth(co.text, fs) + fs * 1.2;
    const h = fs * 1.9;
    // Bubble center defaults to hovering above the anchor.
    const bx = ax + (co.dx ?? 0);
    const by = ay - fs * 4.2 + (co.dy ?? 0);
    nodes.push(
      {
        kind: "line",
        x1: bx,
        y1: by + h / 2 - 1,
        x2: ax,
        y2: ay - 2,
        stroke: style.text,
        strokeWidth: 0.75,
        name: `callout-tail-${i}`,
      },
      {
        kind: "rect",
        x: bx - w / 2,
        y: by - h / 2,
        w,
        h,
        fill: style.background,
        stroke: style.text,
        strokeWidth: 1,
        name: `callout-box-${i}`,
      },
      {
        kind: "text",
        x: bx - w / 2,
        y: by - h / 2,
        w,
        h,
        text: co.text,
        fontSize: fs,
        color: style.text,
        align: "center",
        valign: "middle",
        name: `callout-text-${i}`,
      },
    );
  });

  return nodes;
}

/**
 * Shaded background bands highlighting an axis region — drawn BEHIND the
 * data (the caller prepends these to the scene). axis "y" spans a value
 * range; axis "x" spans category indices.
 */
export function bandNodes(cfg: ChartConfig, style: ChartStyle, decor: Decorations, a: LayoutAnchors): SceneNode[] {
  const nodes: SceneNode[] = [];
  const fs = style.fontSize;
  decor.bands?.forEach((band, i) => {
    const fill = band.color ?? "#f2f1ec";
    let r: { x: number; y: number; w: number; h: number } | null = null;
    if (band.axis === "y" && a.valueToY) {
      // Clip to the plot: band.from/to are data values that may fall outside the
      // value domain, and valueToY extrapolates past the axis, so an unclamped
      // band renders off-frame. (The x-branch below already clamps its indices;
      // a band entirely outside the plot collapses to h<=0 and is dropped below.)
      const y1 = a.valueToY(band.from);
      const y2 = a.valueToY(band.to);
      const top = Math.max(a.plot.y, Math.min(y1, y2));
      const bot = Math.min(a.plot.y + a.plot.h, Math.max(y1, y2));
      r = { x: a.plot.x, y: top, w: a.plot.w, h: bot - top };
    } else if (band.axis === "x" && a.categoryX.length) {
      const c1 = Math.max(0, Math.min(a.categoryX.length - 1, Math.min(band.from, band.to)));
      const c2 = Math.max(0, Math.min(a.categoryX.length - 1, Math.max(band.from, band.to)));
      const x1 = a.categoryX[c1] - a.categoryWidth[c1] * 0.75;
      const x2 = a.categoryX[c2] + a.categoryWidth[c2] * 0.75;
      r = { x: x1, y: a.plot.y, w: x2 - x1, h: a.plot.h };
    }
    if (!r || r.w <= 0 || r.h <= 0) return;
    nodes.push({ kind: "rect", ...r, fill, name: `band-${i}` });
    if (band.label) {
      nodes.push({
        kind: "text",
        x: r.x + 3,
        y: r.y + 1,
        w: Math.max(20, r.w - 6),
        h: fs * 1.3,
        text: band.label,
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: "left",
        valign: "top",
        name: `band-label-${i}`,
      });
    }
  });
  return nodes;
}

function clampPair(p: { from: number; to: number }, a: LayoutAnchors): { from: number; to: number } {
  const n = a.categoryX.length;
  const from = Math.max(0, Math.min(n - 1, p.from));
  const to = Math.max(0, Math.min(n - 1, p.to));
  return from <= to ? { from, to } : { from: to, to: from };
}
