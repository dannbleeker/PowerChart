import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, formatPercent, niceTicks, resolveFormat, segmentLabel } from "../format";
import { seriesColor } from "../style";
import {
  baselineNode,
  breakMarkerNodes,
  categorySlots,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  valueScale,
  type Frame,
  type ValueScale,
} from "./frame";

export interface LayoutResult {
  nodes: SceneNode[];
  anchors: LayoutAnchors;
}

/** Minimum segment thickness (relative to font size) before its label is hidden. */
const LABEL_FIT = 1.25;

/**
 * Stacked / clustered / 100% column charts — and, following think-cell's
 * "a bar chart is a rotated column chart" model, the same layouts in
 * horizontal orientation when cfg.horizontal is set.
 */
export function layoutColumns(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data, kind } = cfg;
  const n = data.categories.length;
  const stacked = kind !== "clustered";
  const pct = kind === "stacked100";
  const H = !!cfg.horizontal;

  const frame = H
    ? computeFrameHorizontal(cfg, style, decor)
    : computeFrame(cfg, style, decor, decor.seriesLabels ? data.series.map((s) => s.name) : []).frame;
  const fs = style.fontSize;

  // Category slots run along x (vertical) or y (horizontal).
  const catStart = H ? frame.y : frame.x;
  const catLen = H ? frame.h : frame.w;
  const slotLen = catLen / Math.max(1, n);
  const colThick = slotLen * (2 / 3);
  const centers = Array.from({ length: n }, (_, i) => catStart + slotLen * (i + 0.5));

  const posTotals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0),
  );
  const negTotals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0),
  );
  const signedTotals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + (s.values[c] ?? 0), 0),
  );
  // Per-category denominator for 100% charts (think-cell's "100%=" row).
  const denominators = data.categories.map((_, c) => {
    const d = data.hundredPercent?.[c];
    return d != null && d > 0 ? d : posTotals[c];
  });
  const fmt = resolveFormat(
    [...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)), ...signedTotals],
    cfg.numberFormat,
  );

  let dataMin: number, dataMax: number;
  if (pct) {
    dataMin = 0;
    dataMax = 1;
  } else if (stacked) {
    dataMin = Math.min(0, ...negTotals);
    dataMax = Math.max(0, ...posTotals);
  } else {
    const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
    dataMin = Math.min(0, ...all);
    dataMax = Math.max(0, ...all);
  }
  const scale: ValueScale = pct
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1], toY: (v: number) => frame.y + frame.h - v * frame.h }
    : valueScale(frame, dataMin, dataMax, cfg.scale, H ? undefined : cfg.axisBreak, !stacked && !H && cfg.logScale);

  // Value coordinate: distance along the value axis from the scale minimum.
  // Vertical charts route through toY so axis breaks apply; horizontal stays linear.
  const valLen = H ? frame.w : frame.h;
  const qOf = H
    ? (v: number) => ((v - scale.min) / (scale.max - scale.min || 1)) * valLen
    : (v: number) => frame.y + frame.h - scale.toY(v);
  /** Rect spanning [v0, v1] on the value axis at category position/thickness. */
  const segRect = (catPos: number, thick: number, v0: number, v1: number) => {
    const q0 = Math.min(qOf(v0), qOf(v1));
    const q1 = Math.max(qOf(v0), qOf(v1));
    return H
      ? { x: frame.x + q0, y: catPos - thick / 2, w: q1 - q0, h: thick }
      : { x: catPos - thick / 2, y: frame.y + frame.h - q1, w: thick, h: q1 - q0 };
  };

  const nodes: SceneNode[] = H
    ? horizontalChrome(cfg, style, decor, frame, centers, scale, qOf)
    : chromeNodes(cfg, style, decor, frame, centers, scale);
  const zeroQ = qOf(0);
  const y0 = H ? frame.x + zeroQ : frame.y + frame.h - zeroQ;
  const columnTop: number[] = [];
  const seriesLevels: number[][] = [];
  /** Segment mid-position of the last category per series, for series labels. */
  const lastSegMid: (number | null)[] = data.series.map(() => null);

  for (let c = 0; c < n; c++) {
    let up = 0; // running positive stack (value units)
    let down = 0;
    const levels: number[] = data.series.map(() => 0);
    const barThick = stacked ? colThick : colThick / Math.max(1, data.series.length);
    // think-cell's Segment Order: stacking order within this column.
    const order = data.series.map((_, i) => i);
    if (cfg.segmentOrder === "reverse") order.reverse();
    else if (cfg.segmentOrder === "ascending" || cfg.segmentOrder === "descending") {
      const sign = cfg.segmentOrder === "ascending" ? 1 : -1;
      order.sort((a, b) => sign * ((data.series[a].values[c] ?? 0) - (data.series[b].values[c] ?? 0)));
    }

    order.forEach((si, position) => {
      const s = data.series[si];
      const raw = s.values[c];
      let v = raw ?? 0;
      if (pct) v = denominators[c] > 0 ? Math.max(0, v) / denominators[c] : 0;
      let r: { x: number; y: number; w: number; h: number } | null = null;
      const fill = seriesColor(style, si, s.color);

      if (raw != null && v !== 0) {
        if (stacked) {
          if (v >= 0) {
            r = segRect(centers[c], colThick, up, up + v);
            up += v;
          } else {
            r = segRect(centers[c], colThick, down + v, down);
            down += v;
          }
        } else {
          const pos = centers[c] - colThick / 2 + (position + 0.5) * barThick;
          r = segRect(pos, barThick - 1, 0, v);
        }
      }
      levels[si] = up + down;
      if (!r) return;

      nodes.push({
        kind: "rect",
        ...r,
        fill,
        stroke: style.background,
        strokeWidth: stacked ? 0.75 : 0,
        name: `seg-${si}-${c}`,
      });
      if (c === n - 1) lastSegMid[si] = H ? r.x + r.w : r.y + r.h / 2;

      if (decor.segmentLabels) {
        // think-cell's label-content dropdown: value / % / series / category.
        const label = segmentLabel(decor.labelContent ?? (pct ? ["percent"] : ["value"]), {
          value: raw!,
          fraction: pct ? v : posTotals[c] > 0 ? Math.max(0, raw!) / posTotals[c] : null,
          series: s.name,
          category: data.categories[c],
          fmt,
        });
        const along = H ? r.w : r.h; // extent along the value axis
        const across = H ? r.h : r.w;
        const fits = H
          ? along >= textWidth(label, fs) + 2 && across >= fs * LABEL_FIT
          : along >= fs * LABEL_FIT && textWidth(label, fs) <= across + 2;
        if (fits) {
          nodes.push({
            kind: "text",
            x: r.x - 4,
            y: r.y + r.h / 2 - fs * 0.75,
            w: r.w + 8,
            h: fs * 1.5,
            text: label,
            fontSize: fs,
            color: contrastInk(fill),
            align: "center",
            valign: "middle",
            name: `label-${si}-${c}`,
          });
        }
      }
    });
    seriesLevels.push(levels);

    const topV = pct ? (denominators[c] > 0 ? posTotals[c] / denominators[c] : 0) : stacked ? up : Math.max(0, ...data.series.map((s) => s.values[c] ?? 0));
    const topQ = qOf(Math.max(0, topV));
    columnTop.push(H ? frame.x + topQ : frame.y + frame.h - topQ);

    if (decor.totals && !pct) {
      if (H) {
        nodes.push({
          kind: "text",
          x: frame.x + topQ + 3,
          y: centers[c] - fs * 0.75,
          w: cfg.width - (frame.x + topQ) - 3,
          h: fs * 1.5,
          text: formatNumber(signedTotals[c], fmt),
          fontSize: fs,
          bold: true,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `total-${c}`,
        });
      } else {
        nodes.push({
          kind: "text",
          x: centers[c] - slotLen / 2,
          y: frame.y + frame.h - topQ - fs * 1.45,
          w: slotLen,
          h: fs * 1.4,
          text: formatNumber(signedTotals[c], fmt),
          fontSize: fs,
          bold: true,
          color: style.text,
          align: "center",
          valign: "bottom",
          name: `total-${c}`,
        });
      }
    }
  }

  if (!H) nodes.push(...breakMarkerNodes(frame, scale, style));

  // Zero baseline: horizontal line (vertical charts) or vertical line (bars).
  if (H) {
    nodes.push({ kind: "line", x1: y0, y1: frame.y, x2: y0, y2: frame.y + frame.h, stroke: style.axis, strokeWidth: 1, name: "baseline" });
  } else {
    nodes.push(baselineNode(frame, y0, style));
  }

  if (decor.seriesLabels && !H) {
    nodes.push(...seriesLabelNodes(cfg, style, frame, lastSegMid as (number | null)[]));
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => colThick),
      columnTop,
      columnValue: pct ? posTotals : signedTotals,
      seriesLevels,
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: pct || H ? undefined : scale.toY,
    },
  };
}

/**
 * Combo chart, think-cell style: stacked columns plus line series drawn over
 * them on the same value axis. Series with `type: "line"` become lines; if
 * none is marked, the last series does.
 */
export function layoutCombo(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const marked = cfg.data.series.some((s) => s.type === "line");
  const lines = marked ? cfg.data.series.filter((s) => s.type === "line") : cfg.data.series.slice(-1);
  const cols = marked
    ? cfg.data.series.filter((s) => s.type !== "line")
    : cfg.data.series.slice(0, Math.max(1, cfg.data.series.length - 1));

  // One shared scale: whichever of stack totals / line values reaches higher.
  const stackMax = Math.max(
    0,
    ...cfg.data.categories.map((_, c) => cols.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0)),
  );
  const lineMax = Math.max(0, ...lines.flatMap((s) => s.values.filter((v): v is number => v != null)));
  const colCfg: ChartConfig = {
    ...cfg,
    kind: "stacked",
    data: { ...cfg.data, series: cols },
    scale: cfg.scale?.max != null ? cfg.scale : { ...cfg.scale, max: niceTicks(0, Math.max(stackMax, lineMax, 1)).pop() },
  };
  const result = layoutColumns(colCfg, style, decor);
  const { anchors, nodes } = result;
  if (!anchors.valueToY) return result;

  const fs = style.fontSize;
  const fmt = resolveFormat(lines.flatMap((s) => s.values.filter((v): v is number => v != null)), cfg.numberFormat);
  lines.forEach((s, li) => {
    const color = seriesColor(style, cols.length + li, s.color);
    let prev: { x: number; y: number } | null = null;
    let lastY: number | null = null;
    s.values.forEach((v, c) => {
      if (v == null || c >= anchors.categoryX.length) {
        prev = null;
        return;
      }
      const pt = { x: anchors.categoryX[c], y: anchors.valueToY!(v) };
      if (prev) nodes.push({ kind: "line", x1: prev.x, y1: prev.y, x2: pt.x, y2: pt.y, stroke: color, strokeWidth: 2, name: `combo-line-${li}-${c}` });
      const r = 2.4;
      nodes.push({ kind: "rect", x: pt.x - r, y: pt.y - r, w: r * 2, h: r * 2, fill: color, stroke: style.background, strokeWidth: 1, name: `combo-marker-${li}-${c}` });
      if (decor.segmentLabels) {
        nodes.push({
          kind: "text", x: pt.x - 30, y: pt.y - fs * 1.65, w: 60, h: fs * 1.4,
          text: formatNumber(v, fmt), fontSize: fs, color: style.text,
          align: "center", valign: "bottom", name: `combo-label-${li}-${c}`,
        });
      }
      prev = pt;
      lastY = pt.y;
    });
    if (decor.seriesLabels && lastY != null) {
      nodes.push({
        kind: "text",
        x: anchors.plot.x + anchors.plot.w + 4,
        y: lastY - fs * 1.6,
        w: cfg.width - (anchors.plot.x + anchors.plot.w) - 4,
        h: fs * 1.4,
        text: s.name,
        fontSize: fs,
        color: style.text,
        align: "left",
        valign: "middle",
        name: `combo-series-label-${li}`,
      });
    }
  });
  return result;
}

/** Chrome for horizontal (bar) orientation: title, legend, left category labels, bottom axis. */
function horizontalChrome(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  frame: Frame,
  centers: number[],
  scale: ValueScale,
  qOf: (v: number) => number,
): SceneNode[] {
  const nodes: SceneNode[] = chromeNodes(
    cfg,
    style,
    { ...decor, categoryAxis: false, valueAxis: false, gridlines: false },
    frame,
    centers,
  );
  const fs = style.fontSize;
  if (decor.gridlines) {
    for (const t of scale.ticks) {
      if (t === 0) continue;
      const x = frame.x + qOf(t);
      nodes.push({ kind: "line", x1: x, y1: frame.y, x2: x, y2: frame.y + frame.h, stroke: style.gridline, strokeWidth: 0.75, name: "gridline" });
    }
  }
  if (decor.valueAxis) {
    const axisFmt = resolveFormat(scale.ticks, cfg.numberFormat);
    for (const t of scale.ticks) {
      const x = frame.x + qOf(t);
      nodes.push({
        kind: "text",
        x: x - 24,
        y: frame.y + frame.h + 2,
        w: 48,
        h: fs * 1.4,
        text: formatNumber(t, axisFmt),
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: "center",
        valign: "top",
        name: "value-axis",
      });
    }
  }
  if (decor.categoryAxis) {
    cfg.data.categories.forEach((cat, i) => {
      nodes.push({
        kind: "text",
        x: 0,
        y: centers[i] - fs * 0.75,
        w: frame.x - 4,
        h: fs * 1.5,
        text: cat,
        fontSize: fs,
        color: style.text,
        align: "right",
        valign: "middle",
        name: `category-${i}`,
      });
    });
  }
  if (decor.seriesLabels && cfg.data.series.length > 1) {
    nodes.push(...legendRow(cfg, style, frame.x, (cfg.title ? fs * 1.6 + 6 : 0) + 2));
  }
  return nodes;
}

/** Horizontal legend row: color chip + series name, left to right. */
export function legendRow(cfg: ChartConfig, style: ChartStyle, x0: number, y: number): SceneNode[] {
  const fs = style.fontSize;
  const nodes: SceneNode[] = [];
  let x = x0;
  cfg.data.series.forEach((s, si) => {
    const chip = fs * 0.7;
    nodes.push(
      { kind: "rect", x, y: y + fs * 0.35, w: chip, h: chip, fill: seriesColor(style, si, s.color), name: `legend-chip-${si}` },
      {
        kind: "text",
        x: x + chip + 3,
        y,
        w: textWidth(s.name, fs) + 6,
        h: fs * 1.4,
        text: s.name,
        fontSize: fs,
        color: style.text,
        align: "left",
        valign: "middle",
        name: `legend-${si}`,
      },
    );
    x += chip + 3 + textWidth(s.name, fs) + 12;
  });
  return nodes;
}

/**
 * Right-hand series labels at the last column's segment midpoints,
 * greedily pushed apart so they never overlap (think-cell placement).
 */
export function seriesLabelNodes(
  cfg: ChartConfig,
  style: ChartStyle,
  frame: { x: number; y: number; w: number; h: number },
  midYs: (number | null)[],
): SceneNode[] {
  const fs = style.fontSize;
  const lineH = fs * 1.35;
  const entries = cfg.data.series
    .map((s, i) => ({ name: s.name, color: seriesColor(style, i, s.color), y: midYs[i] }))
    .filter((e): e is { name: string; color: string; y: number } => e.y != null)
    .sort((a, b) => a.y - b.y);
  // Push overlapping labels apart, then clamp back into the frame.
  for (let i = 1; i < entries.length; i++) {
    if (entries[i].y - entries[i - 1].y < lineH) entries[i].y = entries[i - 1].y + lineH;
  }
  const overflow = entries.length ? entries[entries.length - 1].y + lineH / 2 - (frame.y + frame.h) : 0;
  if (overflow > 0) {
    for (const e of entries) e.y -= overflow;
    for (let i = entries.length - 2; i >= 0; i--) {
      if (entries[i + 1].y - entries[i].y < lineH) entries[i].y = entries[i + 1].y - lineH;
    }
  }
  const x = frame.x + frame.w + 4;
  return entries.map((e, i) => ({
    kind: "text" as const,
    x,
    y: e.y - lineH / 2,
    w: cfg.width - x,
    h: lineH,
    text: e.name,
    fontSize: fs,
    color: style.text,
    align: "left" as const,
    valign: "middle" as const,
    name: `series-label-${i}`,
  }));
}
