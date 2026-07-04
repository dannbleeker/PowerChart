import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { baselineNode, categorySlots, chromeNodes, computeFrame, valueScale } from "./frame";

export interface LayoutResult {
  nodes: SceneNode[];
  anchors: LayoutAnchors;
}

/** Minimum segment height (relative to font size) before its label is hidden. */
const LABEL_FIT = 1.25;

/**
 * Stacked / clustered / 100% column charts.
 * Stacking follows think-cell: positives build up from the baseline,
 * negatives build down, category labels sit below the plot.
 */
export function layoutColumns(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data, kind } = cfg;
  const n = data.categories.length;
  const stacked = kind !== "clustered";
  const pct = kind === "stacked100";

  const { frame } = computeFrame(cfg, style, decor, decor.seriesLabels ? data.series.map((s) => s.name) : []);
  const slots = categorySlots(frame, n);
  const fs = style.fontSize;

  // Column totals (sum of positives / negatives separately for stacking extents).
  const posTotals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0),
  );
  const negTotals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0),
  );
  const signedTotals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + (s.values[c] ?? 0), 0),
  );
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
  const scale = pct
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1], toY: (v: number) => frame.y + frame.h - v * frame.h }
    : valueScale(frame, dataMin, dataMax);

  const nodes: SceneNode[] = chromeNodes(cfg, style, decor, frame, slots.centers, scale);
  const y0 = scale.toY(0);
  const columnTop: number[] = [];
  /** Segment mid-y of the last category per series, for right-hand series labels. */
  const lastSegMid: (number | null)[] = data.series.map(() => null);

  for (let c = 0; c < n; c++) {
    let upY = y0;
    let downY = y0;
    const cx = slots.centers[c];
    const total = pct ? posTotals[c] : 0;
    const barW = stacked ? slots.colWidth : slots.colWidth / Math.max(1, data.series.length);

    data.series.forEach((s, si) => {
      let v = s.values[c];
      if (v == null || v === 0) return;
      if (pct) {
        if (total <= 0) return;
        v = Math.max(0, v) / total;
      }
      const fill = seriesColor(style, si, s.color);
      let x: number, y: number, h: number;
      if (stacked) {
        x = cx - slots.colWidth / 2;
        const hh = (Math.abs(v) / (scale.max - scale.min || 1)) * frame.h;
        if (v >= 0) {
          y = upY - hh;
          upY = y;
        } else {
          y = downY;
          downY = y + hh;
        }
        h = hh;
      } else {
        x = cx - slots.colWidth / 2 + si * barW;
        const vy = scale.toY(v);
        y = Math.min(vy, y0);
        h = Math.abs(vy - y0);
      }
      nodes.push({
        kind: "rect",
        x,
        y,
        w: stacked ? slots.colWidth : barW - 1,
        h,
        fill,
        stroke: style.background,
        strokeWidth: stacked ? 0.75 : 0,
        name: `seg-${si}-${c}`,
      });
      if (c === n - 1) lastSegMid[si] = y + h / 2;
      if (decor.segmentLabels && h >= fs * LABEL_FIT) {
        const label = pct ? formatPercent(v) : formatNumber(s.values[c]!, fmt);
        if (textWidth(label, fs) <= (stacked ? slots.colWidth : barW) + 2) {
          nodes.push({
            kind: "text",
            x: x - 4,
            y: y + h / 2 - fs * 0.75,
            w: (stacked ? slots.colWidth : barW - 1) + 8,
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
    columnTop.push(Math.min(upY, y0));

    if (decor.totals && !pct) {
      nodes.push({
        kind: "text",
        x: cx - slots.slotWidth / 2,
        y: Math.min(upY, y0) - fs * 1.45,
        w: slots.slotWidth,
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

  nodes.push(baselineNode(frame, y0, style));

  if (decor.seriesLabels) {
    nodes.push(...seriesLabelNodes(cfg, style, frame, lastSegMid));
  }

  return {
    nodes,
    anchors: {
      categoryX: slots.centers,
      categoryWidth: data.categories.map(() => slots.colWidth),
      columnTop,
      columnValue: pct ? posTotals : signedTotals,
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: pct ? undefined : scale.toY,
    },
  };
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
