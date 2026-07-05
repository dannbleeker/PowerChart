import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { seriesColor } from "../style";
import {
  baselineNode,
  breakMarkerNodes,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  valueScale,
} from "./frame";
import { horizontalChrome, type LayoutResult } from "./column";

/**
 * think-cell style waterfall: deltas build a running total; categories marked
 * as totals ("e" cells in the datasheet) draw a full bar from the baseline to
 * the running total. Dashed connectors join consecutive bar tops.
 *
 * Stacked waterfall: with several series, each column stacks its per-series
 * contributions from the incoming level and the running total moves by the
 * column sum. Set cfg.horizontal for the rotated (bar) variant.
 */
export function layoutWaterfall(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const nSeries = Math.max(1, data.series.length);
  const stacked = data.series.length > 1;
  const totalSet = new Set(cfg.waterfall?.totalIndices ?? []);
  const H = !!cfg.horizontal;

  interface Seg {
    series: number;
    from: number;
    to: number;
    value: number;
  }
  interface Bar {
    segs: Seg[];
    isTotal: boolean;
    value: number;
    level: number;
  }

  const bars: Bar[] = [];
  let running = 0;
  for (let c = 0; c < n; c++) {
    if (totalSet.has(c)) {
      bars.push({
        segs: [{ series: 0, from: 0, to: running, value: running }],
        isTotal: true,
        value: running,
        level: running,
      });
      continue;
    }
    const segs: Seg[] = [];
    let level = running;
    for (let si = 0; si < nSeries; si++) {
      const v = data.series[si]?.values[c];
      if (v == null || v === 0) continue;
      segs.push({ series: si, from: level, to: level + v, value: v });
      level += v;
    }
    bars.push({ segs, isTotal: false, value: level - running, level });
    running = level;
  }

  const allLevels = bars.flatMap((b) => b.segs.flatMap((s) => [s.from, s.to]));
  const hi = Math.max(0, ...allLevels);
  const lo = Math.min(0, ...allLevels);
  const fmt = resolveFormat(bars.flatMap((b) => b.segs.map((s) => s.value)), cfg.numberFormat);

  const frame = H
    ? computeFrameHorizontal(cfg, style, { ...decor, seriesLabels: false })
    : computeFrame(cfg, style, { ...decor, seriesLabels: false }, []).frame;
  const scale = valueScale(frame, lo, hi, cfg.scale, H ? undefined : cfg.axisBreak);
  const fs = style.fontSize;

  // Category slots along x (vertical) or y (horizontal); value axis the other way.
  const catStart = H ? frame.y : frame.x;
  const catLen = H ? frame.h : frame.w;
  const slotLen = catLen / Math.max(1, n);
  const colThick = slotLen * (2 / 3);
  const centers = Array.from({ length: n }, (_, i) => catStart + slotLen * (i + 0.5));
  const valLen = H ? frame.w : frame.h;
  const qOf = H
    ? (v: number) => ((v - scale.min) / (scale.max - scale.min || 1)) * valLen
    : (v: number) => frame.y + frame.h - scale.toY(v);
  const vRect = (catPos: number, v0: number, v1: number) => {
    const q0 = Math.min(qOf(v0), qOf(v1));
    const q1 = Math.max(qOf(v0), qOf(v1));
    return H
      ? { x: frame.x + q0, y: catPos - colThick / 2, w: Math.max(0.5, q1 - q0), h: colThick }
      : { x: catPos - colThick / 2, y: frame.y + frame.h - q1, w: colThick, h: Math.max(0.5, q1 - q0) };
  };

  const nodes: SceneNode[] = H
    ? horizontalChrome(cfg, style, { ...decor, seriesLabels: false }, frame, centers, scale, qOf)
    : chromeNodes(cfg, style, decor, frame, centers, scale);
  const zeroQ = qOf(0);
  const y0 = H ? frame.x + zeroQ : frame.y + frame.h - zeroQ;
  const columnTop: number[] = [];

  bars.forEach((b, c) => {
    let topQ = 0;

    b.segs.forEach((seg) => {
      const r = vRect(centers[c], seg.from, seg.to);
      topQ = Math.max(topQ, qOf(Math.max(seg.from, seg.to)));
      const fill = b.isTotal
        ? style.neutral
        : stacked
          ? seriesColor(style, seg.series, data.series[seg.series]?.color)
          : seg.value >= 0
            ? style.palette[0]
            : style.negative;
      nodes.push({
        kind: "rect", ...r, fill,
        stroke: stacked ? style.background : undefined,
        strokeWidth: stacked ? 0.75 : 0,
        name: `bar-${c}${stacked && !b.isTotal ? `-s${seg.series}` : ""}`,
      });

      // Base columns (stack starting at zero) get unsigned labels throughout.
      const floating = !b.isTotal && b.segs[0]?.from !== 0;
      const label = formatNumber(seg.value, {
        ...fmt,
        forceSign: floating && (cfg.numberFormat?.forceSign ?? true),
      });
      const along = H ? r.w : r.h;
      const fits = H ? along >= textWidth(label, fs) + 2 : along >= fs * 1.25;
      if (fits || !stacked || b.isTotal) {
        nodes.push({
          kind: "text",
          x: fits ? r.x - 6 : H ? r.x + r.w + 2 : r.x - 6,
          y: fits ? r.y + r.h / 2 - fs * 0.75 : H ? r.y + r.h / 2 - fs * 0.75 : r.y - fs * 1.45,
          w: fits ? r.w + 12 : H ? fs * 4 : r.w + 12,
          h: fs * 1.5,
          text: label,
          fontSize: fs,
          color: fits ? contrastInk(fill) : style.text,
          bold: b.isTotal,
          align: fits ? "center" : H ? "left" : "center",
          valign: "middle",
          name: `label-${c}${stacked && !b.isTotal ? `-s${seg.series}` : ""}`,
        });
      }
    });
    columnTop.push(
      b.segs.length ? (H ? frame.x + topQ : frame.y + frame.h - topQ) : y0,
    );

    // Dashed connector from this column's outgoing level to the next bar.
    if (c < n - 1) {
      const levelQ = qOf(b.level);
      if (H) {
        const x = frame.x + levelQ;
        nodes.push({
          kind: "line",
          x1: x, y1: centers[c] + colThick / 2,
          x2: x, y2: centers[c + 1] - colThick / 2,
          stroke: style.mutedText, strokeWidth: 0.75, dash: [1.5, 1.5], name: `connector-${c}`,
        });
      } else {
        const y = frame.y + frame.h - levelQ;
        nodes.push({
          kind: "line",
          x1: centers[c] + colThick / 2, y1: y,
          x2: centers[c + 1] - colThick / 2, y2: y,
          stroke: style.mutedText, strokeWidth: 0.75, dash: [1.5, 1.5], name: `connector-${c}`,
        });
      }
    }
  });

  if (!H) nodes.push(...breakMarkerNodes(frame, scale, style));
  if (H) {
    nodes.push({ kind: "line", x1: y0, y1: frame.y, x2: y0, y2: frame.y + frame.h, stroke: style.axis, strokeWidth: 1, name: "baseline" });
  } else {
    nodes.push(baselineNode(frame, y0, style));
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => colThick),
      columnTop,
      columnValue: bars.map((b) => (b.isTotal ? b.value : b.level)),
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: H ? undefined : scale.toY,
    },
  };
}

/** Parse a datasheet column value for waterfalls: "e"/"=" marks a computed total. */
export function isTotalToken(raw: string): boolean {
  const t = raw.trim().toLowerCase();
  return t === "e" || t === "=" || t === "Σ";
}
