import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { baselineNode, breakMarkerNodes, chromeNodes, computeFrame, computeFrameHorizontal, valueScale } from "./frame";
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
export interface Seg {
  series: number;
  from: number;
  to: number;
  value: number;
}
export interface Bar {
  segs: Seg[];
  isTotal: boolean;
  value: number;
  level: number;
}

/**
 * Walk the bridge: one bar per category, each carrying its stacked segments and
 * the running level it leaves behind.
 *
 * The single source of truth for what a waterfall draws. It used to be inlined
 * here and re-implemented, differently, by the value extent — which walked only
 * series[0] and ignored spacers, so a stacked bridge's extent stopped at the
 * first series' total. Anything that needs to know how high a waterfall reaches
 * calls this.
 */
/** Category → the chain column its detail group decomposes. */
export function detailParents(cfg: ChartConfig): Map<number, number> {
  const map = new Map<number, number>();
  for (const g of cfg.waterfall?.detailGroups ?? []) {
    for (const i of g.indices) map.set(i, g.of);
  }
  return map;
}

export function waterfallChain(cfg: ChartConfig): Bar[] {
  const { data } = cfg;
  const n = data.categories.length;
  const nSeries = Math.max(1, data.series.length);
  const totalSet = new Set(cfg.waterfall?.totalIndices ?? []);
  const spacerSet = new Set(cfg.waterfall?.spacerIndices ?? []);
  const detailOf = detailParents(cfg);
  const bars: Bar[] = [];
  let running = 0;
  for (let c = 0; c < n; c++) {
    if (spacerSet.has(c)) {
      // Blank grouping gap: no bar, running total carries across unchanged.
      bars.push({ segs: [], isTotal: false, value: 0, level: running });
      continue;
    }
    const parent = detailOf.get(c);
    if (parent != null && bars[parent]) {
      // "Of which": decompose the parent's delta as a sub-bridge rising from
      // the parent's own base, so the group reads as that column taken apart
      // rather than as more steps in the walk. Off the chain — `running` is
      // untouched, exactly like a total, so every downstream total and the
      // final one stay correct whatever these add up to.
      const p = bars[parent];
      const base = p.isTotal ? 0 : p.level - p.value;
      // Chain within the group: pick up where the previous detail left off.
      const prev = detailOf.get(c - 1) === parent ? bars[c - 1] : null;
      let level = prev ? prev.level : base;
      const segs: Seg[] = [];
      for (let si = 0; si < nSeries; si++) {
        const v = data.series[si]?.values[c];
        if (v == null || v === 0) continue;
        segs.push({ series: si, from: level, to: level + v, value: v });
        level += v;
      }
      bars.push({ segs, isTotal: false, value: level - (prev ? prev.level : base), level });
      continue;
    }
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
  return bars;
}

/** The value range a waterfall's bars actually occupy. */
export function waterfallExtent(cfg: ChartConfig): { min: number; max: number } {
  const vals = waterfallChain(cfg).flatMap((b) => [b.level, ...b.segs.flatMap((s) => [s.from, s.to])]);
  return { min: Math.min(0, ...vals), max: Math.max(0, ...vals) };
}

export function layoutWaterfall(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const stacked = data.series.length > 1;
  const spacerSet = new Set(cfg.waterfall?.spacerIndices ?? []);
  const H = !!cfg.horizontal;
  const detailOf = detailParents(cfg);
  /** The next column the chain actually flows to, stepping over any details. */
  const nextChainColumn = (from: number) => {
    let j = from + 1;
    while (j < n && detailOf.has(j)) j++;
    return j;
  };

  const bars = waterfallChain(cfg);

  const allLevels = bars.flatMap((b) => b.segs.flatMap((s) => [s.from, s.to]));
  const hi = Math.max(0, ...allLevels);
  const lo = Math.min(0, ...allLevels);
  const fmt = resolveFormat(
    bars.flatMap((b) => b.segs.map((s) => s.value)),
    cfg.numberFormat,
  );

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
        kind: "rect",
        ...r,
        fill,
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
    columnTop.push(b.segs.length ? (H ? frame.x + topQ : frame.y + frame.h - topQ) : y0);

    // Spacer slot: bridge the connector across the empty gap at the running
    // level so the dashed line reads as continuous through the section break.
    if (b.segs.length === 0 && spacerSet.has(c) && c > 0) {
      const levelQ = qOf(b.level);
      if (H) {
        const x = frame.x + levelQ;
        nodes.push({
          kind: "line",
          x1: x,
          y1: centers[c] - colThick / 2,
          x2: x,
          y2: centers[c] + colThick / 2,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          dash: [1.5, 1.5],
          name: `spacer-bridge-${c}`,
        });
      } else {
        const y = frame.y + frame.h - levelQ;
        nodes.push({
          kind: "line",
          x1: centers[c] - colThick / 2,
          y1: y,
          x2: centers[c] + colThick / 2,
          y2: y,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          dash: [1.5, 1.5],
          name: `spacer-bridge-${c}`,
        });
      }
    }

    // Dashed connector from this column's outgoing level to the next bar. A
    // detail column has no outgoing level to carry — the chain flows past the
    // whole group — so it draws none, and its parent's connector reaches over
    // the group to the next chain column instead. That step-over is what tells
    // the reader the group is a decomposition and not more of the walk.
    const isDetail = detailOf.has(c);
    const to = isDetail ? -1 : nextChainColumn(c);
    if (!isDetail && to > c && to < n) {
      const levelQ = qOf(b.level);
      if (H) {
        const x = frame.x + levelQ;
        nodes.push({
          kind: "line",
          x1: x,
          y1: centers[c] + colThick / 2,
          x2: x,
          y2: centers[to] - colThick / 2,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          dash: [1.5, 1.5],
          name: `connector-${c}`,
        });
      } else {
        const y = frame.y + frame.h - levelQ;
        nodes.push({
          kind: "line",
          x1: centers[c] + colThick / 2,
          y1: y,
          x2: centers[to] - colThick / 2,
          y2: y,
          stroke: style.mutedText,
          strokeWidth: 0.75,
          dash: [1.5, 1.5],
          name: `connector-${c}`,
        });
      }
    }
  });

  if (!H) nodes.push(...breakMarkerNodes(frame, scale, style));
  if (H) {
    nodes.push({
      kind: "line",
      x1: y0,
      y1: frame.y,
      x2: y0,
      y2: frame.y + frame.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
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
  // "Σ".toLowerCase() is "σ" — compare against the lowercased form.
  return t === "e" || t === "=" || t === "σ";
}
