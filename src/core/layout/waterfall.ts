import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { baselineNode, breakMarkerNodes, categorySlots, chromeNodes, computeFrame, valueScale } from "./frame";
import type { LayoutResult } from "./column";

/**
 * think-cell style waterfall: deltas build a running total; categories marked
 * as totals ("e" cells in the datasheet) draw a full bar from the baseline to
 * the running total. Dashed connectors join consecutive bar tops.
 *
 * Stacked waterfall: with several series, each column stacks its per-series
 * contributions from the incoming level (in sheet order, positives up and
 * negatives down from the running level), and the running total moves by the
 * column sum — think-cell's multi-series waterfall.
 */
export function layoutWaterfall(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const nSeries = Math.max(1, data.series.length);
  const stacked = data.series.length > 1;
  const totalSet = new Set(cfg.waterfall?.totalIndices ?? []);

  interface Seg {
    series: number;
    from: number;
    to: number;
    value: number;
  }
  interface Bar {
    segs: Seg[];
    isTotal: boolean;
    /** Column delta (or running total for total bars). */
    value: number;
    /** Running level after this column. */
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

  const { frame } = computeFrame(cfg, style, { ...decor, seriesLabels: false }, []);
  const slots = categorySlots(frame, n);
  const scale = valueScale(frame, lo, hi, cfg.scale, cfg.axisBreak);
  const fs = style.fontSize;
  const y0 = scale.toY(0);

  const nodes: SceneNode[] = chromeNodes(cfg, style, decor, frame, slots.centers, scale);
  const columnTop: number[] = [];

  bars.forEach((b, c) => {
    const cx = slots.centers[c];
    const x = cx - slots.colWidth / 2;
    let top = y0;

    // Base columns (stack starting at zero) get unsigned labels throughout.
    const floating = !b.isTotal && b.segs[0]?.from !== 0;
    b.segs.forEach((seg) => {
      const yA = scale.toY(seg.from);
      const yB = scale.toY(seg.to);
      const y = Math.min(yA, yB);
      const h = Math.max(0.5, Math.abs(yA - yB));
      top = Math.min(top, y);
      const fill = b.isTotal
        ? style.neutral
        : stacked
          ? seriesColor(style, seg.series, data.series[seg.series]?.color)
          : seg.value >= 0
            ? style.palette[0]
            : style.negative;
      nodes.push({
        kind: "rect", x, y, w: slots.colWidth, h, fill,
        stroke: stacked ? style.background : undefined,
        strokeWidth: stacked ? 0.75 : 0,
        name: `bar-${c}${stacked && !b.isTotal ? `-s${seg.series}` : ""}`,
      });

      // Per-segment label, signed for floating deltas.
      const label = formatNumber(seg.value, {
        ...fmt,
        forceSign: floating && (cfg.numberFormat?.forceSign ?? true),
      });
      const fits = h >= fs * 1.25;
      if (fits || !stacked || b.isTotal) {
        nodes.push({
          kind: "text",
          x: x - 6,
          y: fits ? y + h / 2 - fs * 0.75 : y - fs * 1.45,
          w: slots.colWidth + 12,
          h: fs * 1.5,
          text: label,
          fontSize: fs,
          color: fits ? contrastInk(fill) : style.text,
          bold: b.isTotal,
          align: "center",
          valign: "middle",
          name: `label-${c}${stacked && !b.isTotal ? `-s${seg.series}` : ""}`,
        });
      }
    });
    columnTop.push(b.segs.length ? top : y0);

    // Dashed connector from this column's outgoing level to the next bar.
    if (c < n - 1) {
      const levelY = scale.toY(b.level);
      const nextX = slots.centers[c + 1] - slots.colWidth / 2;
      nodes.push({
        kind: "line",
        x1: x + slots.colWidth,
        y1: levelY,
        x2: nextX,
        y2: levelY,
        stroke: style.mutedText,
        strokeWidth: 0.75,
        dash: [1.5, 1.5],
        name: `connector-${c}`,
      });
    }
  });

  nodes.push(...breakMarkerNodes(frame, scale, style));
  nodes.push(baselineNode(frame, y0, style));

  return {
    nodes,
    anchors: {
      categoryX: slots.centers,
      categoryWidth: data.categories.map(() => slots.colWidth),
      columnTop,
      columnValue: bars.map((b) => (b.isTotal ? b.value : b.level)),
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: scale.toY,
    },
  };
}

/** Parse a datasheet column value for waterfalls: "e"/"=" marks a computed total. */
export function isTotalToken(raw: string): boolean {
  const t = raw.trim().toLowerCase();
  return t === "e" || t === "=" || t === "Σ";
}
