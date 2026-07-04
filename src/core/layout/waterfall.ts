import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { baselineNode, breakMarkerNodes, categorySlots, chromeNodes, computeFrame, valueScale } from "./frame";
import type { LayoutResult } from "./column";

/**
 * think-cell style waterfall: deltas build a running total; categories marked
 * as totals ("e" cells in the datasheet) draw a full bar from the baseline to
 * the running total. Dashed connectors join consecutive bar tops.
 */
export function layoutWaterfall(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const totalSet = new Set(cfg.waterfall?.totalIndices ?? []);
  const deltas = data.categories.map((_, c) => data.series[0]?.values[c] ?? 0);

  // Running totals and per-bar extents.
  const bars: { from: number; to: number; isTotal: boolean; value: number }[] = [];
  let running = 0;
  for (let c = 0; c < n; c++) {
    if (totalSet.has(c)) {
      bars.push({ from: 0, to: running, isTotal: true, value: running });
    } else {
      const d = deltas[c] ?? 0;
      bars.push({ from: running, to: running + d, isTotal: false, value: d });
      running += d;
    }
  }

  const hi = Math.max(0, ...bars.map((b) => Math.max(b.from, b.to)));
  const lo = Math.min(0, ...bars.map((b) => Math.min(b.from, b.to)));
  const fmt = resolveFormat(bars.map((b) => b.value), cfg.numberFormat);

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
    const yA = scale.toY(b.from);
    const yB = scale.toY(b.to);
    const y = Math.min(yA, yB);
    const h = Math.max(0.5, Math.abs(yA - yB));
    const fill = b.isTotal ? style.neutral : b.value >= 0 ? style.palette[0] : style.negative;
    nodes.push({ kind: "rect", x, y, w: slots.colWidth, h, fill, name: `bar-${c}` });
    columnTop.push(y);

    // Value label: inside if it fits, otherwise just outside the bar.
    // Signed for floating deltas; plain for totals and bars starting at zero.
    const label = formatNumber(b.value, {
      ...fmt,
      forceSign: !b.isTotal && b.from !== 0 && (cfg.numberFormat?.forceSign ?? true),
    });
    const fits = h >= fs * 1.25;
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
      name: `label-${c}`,
    });

    // Dashed connector from this bar's end level to the next bar.
    if (c < n - 1) {
      const levelY = scale.toY(b.to);
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
      columnValue: bars.map((b) => (b.isTotal ? b.value : b.to)),
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
