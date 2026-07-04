import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { niceTicks } from "../format";
import type { LayoutResult } from "./column";

/**
 * Butterfly (tornado) chart: think-cell models this as two bar charts placed
 * back-to-back, one rotated 180°, sharing the same scale. Uses the first two
 * series; category labels sit in the center gutter.
 */
export function layoutButterfly(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  const left = data.series[0] ?? { name: "", values: [] };
  const right = data.series[1] ?? { name: "", values: [] };

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const headerH = fs * 1.6;
  const gutterW = Math.min(
    cfg.width * 0.3,
    Math.max(0, ...data.categories.map((c) => textWidth(c, fs))) + 12,
  );
  const valueW = fs * 3.4; // room for outside value labels on each flank
  const plot = {
    x: valueW,
    y: titleH + headerH + 2,
    w: cfg.width - valueW * 2,
    h: cfg.height - titleH - headerH - 6,
  };
  const halfW = (plot.w - gutterW) / 2;
  const leftEdge = plot.x + halfW; // right edge of the left half
  const rightEdge = leftEdge + gutterW; // left edge of the right half

  const all = [...left.values, ...right.values].filter((v): v is number => v != null).map((v) => Math.abs(v));
  const max = niceTicks(0, Math.max(1, ...all), 4).pop()!;
  const fmt = resolveFormat(all, cfg.numberFormat);
  const qOf = (v: number) => (Math.abs(v) / max) * halfW;

  const slotH = plot.h / Math.max(1, n);
  const barH = slotH * (2 / 3);

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }
  // Series headers above each half.
  ([[left, plot.x, leftEdge, 0], [right, rightEdge, plot.x + plot.w, 1]] as const).forEach(
    ([s, x0, x1, si]) => {
      nodes.push({
        kind: "text", x: x0, y: titleH, w: x1 - x0, h: headerH,
        text: s.name, fontSize: fs, bold: true,
        color: seriesColor(style, si, s.color), align: "center", valign: "middle", name: `header-${si}`,
      });
    },
  );

  const columnTop: number[] = [];
  for (let c = 0; c < n; c++) {
    const cy = plot.y + slotH * (c + 0.5);
    columnTop.push(cy - barH / 2);
    // Category label in the center gutter.
    nodes.push({
      kind: "text", x: leftEdge, y: cy - fs * 0.75, w: gutterW, h: fs * 1.5,
      text: data.categories[c], fontSize: fs, color: style.text,
      align: "center", valign: "middle", name: `category-${c}`,
    });
    ([[left, -1, 0], [right, 1, 1]] as const).forEach(([s, dir, si]) => {
      const v = s.values[c];
      if (v == null) return;
      const len = qOf(v);
      const x = dir < 0 ? leftEdge - len : rightEdge;
      const fill = seriesColor(style, si, s.color);
      nodes.push({ kind: "rect", x, y: cy - barH / 2, w: len, h: barH, fill, name: `seg-${si}-${c}` });
      if (decor.segmentLabels) {
        const label = formatNumber(v, fmt);
        const inside = len >= textWidth(label, fs) + 4 && barH >= fs * 1.25;
        nodes.push({
          kind: "text",
          x: inside ? x : dir < 0 ? x - fs * 3.4 - 2 : x + len + 2,
          y: cy - fs * 0.75,
          w: inside ? len : fs * 3.4,
          h: fs * 1.5,
          text: label,
          fontSize: fs,
          color: inside ? contrastInk(fill) : style.text,
          align: inside ? "center" : dir < 0 ? "right" : "left",
          valign: "middle",
          name: `label-${si}-${c}`,
        });
      }
    });
  }

  // Center axis lines flanking the gutter.
  for (const x of [leftEdge, rightEdge]) {
    nodes.push({ kind: "line", x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: style.axis, strokeWidth: 1, name: "baseline" });
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.y + slotH * (c + 0.5)),
      categoryWidth: data.categories.map(() => barH),
      columnTop,
      columnValue: data.categories.map((_, c) => (right.values[c] ?? 0) - (left.values[c] ?? 0)),
      baselineY: leftEdge,
      plot,
    },
  };
}
