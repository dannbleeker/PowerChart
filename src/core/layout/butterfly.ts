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
  // Flanks: with butterfly.split, the first `split` series stack on the left
  // and the rest on the right; otherwise the classic series 0 / series 1.
  const withIdx = data.series.map((s, si) => ({ s, si }));
  const split = cfg.butterfly?.split;
  const leftSeries = split != null ? withIdx.slice(0, split) : withIdx.slice(0, 1);
  const rightSeries = split != null ? withIdx.slice(split) : withIdx.slice(1, 2);
  const stacked = leftSeries.length > 1 || rightSeries.length > 1;
  const stackSum = (series: typeof withIdx, c: number) =>
    series.reduce((a, { s }) => a + Math.abs(s.values[c] ?? 0), 0);
  const signedSum = (series: typeof withIdx, c: number) =>
    series.reduce((a, { s }) => a + (s.values[c] ?? 0), 0);

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const headerH = fs * 1.6;
  const gutterW = Math.min(
    cfg.width * 0.3,
    Math.max(0, ...data.categories.map((c) => textWidth(c, fs))) + 12,
  );
  const valueW = fs * 3.4; // room for outside value labels on each flank
  // A value axis reserves a strip at the bottom for tick labels on both flanks.
  const axisH = decor.valueAxis ? fs * 1.5 : 0;
  const plot = {
    x: valueW,
    y: titleH + headerH + 2,
    w: cfg.width - valueW * 2,
    h: cfg.height - titleH - headerH - 6 - axisH,
  };
  const halfW = (plot.w - gutterW) / 2;
  const leftEdge = plot.x + halfW; // right edge of the left half
  const rightEdge = leftEdge + gutterW; // left edge of the right half

  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null)).map((v) => Math.abs(v));
  const sums = data.categories.flatMap((_, c) => [stackSum(leftSeries, c), stackSum(rightSeries, c)]);
  const ticks = niceTicks(0, Math.max(1, ...sums), 4);
  const max = ticks[ticks.length - 1];
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
  if (!stacked) {
    // Series headers above each half (classic two-series butterfly).
    ([[leftSeries[0], plot.x, leftEdge], [rightSeries[0], rightEdge, plot.x + plot.w]] as const).forEach(
      ([entry, x0, x1], i) => {
        nodes.push({
          kind: "text", x: x0, y: titleH, w: x1 - x0, h: headerH,
          text: entry?.s.name ?? "", fontSize: fs, bold: true,
          color: seriesColor(style, entry?.si ?? i, entry?.s.color), align: "center", valign: "middle", name: `header-${i}`,
        });
      },
    );
  } else {
    // Stacked flanks: one legend of all series across the top.
    let lx = plot.x;
    for (const { s, si } of [...leftSeries, ...rightSeries]) {
      const chip = fs * 0.7;
      nodes.push(
        { kind: "rect", x: lx, y: titleH + fs * 0.35, w: chip, h: chip, fill: seriesColor(style, si, s.color), name: `legend-chip-${si}` },
        {
          kind: "text", x: lx + chip + 3, y: titleH, w: textWidth(s.name, fs) + 6, h: headerH,
          text: s.name, fontSize: fs, color: style.text, align: "left", valign: "middle", name: `legend-${si}`,
        },
      );
      lx += chip + 3 + textWidth(s.name, fs) + 12;
    }
  }

  // Value gridlines mirrored on both flanks, drawn behind the bars.
  if (decor.gridlines) {
    for (const tk of ticks) {
      if (tk <= 0) continue;
      const q = qOf(tk);
      for (const x of [leftEdge - q, rightEdge + q]) {
        nodes.push({ kind: "line", x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: style.gridline, strokeWidth: 1, name: `gridline-${tk}` });
      }
    }
  }

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
    const drawSide = (series: typeof withIdx, dir: -1 | 1, edge: number) => {
      let offset = 0;
      for (const { s, si } of series) {
        const v = s.values[c];
        if (v == null) continue;
        const len = qOf(v);
        const x = dir < 0 ? edge - offset - len : edge + offset;
        const fill = seriesColor(style, si, s.color);
        nodes.push({ kind: "rect", x, y: cy - barH / 2, w: len, h: barH, fill, name: `seg-${si}-${c}` });
        if (decor.segmentLabels) {
          const label = formatNumber(v, fmt);
          const single = series.length === 1;
          const inside = len >= textWidth(label, fs) + 4 && barH >= fs * 1.25;
          // Stacked segments only label when the value fits inside; single
          // flanks fall back to an outside label (classic behaviour).
          if (inside || single) {
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
        }
        offset += len;
      }
    };
    drawSide(leftSeries, -1, leftEdge);
    drawSide(rightSeries, 1, rightEdge);
  }

  // Center axis lines flanking the gutter.
  for (const x of [leftEdge, rightEdge]) {
    nodes.push({ kind: "line", x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: style.axis, strokeWidth: 1, name: "baseline" });
  }

  // Value tick labels on both flanks, in the reserved bottom strip.
  if (decor.valueAxis) {
    const ty = plot.y + plot.h + 1;
    for (const tk of ticks) {
      const q = qOf(tk);
      const label = formatNumber(tk, fmt);
      // 0 sits at the inner edges (the gutter sides); other ticks mirror outward.
      const xs = tk === 0 ? [leftEdge, rightEdge] : [leftEdge - q, rightEdge + q];
      xs.forEach((x, side) => {
        nodes.push({
          kind: "text", x: x - valueW / 2, y: ty, w: valueW, h: axisH,
          text: label, fontSize: fs * 0.85, color: style.mutedText,
          align: "center", valign: "middle", name: `tick-${tk}-${side === 0 ? "l" : "r"}`,
        });
      });
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.y + slotH * (c + 0.5)),
      categoryWidth: data.categories.map(() => barH),
      columnTop,
      columnValue: data.categories.map((_, c) => signedSum(rightSeries, c) - signedSum(leftSeries, c)),
      baselineY: leftEdge,
      plot,
    },
  };
}
