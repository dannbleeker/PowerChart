import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { chromeNodes, computeFrame } from "./frame";
import { seriesLabelNodes, type LayoutResult } from "./column";

/**
 * Mekko (Marimekko) chart, think-cell style. Two variants:
 * - %-axis (default): column widths proportional to column totals, columns
 *   normalized to full height, so segment area ∝ absolute value.
 * - "Mekko with units": explicit column widths from the datasheet's
 *   `X extent` row; column heights represent absolute totals on a value scale.
 */
export function layoutMekko(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  const units = !!data.xExtent?.some((v) => v != null && v > 0);

  const totals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0),
  );
  const extents = units
    ? data.categories.map((_, c) => Math.max(0, data.xExtent?.[c] ?? 0))
    : totals;
  const maxTotal = Math.max(1e-9, ...totals);
  const grand = extents.reduce((a, b) => a + b, 0) || 1;
  const fmt = resolveFormat(
    [...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)), ...totals],
    cfg.numberFormat,
  );

  const decorFull = { ...decor, totals: true };
  const { frame } = computeFrame(cfg, style, decorFull, decor.seriesLabels ? data.series.map((s) => s.name) : []);
  const gap = 2;

  const nodes: SceneNode[] = [];
  const centers: number[] = [];
  const widths: number[] = [];
  const columnTop: number[] = [];
  const lastSegMid: (number | null)[] = data.series.map(() => null);

  let x = frame.x;
  const usableW = frame.w - gap * (n - 1);
  for (let c = 0; c < n; c++) {
    const w = (extents[c] / grand) * usableW;
    const cx = x + w / 2;
    centers.push(cx);
    widths.push(w);
    // %-variant: every column fills the plot; units variant: height ∝ total.
    const colH = units ? (totals[c] / maxTotal) * frame.h : frame.h;
    const colTop = frame.y + frame.h - colH;

    let y = frame.y + frame.h;
    data.series.forEach((s, si) => {
      const v = Math.max(0, s.values[c] ?? 0);
      if (v === 0 || totals[c] === 0) return;
      const h = (v / totals[c]) * colH;
      y -= h;
      const fill = seriesColor(style, si, s.color);
      nodes.push({ kind: "rect", x, y, w, h, fill, stroke: style.background, strokeWidth: 0.75, name: `seg-${si}-${c}` });
      if (c === n - 1) lastSegMid[si] = y + h / 2;
      if (decor.segmentLabels && h >= fs * 1.25) {
        const label = formatNumber(v, fmt);
        if (textWidth(label, fs) <= w + 2) {
          nodes.push({
            kind: "text",
            x: x - 2,
            y: y + h / 2 - fs * 0.75,
            w: w + 4,
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
    columnTop.push(colTop);

    // Column total (absolute) and width share on top — the Mekko signature.
    nodes.push({
      kind: "text",
      x: x - 4,
      y: colTop - fs * 1.45,
      w: w + 8,
      h: fs * 1.4,
      text: formatNumber(totals[c], fmt),
      fontSize: fs,
      bold: true,
      color: style.text,
      align: "center",
      valign: "bottom",
      name: `total-${c}`,
    });

    x += w + gap;
  }

  // Category labels need the variable-width centers, so emit chrome without them
  // and add custom category labels.
  nodes.push(...chromeNodes(cfg, style, { ...decorFull, categoryAxis: false, valueAxis: false, gridlines: false }, frame, centers));
  if (decor.categoryAxis) {
    for (let c = 0; c < n; c++) {
      nodes.push({
        kind: "text",
        x: centers[c] - widths[c] / 2 - 4,
        y: frame.y + frame.h + 3,
        w: widths[c] + 8,
        h: fs * 1.4,
        text: units
          ? data.categories[c]
          : `${data.categories[c]} (${formatPercent(extents[c] / grand)})`,
        fontSize: fs,
        color: style.text,
        align: "center",
        valign: "top",
        name: `category-${c}`,
      });
    }
  }
  nodes.push({
    kind: "line",
    x1: frame.x,
    y1: frame.y + frame.h,
    x2: frame.x + frame.w,
    y2: frame.y + frame.h,
    stroke: style.axis,
    strokeWidth: 1,
    name: "baseline",
  });

  if (decor.seriesLabels) nodes.push(...seriesLabelNodes(cfg, style, frame, lastSegMid));

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: widths,
      columnTop,
      columnValue: totals,
      baselineY: frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
    },
  };
}
