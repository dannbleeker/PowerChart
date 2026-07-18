import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Funnel chart: pipeline stages as horizontally centered bands whose width
 * is proportional to the value (the Power BI convention — plain rectangles,
 * so it renders natively in every renderer). Between stages, a muted
 * conversion label states the % of the previous stage. Order the values
 * ascending for a pyramid.
 */
export function layoutFunnel(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const values = data.categories.map((_, c) => Math.max(0, data.series[0]?.values[c] ?? 0));
  const vMax = Math.max(1, ...values);
  const fmt = resolveFormat(values, cfg.numberFormat);

  const titleH = titleHeight(cfg, style);
  // Left gutter for stage names when category labels are on (default).
  const catW = decor.categoryAxis
    ? Math.min(cfg.width * 0.28, Math.max(0, ...data.categories.map((c) => textWidth(c, fs))) + 10)
    : 2;
  const plot = {
    x: catW,
    y: titleH + 2,
    w: cfg.width - catW - 4,
    h: cfg.height - titleH - 2 - footnoteH(cfg, style, decor) - 4,
  };
  const gap = fs * 1.5; // room for the conversion label between bands
  // Floor at a positive height: on a short frame with many stages the gaps can
  // exceed the plot, driving every band negative (SVG then drops the rects).
  const bandH = Math.max(1, (plot.h - gap * (n - 1)) / Math.max(1, n));
  const cx = plot.x + plot.w / 2;

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  const columnTop: number[] = [];
  values.forEach((v, c) => {
    const y = plot.y + c * (bandH + gap);
    columnTop.push(y);
    const w = Math.max(2, (v / vMax) * plot.w);
    const fill = data.series[0]?.colors?.[c] ?? style.palette[c % style.palette.length];
    nodes.push({ kind: "rect", x: cx - w / 2, y, w, h: bandH, fill, name: `stage-${c}` });

    if (decor.categoryAxis) {
      nodes.push({
        kind: "text",
        x: 0,
        y: y + bandH / 2 - fs * 0.75,
        w: catW - 4,
        h: fs * 1.5,
        text: data.categories[c],
        fontSize: fs,
        color: style.text,
        align: "right",
        valign: "middle",
        name: `category-${c}`,
      });
    }
    if (decor.segmentLabels) {
      const label = formatNumber(v, fmt);
      const inside = w >= textWidth(label, fs) + 8 && bandH >= fs * 1.3;
      nodes.push({
        kind: "text",
        x: inside ? cx - w / 2 : cx + w / 2 + 4,
        y: y + bandH / 2 - fs * 0.75,
        w: inside ? w : textWidth(label, fs) + 6,
        h: fs * 1.5,
        text: label,
        fontSize: fs,
        bold: true,
        color: inside ? contrastInk(fill) : style.text,
        align: inside ? "center" : "left",
        valign: "middle",
        name: `stage-value-${c}`,
      });
    }
    // Conversion rate vs the previous stage, in the gap between bands.
    if (c > 0 && values[c - 1] > 0) {
      nodes.push({
        kind: "text",
        x: cx - 40,
        y: y - gap,
        w: 80,
        h: gap,
        text: `▾ ${formatPercent(v / values[c - 1], 1)}`,
        fontSize: fs * 0.85,
        color: style.mutedText,
        align: "center",
        valign: "middle",
        name: `conversion-${c}`,
      });
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: values.map(() => cx),
      categoryWidth: values.map((v) => Math.max(2, (v / vMax) * plot.w)),
      columnTop,
      columnValue: values,
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
