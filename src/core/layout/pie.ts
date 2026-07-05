import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { polar, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat, segmentLabel } from "../format";
import type { LayoutResult } from "./column";

/**
 * Pie / doughnut chart from the first series; one slice per category,
 * colored by category. Slices start at 12 o'clock, clockwise, largest-first
 * ordering left to the data (think-cell keeps sheet order too).
 * Labels sit inside the slice when it is wide enough, otherwise outside.
 */
export function layoutPie(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const values = data.categories.map((_, c) => Math.max(0, data.series[0]?.values[c] ?? 0));
  const total = values.reduce((a, b) => a + b, 0) || 1;
  const fmt = resolveFormat(values, cfg.numberFormat);

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const cx = cfg.width / 2;
  const cy = titleH + (cfg.height - titleH) / 2;
  const r = Math.min(cfg.width * 0.5 - fs * 7, (cfg.height - titleH) / 2 - fs * 2.2);
  const doughnut = cfg.kind === "doughnut";

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }

  let angle = 0;
  values.forEach((v, c) => {
    const span = (v / total) * 360;
    if (span <= 0) return;
    const fill = style.palette[c % style.palette.length];
    nodes.push({
      kind: "wedge", cx, cy, r, innerR: 0,
      startAngle: angle, endAngle: angle + span,
      fill, stroke: style.background, strokeWidth: 1, name: `slice-${c}`,
    });

    if (decor.segmentLabels) {
      const mid = angle + span / 2;
      const label = segmentLabel(decor.labelContent ?? ["category", "percent"], {
        value: v,
        fraction: v / total,
        series: data.series[0]?.name ?? "",
        category: data.categories[c],
        fmt,
      });
      const inside = span >= 30 && !doughnut;
      const p = polar(cx, cy, inside ? r * 0.62 : r + fs * 0.8, mid);
      const w = textWidth(label, fs) + 4;
      const rightHalf = mid % 360 < 180;
      nodes.push({
        kind: "text",
        x: inside ? p.x - w / 2 : rightHalf ? p.x : p.x - w,
        y: p.y - fs * 0.75,
        w,
        h: fs * 1.5,
        text: label,
        fontSize: fs,
        color: inside ? "#ffffff" : style.text,
        align: inside ? "center" : rightHalf ? "left" : "right",
        valign: "middle",
        name: `label-${c}`,
      });
    }
    angle += span;
  });

  if (doughnut) {
    nodes.push({ kind: "ellipse", cx, cy, rx: r * 0.55, ry: r * 0.55, fill: style.background, name: "hole" });
    nodes.push({
      kind: "text", x: cx - r * 0.5, y: cy - fs * 0.9, w: r, h: fs * 1.8,
      text: formatNumber(total, fmt), fontSize: fs * 1.3, bold: true, color: style.text,
      align: "center", valign: "middle", name: "hole-label",
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: values.map(() => cx),
      categoryWidth: values.map(() => r),
      columnTop: values.map(() => cy - r),
      columnValue: values,
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
    },
  };
}
