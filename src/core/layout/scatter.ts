import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatNumber, niceTicks, resolveFormat } from "../format";
import { placeLabels, type Box, type LabelRequest } from "../labels";
import { PALETTE } from "../style";
import type { LayoutResult } from "./column";

/**
 * Scatter / bubble chart, think-cell datasheet convention: categories are
 * point labels; rows named X, Y (and Size for bubbles) carry coordinates.
 * A row named Group (values 1..k) colors points by group.
 * Point labels are placed by the greedy collision-avoiding placer and hidden
 * when the chart gets too dense.
 */
export function layoutScatter(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()));
  const xs = find(/^x$/i)?.values ?? [];
  const ys = find(/^y$/i)?.values ?? [];
  const sizes = cfg.kind === "bubble" ? (find(/^size$/i)?.values ?? []) : [];
  const groups = find(/^group$/i)?.values ?? [];

  const pts = data.categories
    .map((label, i) => ({ label, x: xs[i], y: ys[i], size: sizes[i] ?? null, group: groups[i] ?? 1 }))
    .filter((p): p is typeof p & { x: number; y: number } => p.x != null && p.y != null);

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const axisW = 34;
  const plot = { x: axisW, y: titleH + 6, w: cfg.width - axisW - 8, h: cfg.height - titleH - 6 - fs * 1.6 };

  const xTicks = niceTicks(Math.min(0, ...pts.map((p) => p.x)), Math.max(1, ...pts.map((p) => p.x)), 5);
  const yTicks = niceTicks(Math.min(0, ...pts.map((p) => p.y)), Math.max(1, ...pts.map((p) => p.y)), 5);
  const x0 = xTicks[0];
  const x1 = xTicks[xTicks.length - 1];
  const y0 = yTicks[0];
  const y1 = yTicks[yTicks.length - 1];
  const toX = (v: number) => plot.x + ((v - x0) / (x1 - x0 || 1)) * plot.w;
  const toY = (v: number) => plot.y + plot.h - ((v - y0) / (y1 - y0 || 1)) * plot.h;

  const xFmt = resolveFormat(xTicks, cfg.numberFormat);
  const yFmt = resolveFormat(yTicks, cfg.numberFormat);

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }
  // Gridlines + axis labels on both axes.
  for (const t of yTicks) {
    const y = toY(t);
    nodes.push(
      { kind: "line", x1: plot.x, y1: y, x2: plot.x + plot.w, y2: y, stroke: style.gridline, strokeWidth: 0.75, name: "gridline-y" },
      { kind: "text", x: 0, y: y - fs * 0.7, w: plot.x - 4, h: fs * 1.4, text: formatNumber(t, yFmt), fontSize: fs * 0.9, color: style.mutedText, align: "right", valign: "middle", name: "y-axis" },
    );
  }
  for (const t of xTicks) {
    const x = toX(t);
    nodes.push({
      kind: "text", x: x - 24, y: plot.y + plot.h + 2, w: 48, h: fs * 1.4,
      text: formatNumber(t, xFmt), fontSize: fs * 0.9, color: style.mutedText, align: "center", valign: "top", name: "x-axis",
    });
  }
  nodes.push(
    { kind: "line", x1: plot.x, y1: plot.y + plot.h, x2: plot.x + plot.w, y2: plot.y + plot.h, stroke: style.axis, strokeWidth: 1, name: "baseline" },
    { kind: "line", x1: toX(0) >= plot.x ? toX(0) : plot.x, y1: plot.y, x2: toX(0) >= plot.x ? toX(0) : plot.x, y2: plot.y + plot.h, stroke: style.axis, strokeWidth: 1, name: "y-axis-line" },
  );

  // Bubble radius: area ∝ size, max radius 9% of the smaller plot dimension.
  const maxSize = Math.max(1e-9, ...pts.map((p) => Math.abs(p.size ?? 0)));
  const maxR = Math.min(plot.w, plot.h) * 0.09;
  const radius = (p: (typeof pts)[number]) =>
    cfg.kind === "bubble" && p.size != null ? Math.max(2.5, Math.sqrt(Math.abs(p.size) / maxSize) * maxR) : 3;

  const markerBoxes: Box[] = [];
  pts.forEach((p, i) => {
    const r = radius(p);
    const gi = Math.max(0, Math.round(Number(p.group)) - 1);
    nodes.push({
      kind: "ellipse",
      cx: toX(p.x),
      cy: toY(p.y),
      rx: r,
      ry: r,
      fill: (cfg.style?.palette ?? PALETTE)[gi % 8],
      stroke: style.background,
      strokeWidth: 1,
      name: `point-${i}`,
    });
    markerBoxes.push({ x: toX(p.x) - r, y: toY(p.y) - r, w: r * 2, h: r * 2 });
  });

  // Greedy label placement, biggest bubbles first so important points win.
  if (decor.segmentLabels !== false) {
    const order = pts.map((_, i) => i).sort((a, b) => radius(pts[b]) - radius(pts[a]));
    const reqs: LabelRequest[] = order.map((i) => ({
      cx: toX(pts[i].x),
      cy: toY(pts[i].y),
      r: radius(pts[i]),
      w: textWidth(pts[i].label, fs) + 2,
      h: fs * 1.3,
    }));
    for (const placed of placeLabels(reqs, { x: 0, y: plot.y, w: cfg.width, h: plot.h + fs * 1.5 }, markerBoxes)) {
      const p = pts[order[placed.index]];
      nodes.push({
        kind: "text",
        x: placed.box.x,
        y: placed.box.y,
        w: placed.box.w,
        h: placed.box.h,
        text: p.label,
        fontSize: fs,
        color: style.text,
        align: "left",
        valign: "middle",
        name: `label-${order[placed.index]}`,
      });
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: pts.map((p) => toX(p.x)),
      categoryWidth: pts.map((p) => radius(p) * 2),
      columnTop: pts.map((p) => toY(p.y)),
      columnValue: pts.map((p) => p.y),
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
