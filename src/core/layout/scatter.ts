import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatNumber, formatP, niceTicks, resolveFormat, trendStats } from "../format";
import { placeLabels, type Box, type LabelRequest } from "../labels";
import { PALETTE } from "../style";
import { sequentialScale } from "../color";
import { footnoteH } from "./frame";
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
  // A numeric "Color" row encodes a third/fourth variable on a sequential ramp.
  const colorVals = find(/^colou?r$/i)?.values ?? [];
  // Partition lines and trend line, think-cell's scatter decorations.
  const xLines = (find(/^x\s*line$/i)?.values ?? []).filter((v): v is number => v != null);
  const yLines = (find(/^y\s*line$/i)?.values ?? []).filter((v): v is number => v != null);
  const wantTrend = (find(/^trend$/i)?.values ?? []).some((v) => v != null);

  const pts = data.categories
    .map((label, i) => ({ label, x: xs[i], y: ys[i], size: sizes[i] ?? null, group: groups[i] ?? 1, color: colorVals[i] ?? null }))
    .filter((p): p is typeof p & { x: number; y: number } => p.x != null && p.y != null);

  // Continuous color scale (a "Color" row): maps each point onto a sequential
  // ramp; supersedes group coloring and swaps the chip legend for a gradient.
  const colorNums = pts.map((p) => p.color).filter((v): v is number => v != null);
  const colorScale =
    colorNums.length > 0
      ? { min: Math.min(...colorNums), max: Math.max(...colorNums), of: sequentialScale(Math.min(...colorNums), Math.max(...colorNums), (cfg.style?.palette ?? PALETTE)[0]) }
      : null;

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const axisW = 34;
  const multiGroup = !colorScale && new Set(pts.map((p) => Math.round(Number(p.group)))).size > 1;
  const legendH = multiGroup || colorScale ? fs * 1.8 : 0;
  const plot = {
    x: axisW,
    y: titleH + 6 + legendH,
    w: cfg.width - axisW - 8,
    h: cfg.height - titleH - 6 - legendH - fs * 1.6 - footnoteH(cfg, style, decor),
  };

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
  // Quadrant preset: one X/Y crossing → four tinted zones with corner
  // labels and the crossing lines — BCG-matrix framing in one step.
  if (decor.quadrants) {
    const { x: qx, y: qy, labels } = decor.quadrants;
    const cx = Math.max(plot.x, Math.min(plot.x + plot.w, toX(qx)));
    const cy = Math.max(plot.y, Math.min(plot.y + plot.h, toY(qy)));
    const zones: { x: number; y: number; w: number; h: number }[] = [
      { x: plot.x, y: plot.y, w: cx - plot.x, h: cy - plot.y }, // TL
      { x: cx, y: plot.y, w: plot.x + plot.w - cx, h: cy - plot.y }, // TR
      { x: plot.x, y: cy, w: cx - plot.x, h: plot.y + plot.h - cy }, // BL
      { x: cx, y: cy, w: plot.x + plot.w - cx, h: plot.y + plot.h - cy }, // BR
    ];
    zones.forEach((z, i) => {
      if (z.w <= 0 || z.h <= 0) return;
      // Checkerboard tint so adjacent zones read as distinct regions.
      nodes.push({ kind: "rect", ...z, fill: i === 0 || i === 3 ? "#f2f1ec" : "#faf9f6", name: `quadrant-${i}` });
      const label = labels?.[i];
      if (label) {
        nodes.push({
          kind: "text", x: z.x + 4, y: z.y + 2, w: Math.max(20, z.w - 8), h: fs * 1.3,
          text: label, fontSize: fs * 0.9, bold: true, color: style.mutedText,
          align: i === 1 || i === 3 ? "right" : "left", valign: "top", name: `quadrant-label-${i}`,
        });
      }
    });
    nodes.push(
      { kind: "line", x1: cx, y1: plot.y, x2: cx, y2: plot.y + plot.h, stroke: style.mutedText, strokeWidth: 1, dash: [3, 2], name: "quadrant-x" },
      { kind: "line", x1: plot.x, y1: cy, x2: plot.x + plot.w, y2: cy, stroke: style.mutedText, strokeWidth: 1, dash: [3, 2], name: "quadrant-y" },
    );
  }

  // Background bands (both axes in value units), behind gridlines and points.
  decor.bands?.forEach((band, i) => {
    const clampX = (v: number) => Math.max(plot.x, Math.min(plot.x + plot.w, toX(v)));
    const clampY = (v: number) => Math.max(plot.y, Math.min(plot.y + plot.h, toY(v)));
    const r =
      band.axis === "x"
        ? { x: Math.min(clampX(band.from), clampX(band.to)), y: plot.y, w: Math.abs(clampX(band.to) - clampX(band.from)), h: plot.h }
        : { x: plot.x, y: Math.min(clampY(band.from), clampY(band.to)), w: plot.w, h: Math.abs(clampY(band.to) - clampY(band.from)) };
    if (r.w <= 0 || r.h <= 0) return;
    nodes.push({ kind: "rect", ...r, fill: band.color ?? "#f2f1ec", name: `band-${i}` });
    if (band.label) {
      nodes.push({
        kind: "text", x: r.x + 3, y: r.y + 1, w: Math.max(20, r.w - 6), h: fs * 1.3,
        text: band.label, fontSize: fs * 0.9, color: style.mutedText, align: "left", valign: "top", name: `band-label-${i}`,
      });
    }
  });

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

  // Partition lines (dashed) at fixed x / y values.
  for (const v of xLines) {
    const x = toX(v);
    nodes.push({ kind: "line", x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: style.mutedText, strokeWidth: 1, dash: [3, 2], name: "x-line" });
  }
  for (const v of yLines) {
    const y = toY(v);
    nodes.push({ kind: "line", x1: plot.x, y1: y, x2: plot.x + plot.w, y2: y, stroke: style.mutedText, strokeWidth: 1, dash: [3, 2], name: "y-line" });
  }

  // OLS trend line across all points — always stating fit and significance.
  if (wantTrend && pts.length >= 2) {
    const mx = pts.reduce((s, p) => s + p.x, 0) / pts.length;
    const my = pts.reduce((s, p) => s + p.y, 0) / pts.length;
    const sxx = pts.reduce((s, p) => s + (p.x - mx) ** 2, 0);
    if (sxx > 0) {
      const slope = pts.reduce((s, p) => s + (p.x - mx) * (p.y - my), 0) / sxx;
      const at = (x: number) => my + slope * (x - mx);
      nodes.push({
        kind: "line",
        x1: toX(x0), y1: toY(at(x0)), x2: toX(x1), y2: toY(at(x1)),
        stroke: style.negative, strokeWidth: 1.25, dash: [4, 2], name: "trend",
      });
      const stats = trendStats(pts);
      if (stats) {
        const label = `R² = ${stats.r2.toFixed(2)}${stats.p != null ? `, p ${formatP(stats.p)}` : ""}`;
        const endY = toY(at(x1));
        nodes.push({
          kind: "text",
          x: plot.x + plot.w - textWidth(label, fs * 0.9) - 4,
          y: Math.max(plot.y, Math.min(plot.y + plot.h - fs * 1.3, endY + (slope >= 0 ? fs * 0.4 : -fs * 1.7))),
          w: textWidth(label, fs * 0.9) + 4,
          h: fs * 1.3,
          text: label,
          fontSize: fs * 0.9,
          color: style.negative,
          align: "right",
          valign: "middle",
          name: "trend-stats",
        });
      }
    }
  }

  // Continuous color legend: a discretized gradient bar with min/max labels
  // (renderer-safe — small rects, no SVG gradient).
  if (colorScale) {
    const cFmt = resolveFormat([colorScale.min, colorScale.max], cfg.numberFormat);
    const steps = 24;
    const barW = 90;
    const cell = barW / steps;
    const bx = plot.x;
    const by = plot.y - fs * 1.35;
    for (let i = 0; i < steps; i++) {
      const t = (i + 0.5) / steps;
      nodes.push({ kind: "rect", x: bx + i * cell, y: by, w: cell + 0.5, h: fs * 0.7, fill: colorScale.of(colorScale.min + t * (colorScale.max - colorScale.min)), name: `color-legend-${i}` });
    }
    const colorName = find(/^colou?r$/i)?.name ?? "Color";
    nodes.push(
      { kind: "text", x: bx - 40, y: by - fs * 0.15, w: 38, h: fs, text: colorName, fontSize: fs * 0.85, color: style.mutedText, align: "right", valign: "middle", name: "color-legend-title" },
      { kind: "text", x: bx, y: by + fs * 0.75, w: barW, h: fs, text: formatNumber(colorScale.min, cFmt), fontSize: fs * 0.75, color: style.mutedText, align: "left", valign: "middle", name: "color-legend-min" },
      { kind: "text", x: bx, y: by + fs * 0.75, w: barW, h: fs, text: formatNumber(colorScale.max, cFmt), fontSize: fs * 0.75, color: style.mutedText, align: "right", valign: "middle", name: "color-legend-max" },
    );
  }

  // Group legend when points are colored by group (skipped under a color scale).
  const groupIds = colorScale ? [] : [...new Set(pts.map((p) => Math.max(1, Math.round(Number(p.group)))))].sort((a, b) => a - b);
  if (groupIds.length > 1) {
    let lx = plot.x;
    for (const g of groupIds) {
      const chip = fs * 0.7;
      const label = `Group ${g}`;
      nodes.push(
        { kind: "rect", x: lx, y: plot.y - fs * 1.2, w: chip, h: chip, fill: (cfg.style?.palette ?? PALETTE)[(g - 1) % 8], name: `legend-chip-${g}` },
        { kind: "text", x: lx + chip + 3, y: plot.y - fs * 1.55, w: textWidth(label, fs) + 6, h: fs * 1.4, text: label, fontSize: fs, color: style.text, align: "left", valign: "middle", name: `legend-${g}` },
      );
      lx += chip + 3 + textWidth(label, fs) + 14;
    }
  }

  // Bubble radius: area ∝ size, max radius 9% of the smaller plot dimension.
  const maxSize = Math.max(1e-9, ...pts.map((p) => Math.abs(p.size ?? 0)));
  const maxR = Math.min(plot.w, plot.h) * 0.09;
  const radius = (p: (typeof pts)[number]) =>
    cfg.kind === "bubble" && p.size != null ? Math.max(2.5, Math.sqrt(Math.abs(p.size) / maxSize) * maxR) : 3;

  // Bubble size legend: without a key, bubble AREA is unreadable. Two
  // outline reference circles (a nice maximum and its half), top-right.
  const legendBoxes: Box[] = [];
  if (cfg.kind === "bubble" && pts.some((p) => p.size != null)) {
    const sizeFmt = resolveFormat(pts.map((p) => Math.abs(p.size ?? 0)), cfg.numberFormat);
    const refMax = niceTicks(0, maxSize, 3).pop()!;
    const refs = [refMax, refMax / 2];
    let lx = plot.x + plot.w - 4;
    refs.forEach((v, i) => {
      const r = Math.max(2.5, Math.sqrt(v / maxSize) * maxR);
      const cx = lx - r;
      const cy = plot.y + maxR * 1.1 + (Math.sqrt(refMax / maxSize) * maxR - r); // bottom-aligned circles
      nodes.push(
        { kind: "ellipse", cx, cy, rx: r, ry: r, fill: "none", stroke: style.mutedText, strokeWidth: 1, name: `size-legend-${i}` },
        {
          kind: "text", x: cx - r, y: cy - Math.sqrt(refMax / maxSize) * maxR - fs * 1.35, w: r * 2, h: fs * 1.2,
          text: formatNumber(v, sizeFmt), fontSize: fs * 0.8, color: style.mutedText,
          align: "center", valign: "bottom", name: `size-legend-label-${i}`,
        },
      );
      legendBoxes.push({ x: cx - r, y: cy - r - fs * 1.4, w: r * 2, h: r * 2 + fs * 1.4 });
      lx = cx - r - fs * 0.8;
    });
  }

  // Trajectory / trail: connect the points in datasheet (row) order with a
  // direction arrowhead at each segment midpoint, drawn behind the markers —
  // a Gapminder-style path of one entity through the X/Y space over time.
  if (decor.trajectory && pts.length > 1) {
    for (let i = 0; i < pts.length - 1; i++) {
      const ax = toX(pts[i].x);
      const ay = toY(pts[i].y);
      const bx = toX(pts[i + 1].x);
      const by = toY(pts[i + 1].y);
      nodes.push({ kind: "line", x1: ax, y1: ay, x2: bx, y2: by, stroke: style.mutedText, strokeWidth: 1.5, name: `trajectory-${i}` });
      const angle = (Math.atan2(by - ay, bx - ax) * 180) / Math.PI;
      nodes.push({ kind: "arrowhead", x: (ax + bx) / 2, y: (ay + by) / 2, angle, size: 4, fill: style.mutedText, name: `trajectory-head-${i}` });
    }
  }

  // Point labels treat the size legend as an obstacle.
  const markerBoxes: Box[] = [...legendBoxes];
  // Paint back-to-front. Emitting in datasheet order lets a large bubble drawn
  // early bury a smaller one drawn later — completely, in all three renderers,
  // with nothing to tell the reader a point is missing. Largest first puts the
  // big ones at the back. This is paint order only: every marker keeps its
  // datasheet index in its name and its exact position. Ties keep datasheet
  // order (Array.sort is stable), so the result stays deterministic. Labels
  // already run biggest-first below, for the same reason.
  const paintOrder = pts.map((_, i) => i).sort((a, b) => radius(pts[b]) - radius(pts[a]));
  for (const i of paintOrder) {
    const p = pts[i];
    const r = radius(p);
    const gi = Math.max(0, Math.round(Number(p.group)) - 1);
    const fill =
      colorScale && p.color != null ? colorScale.of(p.color) : colorScale ? style.mutedText : (cfg.style?.palette ?? PALETTE)[gi % 8];
    nodes.push({
      kind: "ellipse",
      cx: toX(p.x),
      cy: toY(p.y),
      rx: r,
      ry: r,
      fill,
      stroke: style.background,
      strokeWidth: 1,
      name: `point-${i}`,
    });
    markerBoxes.push({ x: toX(p.x) - r, y: toY(p.y) - r, w: r * 2, h: r * 2 });
  }

  // Greedy label placement, biggest bubbles first so important points win.
  if (decor.segmentLabels !== false) {
    const order = pts.map((_, i) => i).sort((a, b) => radius(pts[b]) - radius(pts[a]));
    // Point label content: category (default), optionally with "(x, y)".
    const pointLabel = (p: (typeof pts)[number]) => {
      const parts = decor.labelContent ?? ["category"];
      const fmt = resolveFormat([p.x, p.y], cfg.numberFormat);
      return parts
        .map((part) =>
          part === "category" ? p.label : part === "value" ? `(${formatNumber(p.x, fmt)}, ${formatNumber(p.y, fmt)})` : null,
        )
        .filter(Boolean)
        .join(" ");
    };
    const reqs: LabelRequest[] = order.map((i) => ({
      cx: toX(pts[i].x),
      cy: toY(pts[i].y),
      r: radius(pts[i]),
      w: textWidth(pointLabel(pts[i]), fs) + 2,
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
        text: pointLabel(p),
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
