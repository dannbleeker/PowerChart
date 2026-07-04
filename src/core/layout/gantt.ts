import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatNumber, niceTicks, resolveFormat } from "../format";
import { seriesColor } from "../style";
import type { LayoutResult } from "./column";

/**
 * Simplified Gantt / timeline: categories are activities; rows named
 * Start and End give each activity's span on a numeric timeline (week,
 * month index, year — any number). A row named Milestone adds a diamond
 * marker at that position. think-cell's calendar-based Gantt is richer;
 * this covers the project-on-a-slide case.
 */
export function layoutGantt(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()));
  const starts = find(/^start$/i)?.values ?? [];
  const ends = find(/^end$/i)?.values ?? [];
  const milestones = find(/^milestone$/i)?.values ?? [];

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const headerH = fs * 1.6;
  const catW = Math.min(
    cfg.width * 0.32,
    Math.max(0, ...data.categories.map((c) => textWidth(c, fs))) + 10,
  );
  const plot = { x: catW, y: titleH + headerH, w: cfg.width - catW - 6, h: cfg.height - titleH - headerH - 6 };

  const all = [...starts, ...ends, ...milestones].filter((v): v is number => v != null);
  const ticks = niceTicks(Math.min(...(all.length ? all : [0])), Math.max(...(all.length ? all : [1])), 6);
  const t0 = ticks[0];
  const t1 = ticks[ticks.length - 1];
  const toX = (v: number) => plot.x + ((v - t0) / (t1 - t0 || 1)) * plot.w;
  const fmt = resolveFormat(ticks, cfg.numberFormat);

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }
  // Timeline header on top + vertical gridlines (think-cell's calendar strip).
  for (const t of ticks) {
    const x = toX(t);
    nodes.push(
      { kind: "text", x: x - 24, y: titleH, w: 48, h: headerH, text: formatNumber(t, fmt), fontSize: fs * 0.9, color: style.mutedText, align: "center", valign: "middle", name: "timeline" },
      { kind: "line", x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: style.gridline, strokeWidth: 0.75, name: "gridline" },
    );
  }

  const slotH = plot.h / Math.max(1, data.categories.length);
  const barH = Math.min(slotH * 0.55, fs * 1.4);
  const columnTop: number[] = [];

  data.categories.forEach((activity, c) => {
    const cy = plot.y + slotH * (c + 0.5);
    columnTop.push(cy - barH / 2);
    nodes.push({
      kind: "text", x: 0, y: cy - fs * 0.75, w: catW - 6, h: fs * 1.5,
      text: activity, fontSize: fs, color: style.text, align: "left", valign: "middle", name: `category-${c}`,
    });
    // Faint row separator.
    if (c > 0) {
      nodes.push({ kind: "line", x1: plot.x, y1: cy - slotH / 2, x2: plot.x + plot.w, y2: cy - slotH / 2, stroke: style.gridline, strokeWidth: 0.5, name: `row-${c}` });
    }
    const s = starts[c];
    const e = ends[c];
    if (s != null && e != null && e > s) {
      nodes.push({
        kind: "rect", x: toX(s), y: cy - barH / 2, w: toX(e) - toX(s), h: barH,
        fill: seriesColor(style, 0), name: `bar-${c}`,
      });
      if (decor.segmentLabels) {
        const label = `${formatNumber(s, fmt)}–${formatNumber(e, fmt)}`;
        if (toX(e) - toX(s) >= textWidth(label, fs * 0.9) + 4) {
          nodes.push({
            kind: "text", x: toX(s), y: cy - fs * 0.7, w: toX(e) - toX(s), h: fs * 1.4,
            text: label, fontSize: fs * 0.9, color: "#ffffff", align: "center", valign: "middle", name: `bar-label-${c}`,
          });
        }
      }
    }
    const m = milestones[c];
    if (m != null) {
      const r = barH * 0.45;
      // Diamond milestone marker.
      nodes.push({
        kind: "ellipse", cx: toX(m), cy, rx: r, ry: r,
        fill: style.text, name: `milestone-${c}`,
      });
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.y + slotH * (c + 0.5)),
      categoryWidth: data.categories.map(() => barH),
      columnTop,
      columnValue: data.categories.map((_, c) => ends[c] ?? 0),
      baselineY: plot.x,
      plot,
    },
  };
}
