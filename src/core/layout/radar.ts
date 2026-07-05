import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { polar, textWidth, type SceneNode } from "../scene";
import { formatNumber, niceTicks, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { footnoteH } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Radar (spider) chart: categories = spokes (first at 12 o'clock,
 * clockwise), series = polygons with translucent fills. Gridlines are
 * straight polygons by default (business style); scale is shared across
 * all spokes and ticked on the 12 o'clock spoke only.
 */
export function layoutRadar(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const footH = footnoteH(cfg, style, decor);
  const legendH = decor.seriesLabels && data.series.length > 1 ? fs * 1.6 : 0;
  const cx = cfg.width / 2;
  const cy = titleH + legendH + (cfg.height - titleH - legendH - footH) / 2;
  // Perimeter labels need a margin around the web.
  const labelW = Math.max(0, ...data.categories.map((c) => textWidth(c, fs)));
  const r = Math.max(
    10,
    Math.min(cfg.width / 2 - labelW - fs, (cfg.height - titleH - legendH - footH) / 2 - fs * 1.9),
  );

  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const ticks = niceTicks(Math.min(0, cfg.scale?.min ?? 0), Math.max(cfg.scale?.max ?? Math.max(1, ...all), 1), 4);
  const min = cfg.scale?.min ?? ticks[0];
  const max = cfg.scale?.max ?? ticks[ticks.length - 1];
  const fmt = resolveFormat(ticks, cfg.numberFormat);
  const toR = (v: number) => ((v - min) / (max - min || 1)) * r;
  const angle = (c: number) => (360 / Math.max(1, n)) * c;

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }

  // Grid: concentric polygons (or circles) at each tick, plus the spokes.
  const gridShape = decor.gridShape ?? "polygon";
  for (const t of ticks) {
    if (t <= min) continue;
    const rr = toR(t);
    if (gridShape === "circle") {
      nodes.push({ kind: "ellipse", cx, cy, rx: rr, ry: rr, fill: "none", stroke: style.gridline, strokeWidth: 0.75, name: `grid-${t}` });
    } else {
      nodes.push({
        kind: "polygon",
        points: data.categories.map((_, c) => polar(cx, cy, rr, angle(c))),
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: `grid-${t}`,
      });
    }
    // Tick label on the 12 o'clock spoke only.
    nodes.push({
      kind: "text", x: cx + 3, y: cy - rr - fs * 0.6, w: fs * 3.4, h: fs * 1.2,
      text: formatNumber(t, fmt), fontSize: fs * 0.85, color: style.mutedText,
      align: "left", valign: "middle", name: `tick-${t}`,
    });
  }
  data.categories.forEach((cat, c) => {
    const end = polar(cx, cy, r, angle(c));
    nodes.push({ kind: "line", x1: cx, y1: cy, x2: end.x, y2: end.y, stroke: style.gridline, strokeWidth: 0.75, name: `spoke-${c}` });
    // Perimeter category label, anchored by which side of the web it sits on.
    const p = polar(cx, cy, r + fs * 0.6, angle(c));
    const a = angle(c) % 360;
    const align = a < 10 || a > 350 || Math.abs(a - 180) < 10 ? "center" : a < 180 ? "left" : "right";
    const w = textWidth(cat, fs) + 4;
    nodes.push({
      kind: "text",
      x: align === "center" ? p.x - w / 2 : align === "left" ? p.x : p.x - w,
      y: p.y - (a < 10 || a > 350 ? fs * 1.4 : Math.abs(a - 180) < 10 ? 0 : fs * 0.7),
      w,
      h: fs * 1.4,
      text: cat,
      fontSize: fs,
      color: style.text,
      align,
      valign: "middle",
      name: `category-${c}`,
    });
  });

  // Series polygons: translucent fill (SVG), full-opacity outline + markers.
  const defaultOpacity = data.series.length === 1 ? 0.25 : 0.18;
  data.series.forEach((s, si) => {
    const color = seriesColor(style, si, s.color);
    const pts = data.categories.map((_, c) => polar(cx, cy, toR(Math.max(min, s.values[c] ?? min)), angle(c)));
    nodes.push({
      kind: "polygon",
      points: pts,
      fill: color,
      fillOpacity: decor.fillOpacity ?? defaultOpacity,
      stroke: color,
      strokeWidth: 1.6,
      name: `series-${si}`,
    });
    pts.forEach((p, c) => {
      if (s.values[c] == null) return;
      nodes.push({ kind: "ellipse", cx: p.x, cy: p.y, rx: 2.4, ry: 2.4, fill: color, stroke: style.background, strokeWidth: 1, name: `marker-${si}-${c}` });
    });
  });

  // Legend row under the title when there are multiple series.
  if (legendH) {
    let x = 0;
    data.series.forEach((s, si) => {
      const chip = fs * 0.7;
      const color = seriesColor(style, si, s.color);
      nodes.push(
        { kind: "rect", x, y: titleH + fs * 0.35, w: chip, h: chip, fill: color, name: `legend-chip-${si}` },
        {
          kind: "text", x: x + chip + 3, y: titleH, w: textWidth(s.name, fs) + 6, h: fs * 1.4,
          text: s.name, fontSize: fs, color: style.text, align: "left", valign: "middle", name: `legend-${si}`,
        },
      );
      x += chip + 3 + textWidth(s.name, fs) + 12;
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).x),
      categoryWidth: data.categories.map(() => r / 2),
      columnTop: data.categories.map((_, c) => polar(cx, cy, r, angle(c)).y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: cy,
      plot: { x: cx - r, y: cy - r, w: r * 2, h: r * 2 },
    },
  };
}
