import type { ChartConfig, ChartStyle, Decorations } from "../types";
import type { SceneNode } from "../scene";
import { formatNumber, parseDateToken, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { baselineNode, categorySlots, chromeNodes, computeFrame, valueScale } from "./frame";
import { seriesLabelNodes, type LayoutResult } from "./column";

/** Line and area charts over categories. Lines are 2pt with ≥3pt markers. */
export function layoutLine(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const area = cfg.kind === "area";
  const fs = style.fontSize;

  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  // Area charts stack; lines share one scale.
  const stackedTotals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0),
  );
  const dataMax = area ? Math.max(0, ...stackedTotals) : Math.max(0, ...all);
  const dataMin = area ? 0 : Math.min(0, ...all);
  const fmt = resolveFormat(all, cfg.numberFormat);

  const { frame } = computeFrame(cfg, style, decor, decor.seriesLabels ? data.series.map((s) => s.name) : []);
  const slots = categorySlots(frame, n, 0);
  // Date categories ("2024-03", "Jan 2025", …) space x proportionally to time.
  const days = data.categories.map((c) => parseDateToken(c));
  if (!area && n > 1 && days.every((d): d is number => d != null)) {
    const d0 = Math.min(...(days as number[]));
    const d1 = Math.max(...(days as number[]));
    const inset = slots.slotWidth / 2;
    for (let c = 0; c < n; c++) {
      slots.centers[c] = frame.x + inset + (((days[c] as number) - d0) / (d1 - d0 || 1)) * (frame.w - inset * 2);
    }
  }
  const scale = valueScale(frame, dataMin, dataMax, cfg.scale, undefined, !area && cfg.logScale);
  const y0 = scale.toY(0);

  const nodes: SceneNode[] = chromeNodes(cfg, style, decor, frame, slots.centers, scale);
  const lastSegMid: (number | null)[] = data.series.map(() => null);
  const columnTop: number[] = data.categories.map(() => y0);

  if (area) {
    // Stacked areas drawn as per-category slabs (renderers have no polygon fill),
    // bottom-up so later series sit on top visually.
    const base = data.categories.map(() => 0);
    data.series.forEach((s, si) => {
      const fill = seriesColor(style, si, s.color);
      for (let c = 0; c < n - 1; c++) {
        const v0 = Math.max(0, s.values[c] ?? 0);
        const v1 = Math.max(0, s.values[c + 1] ?? 0);
        // Approximate the trapezoid with a rect at the average height.
        const yTop0 = scale.toY(base[c] + v0);
        const yTop1 = scale.toY(base[c + 1] + v1);
        const yBot0 = scale.toY(base[c]);
        const yBot1 = scale.toY(base[c + 1]);
        const steps = 24;
        const w = (slots.centers[c + 1] - slots.centers[c]) / steps;
        for (let k = 0; k < steps; k++) {
          const t = (k + 0.5) / steps;
          const yT = yTop0 + (yTop1 - yTop0) * t;
          const yB = yBot0 + (yBot1 - yBot0) * t;
          nodes.push({ kind: "rect", x: slots.centers[c] + k * w, y: yT, w: w + 0.5, h: Math.max(0, yB - yT), fill, name: `area-${si}-${c}-${k}` });
        }
      }
      for (let c = 0; c < n; c++) base[c] += Math.max(0, s.values[c] ?? 0);
      lastSegMid[si] = scale.toY(base[n - 1] - Math.max(0, s.values[n - 1] ?? 0) / 2);
    });
    for (let c = 0; c < n; c++) columnTop[c] = scale.toY(base[c]);
  } else {
    data.series.forEach((s, si) => {
      const color = seriesColor(style, si, s.color);
      let prev: { x: number; y: number } | null = null;
      for (let c = 0; c < n; c++) {
        const v = s.values[c];
        if (v == null) {
          prev = null;
          continue;
        }
        const pt = { x: slots.centers[c], y: scale.toY(v) };
        columnTop[c] = Math.min(columnTop[c], pt.y);
        if (prev) {
          nodes.push({ kind: "line", x1: prev.x, y1: prev.y, x2: pt.x, y2: pt.y, stroke: color, strokeWidth: 2, name: `line-${si}-${c}` });
        }
        // Marker: small square with a background ring so crossings stay legible.
        const r = 2.4;
        nodes.push({ kind: "rect", x: pt.x - r, y: pt.y - r, w: r * 2, h: r * 2, fill: color, stroke: style.background, strokeWidth: 1, name: `marker-${si}-${c}` });
        if (decor.segmentLabels) {
          nodes.push({
            kind: "text",
            x: pt.x - 30,
            y: pt.y - fs * 1.65,
            w: 60,
            h: fs * 1.4,
            text: formatNumber(v, fmt),
            fontSize: fs,
            color: style.text,
            align: "center",
            valign: "bottom",
            name: `label-${si}-${c}`,
          });
        }
        prev = pt;
        if (c === n - 1) lastSegMid[si] = pt.y;
      }
    });
  }

  nodes.push(baselineNode(frame, y0, style));
  if (decor.seriesLabels) nodes.push(...seriesLabelNodes(cfg, style, frame, lastSegMid));

  return {
    nodes,
    anchors: {
      categoryX: slots.centers,
      categoryWidth: data.categories.map(() => slots.colWidth || 10),
      columnTop,
      columnValue: area ? stackedTotals : data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: scale.toY,
    },
  };
}
