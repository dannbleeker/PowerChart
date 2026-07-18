import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatNumber, parseDateToken, resolveFormat, segmentLabel } from "../format";
import { seriesColor } from "../style";
import { lerpColor } from "../color";
import {
  baselineNode,
  categorySlots,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  footnoteH,
  titleHeight,
  titleNode,
  valueScale,
} from "./frame";
import { horizontalChrome, seriesLabelNodes, type LayoutResult } from "./column";

/**
 * Slab count for a filled area/ribbon segment.
 *
 * The renderers have no polygon fill, so a segment is tiled with thin rects.
 * This was a flat 24 regardless of on-screen width, which put sub-pixel rects on
 * narrow charts and multiplied the Office.js shape count ~24× per segment (a
 * 30×5 stacked area is ~3,500 rects, right against the host's shape budget).
 * Scale to the segment's pixel span (~4px per slab, the eye can't resolve a
 * finer staircase) and cap at the old 24 so a wide segment never GAINS shapes;
 * a stepped (staircase) segment has a flat top and needs no tessellation.
 */
function slabSteps(spanPx: number, stepped?: "before" | "after" | "center"): number {
  if (stepped) return stepped === "center" ? 2 : 1;
  return Math.max(2, Math.min(24, Math.ceil(Math.abs(spanPx) / 4)));
}

/** Line and area charts over categories. Lines are 2pt with ≥3pt markers. */
export function layoutLine(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  if ((cfg.kind === "line" || cfg.kind === "area") && decor.sparkline) {
    return layoutSparkline(cfg, style, decor);
  }
  if (cfg.kind === "line" && decor.slope && cfg.data.categories.length >= 2) {
    return layoutSlope(cfg, style, decor);
  }
  if (cfg.kind === "line" && decor.bump && cfg.data.categories.length >= 2) {
    return layoutBump(cfg, style, decor);
  }
  if ((cfg.kind === "line" || cfg.kind === "area") && cfg.horizontal) {
    return layoutLineHorizontal(cfg, style, decor);
  }
  const rawData = cfg.data;
  // Band low/high rows shade an uncertainty ribbon instead of drawing lines.
  const BAND_LOW = /^band\s*low$/i;
  const BAND_HIGH = /^band\s*high$/i;
  const bandLow = rawData.series.find((s) => BAND_LOW.test(s.name.trim()))?.values;
  const bandHigh = rawData.series.find((s) => BAND_HIGH.test(s.name.trim()))?.values;
  const data = {
    ...rawData,
    series: rawData.series.filter((s) => !BAND_LOW.test(s.name.trim()) && !BAND_HIGH.test(s.name.trim())),
  };
  const n = data.categories.length;
  const area = cfg.kind === "area";
  const fs = style.fontSize;

  const all = [
    ...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)),
    ...(bandLow ?? []).filter((v): v is number => v != null),
    ...(bandHigh ?? []).filter((v): v is number => v != null),
  ];
  // Area charts stack (positives above zero, negatives below); lines share
  // one scale. Negative areas dip under the baseline — think-cell parity for
  // P&L-over-time where a series can go negative.
  const stackedPos = data.categories.map((_, c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
  const stackedNeg = data.categories.map((_, c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
  const dataMax = area ? Math.max(0, ...stackedPos) : Math.max(0, ...all);
  const dataMin = area ? Math.min(0, ...stackedNeg) : Math.min(0, ...all);
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

  /** Shaded ribbon between two per-category value arrays (slab technique). */
  const ribbon = (lows: (number | null)[], highs: (number | null)[], fill: string, name: string) => {
    for (let c = 0; c < n - 1; c++) {
      const l0 = lows[c];
      const l1 = lows[c + 1];
      const h0 = highs[c];
      const h1 = highs[c + 1];
      if (l0 == null || l1 == null || h0 == null || h1 == null) continue;
      const span = slots.centers[c + 1] - slots.centers[c];
      const steps = slabSteps(span);
      const w = span / steps;
      for (let k = 0; k < steps; k++) {
        const t = (k + 0.5) / steps;
        const yT = scale.toY(h0 + (h1 - h0) * t);
        const yB = scale.toY(l0 + (l1 - l0) * t);
        nodes.push({
          kind: "rect",
          x: slots.centers[c] + k * w,
          y: Math.min(yT, yB),
          w: w + 0.5,
          h: Math.abs(yB - yT),
          fill,
          name: `${name}-${c}-${k}`,
        });
      }
    }
  };
  // Confidence/uncertainty band from Band low / Band high rows.
  if (!area && bandLow && bandHigh) {
    ribbon(bandLow, bandHigh, lerpColor("#ffffff", seriesColor(style, 0, data.series[0]?.color), 0.18), "band-ribbon");
  }
  // Filled gap between two named series (plan-vs-actual ribbon).
  if (!area && decor.fillBetween) {
    const [ai, bi] = decor.fillBetween;
    const sa = data.series[ai]?.values;
    const sb = data.series[bi]?.values;
    if (sa && sb)
      ribbon(sa, sb, lerpColor("#ffffff", seriesColor(style, ai, data.series[ai]?.color), 0.22), "fill-between");
  }

  if (area) {
    // Stacked areas drawn as per-category slabs (renderers have no polygon fill),
    // bottom-up so later series sit on top visually. Positives stack above the
    // zero baseline and negatives below it, so a negative series dips under zero.
    const posBase = data.categories.map(() => 0);
    const negBase = data.categories.map(() => 0);
    data.series.forEach((s, si) => {
      const fill = seriesColor(style, si, s.color);
      // This series' band boundaries per category (value units).
      const lower: number[] = [];
      const upper: number[] = [];
      for (let c = 0; c < n; c++) {
        const v = s.values[c] ?? 0;
        if (v >= 0) {
          lower[c] = posBase[c];
          upper[c] = posBase[c] + v;
          posBase[c] += v;
        } else {
          upper[c] = negBase[c];
          lower[c] = negBase[c] + v;
          negBase[c] += v;
        }
      }
      for (let c = 0; c < n - 1; c++) {
        // Approximate the trapezoid with a rect at the average height.
        const yTop0 = scale.toY(upper[c]);
        const yTop1 = scale.toY(upper[c + 1]);
        const yBot0 = scale.toY(lower[c]);
        const yBot1 = scale.toY(lower[c + 1]);
        const span = slots.centers[c + 1] - slots.centers[c];
        const steps = slabSteps(span, decor.stepped);
        const w = span / steps;
        for (let k = 0; k < steps; k++) {
          const t = (k + 0.5) / steps;
          // Stepped areas hold a flat top across the interval (staircase);
          // "after" carries the left value, "before" the right, "center" both.
          let yT: number;
          let yB: number;
          if (decor.stepped === "after" || (decor.stepped === "center" && t < 0.5)) {
            yT = yTop0;
            yB = yBot0;
          } else if (decor.stepped === "before" || decor.stepped === "center") {
            yT = yTop1;
            yB = yBot1;
          } else {
            yT = yTop0 + (yTop1 - yTop0) * t;
            yB = yBot0 + (yBot1 - yBot0) * t;
          }
          nodes.push({
            kind: "rect",
            x: slots.centers[c] + k * w,
            y: Math.min(yT, yB),
            w: w + 0.5,
            h: Math.abs(yB - yT),
            fill,
            name: `area-${si}-${c}-${k}`,
          });
        }
      }
      lastSegMid[si] = n > 0 ? scale.toY((lower[n - 1] + upper[n - 1]) / 2) : null;
    });
    for (let c = 0; c < n; c++) columnTop[c] = scale.toY(posBase[c]);
  } else {
    // Forecast boundary: categories from this index on draw dashed with
    // hollow markers; a subtle divider marks the actuals/forecast split.
    const fc = decor.forecastFrom;
    if (fc != null && fc > 0 && fc < n) {
      const dx = slots.centers[fc - 1] + (slots.centers[fc] - slots.centers[fc - 1]) / 2;
      nodes.push({
        kind: "line",
        x1: dx,
        y1: frame.y,
        x2: dx,
        y2: frame.y + frame.h,
        stroke: style.gridline,
        strokeWidth: 1,
        dash: [2, 3],
        name: "forecast-divider",
      });
    }
    // Smooth (Catmull-Rom) curves, sampled to a dense polyline. Ignored when
    // stepped is set (mutually exclusive shapes).
    const smooth = !!decor.smooth && !decor.stepped;
    data.series.forEach((s, si) => {
      const color = seriesColor(style, si, s.color);
      if (smooth) {
        // Split into contiguous runs (nulls break the line), then draw each
        // run as a Catmull-Rom spline sampled at STEPS points per segment.
        const seq = s.values.map((v, c) => (v == null ? null : { x: slots.centers[c], y: scale.toY(v), c }));
        const runs: { x: number; y: number; c: number }[][] = [];
        let cur: { x: number; y: number; c: number }[] = [];
        for (const p of seq) {
          if (p) cur.push(p);
          else if (cur.length) {
            runs.push(cur);
            cur = [];
          }
        }
        if (cur.length) runs.push(cur);
        const STEPS = 16;
        for (const run of runs) {
          for (let i = 0; i < run.length - 1; i++) {
            const p0 = run[Math.max(0, i - 1)];
            const p1 = run[i];
            const p2 = run[i + 1];
            const p3 = run[Math.min(run.length - 1, i + 2)];
            const forecast = fc != null && p2.c >= fc;
            let pp = { x: p1.x, y: p1.y };
            for (let k = 1; k <= STEPS; k++) {
              const t = k / STEPS;
              const t2 = t * t;
              const t3 = t2 * t;
              const cx =
                0.5 *
                (2 * p1.x +
                  (-p0.x + p2.x) * t +
                  (2 * p0.x - 5 * p1.x + 4 * p2.x - p3.x) * t2 +
                  (-p0.x + 3 * p1.x - 3 * p2.x + p3.x) * t3);
              const cy =
                0.5 *
                (2 * p1.y +
                  (-p0.y + p2.y) * t +
                  (2 * p0.y - 5 * p1.y + 4 * p2.y - p3.y) * t2 +
                  (-p0.y + 3 * p1.y - 3 * p2.y + p3.y) * t3);
              nodes.push({
                kind: "line",
                x1: pp.x,
                y1: pp.y,
                x2: cx,
                y2: cy,
                stroke: color,
                strokeWidth: 2,
                ...(forecast ? { dash: [4, 3] } : {}),
                name: `line-${si}-${p2.c}-s${k}`,
              });
              pp = { x: cx, y: cy };
            }
          }
        }
      }
      let prev: { x: number; y: number } | null = null;
      for (let c = 0; c < n; c++) {
        const v = s.values[c];
        if (v == null) {
          // Bridge gaps: keep the previous point so the next value connects
          // straight across the missing categories instead of breaking.
          if (!decor.bridgeGaps) prev = null;
          continue;
        }
        const forecast = fc != null && c >= fc;
        const pt = { x: slots.centers[c], y: scale.toY(v) };
        columnTop[c] = Math.min(columnTop[c], pt.y);
        if (prev && !smooth) {
          const p = prev;
          const dashOpt = forecast ? { dash: [4, 3] } : {};
          const seg = (x1: number, y1: number, x2: number, y2: number, suffix: string) =>
            nodes.push({
              kind: "line",
              x1,
              y1,
              x2,
              y2,
              stroke: color,
              strokeWidth: 2,
              ...dashOpt,
              name: `line-${si}-${c}${suffix}`,
            });
          if (decor.stepped === "after") {
            seg(p.x, p.y, pt.x, p.y, "a");
            seg(pt.x, p.y, pt.x, pt.y, "b");
          } else if (decor.stepped === "before") {
            seg(p.x, p.y, p.x, pt.y, "a");
            seg(p.x, pt.y, pt.x, pt.y, "b");
          } else if (decor.stepped === "center") {
            const mx = (p.x + pt.x) / 2;
            seg(p.x, p.y, mx, p.y, "a");
            seg(mx, p.y, mx, pt.y, "b");
            seg(mx, pt.y, pt.x, pt.y, "c");
          } else {
            seg(p.x, p.y, pt.x, pt.y, "");
          }
        }
        // Marker: small square with a background ring so crossings stay legible.
        // A per-cell color override highlights the point (max/min/last…) with
        // a larger, recolored marker. Forecast points render hollow.
        const cellColor = s.colors?.[c];
        const r = cellColor ? 3.4 : 2.4;
        nodes.push({
          kind: "rect",
          x: pt.x - r,
          y: pt.y - r,
          w: r * 2,
          h: r * 2,
          fill: forecast && !cellColor ? style.background : (cellColor ?? color),
          stroke: forecast && !cellColor ? color : style.background,
          strokeWidth: 1,
          name: `marker-${si}-${c}`,
        });
        if (decor.segmentLabels) {
          nodes.push({
            kind: "text",
            x: pt.x - 30,
            y: pt.y - fs * 1.65,
            w: 60,
            h: fs * 1.4,
            text: segmentLabel(decor.labelContent ?? ["value"], {
              value: v,
              fraction: null,
              series: s.name,
              category: data.categories[c],
              fmt,
            }),
            fontSize: fs,
            bold: !!cellColor,
            color: cellColor ?? style.text,
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
      columnValue: area
        ? data.categories.map((_, c) => stackedPos[c] + stackedNeg[c])
        : data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: scale.toY,
    },
  };
}

/**
 * Sparkline: a compact, chrome-less trend line sized to sit inline (think
 * Tufte's "word-sized graphic"). No axes, gridlines or category labels — just
 * a thin line, an optional leading label (the title / series name) and a
 * trailing value, with dots on the min (red), max (green) and last points.
 * Pair with `multiples` to get a table of sparklines, one per series.
 */
function layoutSparkline(cfg: ChartConfig, style: ChartStyle, _decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const area = cfg.kind === "area";
  const single = data.series.length === 1;
  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const fmt = resolveFormat(all, cfg.numberFormat);
  const lo = all.length ? Math.min(...all) : 0;
  const hi = all.length ? Math.max(...all) : 1;
  const span = hi - lo || 1;

  // Leading label (title/series name) and trailing last value reserve gutters.
  const label = cfg.title ?? (single ? data.series[0].name : "");
  const labelW = label ? Math.min(cfg.width * 0.38, textWidth(label, fs) + 8) : 0;
  const lastVal = single ? ([...data.series[0].values].reverse().find((v): v is number => v != null) ?? null) : null;
  const endText = lastVal != null ? formatNumber(lastVal, fmt) : "";
  const endW = endText ? textWidth(endText, fs) + 8 : 0;
  const padY = Math.max(2, cfg.height * 0.16);
  const plot = {
    x: labelW,
    y: padY,
    w: Math.max(10, cfg.width - labelW - endW - 4),
    h: Math.max(4, cfg.height - padY * 2),
  };
  const xs = data.categories.map((_, c) => plot.x + (n === 1 ? plot.w / 2 : (c / (n - 1)) * plot.w));
  const toY = (v: number) => plot.y + (1 - (v - lo) / span) * plot.h;

  const nodes: SceneNode[] = [];
  if (label) {
    nodes.push({
      kind: "text",
      x: 0,
      y: plot.y + plot.h / 2 - fs * 0.75,
      w: labelW - 6,
      h: fs * 1.5,
      text: label,
      fontSize: fs,
      color: style.text,
      align: "left",
      valign: "middle",
      name: "spark-label",
    });
  }

  data.series.forEach((s, si) => {
    const color = seriesColor(style, si, s.color);
    const pts = s.values
      .map((v, c) => (v == null ? null : { x: xs[c], y: toY(v), v, c }))
      .filter((p): p is { x: number; y: number; v: number; c: number } => p != null);
    // Light area fill under the line (area kind only), per-segment rects to the floor.
    if (area) {
      const fill = lerpColor("#ffffff", color, 0.16);
      const floor = plot.y + plot.h;
      for (let i = 0; i < pts.length - 1; i++) {
        const span = pts[i + 1].x - pts[i].x;
        const steps = slabSteps(span);
        const w = span / steps;
        for (let k = 0; k < steps; k++) {
          const t = (k + 0.5) / steps;
          const y = pts[i].y + (pts[i + 1].y - pts[i].y) * t;
          nodes.push({
            kind: "rect",
            x: pts[i].x + k * w,
            y,
            w: w + 0.5,
            h: Math.max(0, floor - y),
            fill,
            name: `spark-fill-${si}-${i}-${k}`,
          });
        }
      }
    }
    for (let i = 0; i < pts.length - 1; i++) {
      nodes.push({
        kind: "line",
        x1: pts[i].x,
        y1: pts[i].y,
        x2: pts[i + 1].x,
        y2: pts[i + 1].y,
        stroke: color,
        strokeWidth: 1.25,
        name: `spark-${si}-${pts[i + 1].c}`,
      });
    }
    // Min / max / last dots (single-series only, to stay uncluttered).
    if (single && pts.length) {
      let minP = pts[0];
      let maxP = pts[0];
      for (const p of pts) {
        if (p.v < minP.v) minP = p;
        if (p.v > maxP.v) maxP = p;
      }
      const last = pts[pts.length - 1];
      nodes.push(
        { kind: "ellipse", cx: minP.x, cy: minP.y, rx: 1.9, ry: 1.9, fill: style.negative, name: `spark-min-${si}` },
        { kind: "ellipse", cx: maxP.x, cy: maxP.y, rx: 1.9, ry: 1.9, fill: "#1a9e6e", name: `spark-max-${si}` },
        { kind: "ellipse", cx: last.x, cy: last.y, rx: 2.1, ry: 2.1, fill: color, name: `spark-last-${si}` },
      );
      if (endText) {
        nodes.push({
          kind: "text",
          x: plot.x + plot.w + 4,
          y: last.y - fs * 0.75,
          w: endW,
          h: fs * 1.5,
          text: endText,
          fontSize: fs,
          bold: true,
          color,
          align: "left",
          valign: "middle",
          name: `spark-end-${si}`,
        });
      }
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: xs,
      categoryWidth: data.categories.map(() => plot.w / Math.max(1, n)),
      columnTop: data.categories.map(() => plot.y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: plot.y + plot.h,
      plot,
      valueToY: toY,
    },
  };
}

/**
 * Slope chart: the before/after comparison. No value axis or gridlines —
 * two vertical hairlines carry the periods, every series is a straight
 * line (or polyline for >2 categories) with a "Name value" label at both
 * ends, colored like its line. Labels de-overlap vertically per side.
 */
function layoutSlope(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const last = n - 1;
  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const fmt = resolveFormat(all, cfg.numberFormat);
  const lo = all.length ? Math.min(...all) : 0;
  const hi = all.length ? Math.max(...all) : 1;
  const span = hi - lo || 1;

  const endLabel = (s: (typeof data.series)[number], c: number) =>
    s.values[c] == null ? "" : `${s.name} ${formatNumber(s.values[c]!, fmt)}`;
  const gutterL = Math.min(
    cfg.width * 0.34,
    Math.max(fs, ...data.series.map((s) => textWidth(endLabel(s, 0), fs))) + 10,
  );
  const gutterR = Math.min(
    cfg.width * 0.34,
    Math.max(fs, ...data.series.map((s) => textWidth(endLabel(s, last), fs))) + 10,
  );

  const titleH = titleHeight(cfg, style);
  const headerH = fs * 1.5; // period labels above the rails
  const plot = {
    x: gutterL,
    y: titleH + headerH + 4,
    w: cfg.width - gutterL - gutterR,
    h: cfg.height - titleH - headerH - 4 - footnoteH(cfg, style, decor) - 6,
  };
  const pad = plot.h * 0.08;
  const toY = (v: number) => plot.y + pad + (1 - (v - lo) / span) * (plot.h - pad * 2);
  const xs = data.categories.map((_, c) => plot.x + (n === 1 ? plot.w / 2 : (c / (n - 1)) * plot.w));

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // Rails and period labels at the two ends only.
  for (const c of [0, last]) {
    nodes.push({
      kind: "line",
      x1: xs[c],
      y1: plot.y,
      x2: xs[c],
      y2: plot.y + plot.h,
      stroke: style.gridline,
      strokeWidth: 1,
      name: `slope-rail-${c}`,
    });
    nodes.push({
      kind: "text",
      x: xs[c] - 60,
      y: titleH,
      w: 120,
      h: headerH,
      text: data.categories[c],
      fontSize: fs,
      color: style.mutedText,
      align: "center",
      valign: "middle",
      name: `category-${c}`,
    });
  }

  // Per-side label placement: keep each label at its line end, then push
  // apart to a minimum gap and clamp back inside the plot.
  const place = (ys: (number | null)[]): (number | null)[] => {
    const idx = ys
      .map((y, i) => ({ y, i }))
      .filter((e): e is { y: number; i: number } => e.y != null)
      .sort((a, b) => a.y - b.y);
    const minGap = fs * 1.5;
    for (let k = 1; k < idx.length; k++) idx[k].y = Math.max(idx[k].y, idx[k - 1].y + minGap);
    if (idx.length) {
      idx[idx.length - 1].y = Math.min(idx[idx.length - 1].y, plot.y + plot.h);
      for (let k = idx.length - 2; k >= 0; k--) idx[k].y = Math.min(idx[k].y, idx[k + 1].y - minGap);
    }
    const out: (number | null)[] = ys.map(() => null);
    idx.forEach((e) => (out[e.i] = e.y));
    return out;
  };
  const leftYs = place(data.series.map((s) => (s.values[0] == null ? null : toY(s.values[0]!))));
  const rightYs = place(data.series.map((s) => (s.values[last] == null ? null : toY(s.values[last]!))));

  const columnTop: number[] = data.categories.map(() => plot.y + plot.h);
  data.series.forEach((s, si) => {
    const color = seriesColor(style, si, s.color);
    let prev: { x: number; y: number } | null = null;
    for (let c = 0; c < n; c++) {
      const v = s.values[c];
      if (v == null) {
        prev = null;
        continue;
      }
      const pt = { x: xs[c], y: toY(v) };
      columnTop[c] = Math.min(columnTop[c], pt.y);
      if (prev) {
        nodes.push({
          kind: "line",
          x1: prev.x,
          y1: prev.y,
          x2: pt.x,
          y2: pt.y,
          stroke: color,
          strokeWidth: 2,
          name: `line-${si}-${c}`,
        });
      }
      const r = 2.4;
      nodes.push({
        kind: "rect",
        x: pt.x - r,
        y: pt.y - r,
        w: r * 2,
        h: r * 2,
        fill: color,
        stroke: style.background,
        strokeWidth: 1,
        name: `marker-${si}-${c}`,
      });
      prev = pt;
    }
    if (leftYs[si] != null) {
      nodes.push({
        kind: "text",
        x: 0,
        y: leftYs[si]! - fs * 0.75,
        w: gutterL - 6,
        h: fs * 1.5,
        text: endLabel(s, 0),
        fontSize: fs,
        color,
        align: "right",
        valign: "middle",
        name: `slope-left-${si}`,
      });
    }
    if (rightYs[si] != null) {
      nodes.push({
        kind: "text",
        x: plot.x + plot.w + 6,
        y: rightYs[si]! - fs * 0.75,
        w: gutterR - 6,
        h: fs * 1.5,
        text: endLabel(s, last),
        fontSize: fs,
        color,
        align: "left",
        valign: "middle",
        name: `slope-right-${si}`,
      });
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: xs,
      categoryWidth: data.categories.map(() => fs * 2),
      columnTop,
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: plot.y + plot.h,
      plot,
      valueToY: toY,
    },
  };
}

/**
 * Bump chart: rank-over-time. Categories are periods (x); each series is an
 * entity whose values are ranks (1 = best). Ranks map onto an inverted integer
 * axis (rank 1 at the top) with thick lines, round markers and a "Name" label
 * at both ends of every line.
 */
function layoutBump(cfg: ChartConfig, style: ChartStyle, _decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const titleH = titleHeight(cfg, style);
  const headerH = fs * 1.6;
  const maxRank = Math.max(1, ...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)));
  const nameW = Math.max(fs * 3, ...data.series.map((s) => textWidth(s.name, fs))) + fs;
  const plot = {
    x: nameW,
    y: titleH + headerH,
    w: cfg.width - nameW * 2,
    h: cfg.height - titleH - headerH - fs * 1.6,
  };
  const xs = data.categories.map((_, c) => plot.x + (n === 1 ? 0 : (c / (n - 1)) * plot.w));
  const toY = (rank: number) => plot.y + (maxRank === 1 ? plot.h / 2 : ((rank - 1) / (maxRank - 1)) * plot.h);

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // Period headers along the top.
  data.categories.forEach((cat, c) => {
    nodes.push({
      kind: "text",
      x: xs[c] - 40,
      y: titleH,
      w: 80,
      h: headerH,
      text: cat,
      fontSize: fs,
      bold: true,
      color: style.text,
      align: "center",
      valign: "middle",
      name: `period-${c}`,
    });
  });

  data.series.forEach((s, si) => {
    const color = seriesColor(style, si, s.color);
    let prev: { x: number; y: number } | null = null;
    let firstC = -1;
    let lastC = -1;
    s.values.forEach((v, c) => {
      if (v == null) {
        prev = null;
        return;
      }
      if (firstC < 0) firstC = c;
      lastC = c;
      const pt = { x: xs[c], y: toY(v) };
      if (prev)
        nodes.push({
          kind: "line",
          x1: prev.x,
          y1: prev.y,
          x2: pt.x,
          y2: pt.y,
          stroke: color,
          strokeWidth: 3.5,
          name: `bump-${si}-${c}`,
        });
      prev = pt;
    });
    s.values.forEach((v, c) => {
      if (v == null) return;
      nodes.push({
        kind: "ellipse",
        cx: xs[c],
        cy: toY(v),
        rx: 4,
        ry: 4,
        fill: color,
        stroke: style.background,
        strokeWidth: 1.5,
        name: `bump-marker-${si}-${c}`,
      });
    });
    // "Name" labels at both ends of the line.
    if (firstC >= 0) {
      const y = toY(s.values[firstC]!);
      nodes.push({
        kind: "text",
        x: 0,
        y: y - fs * 0.75,
        w: nameW - fs * 0.5,
        h: fs * 1.5,
        text: s.name,
        fontSize: fs,
        bold: true,
        color,
        align: "right",
        valign: "middle",
        name: `bump-label-l-${si}`,
      });
    }
    if (lastC >= 0) {
      const y = toY(s.values[lastC]!);
      nodes.push({
        kind: "text",
        x: plot.x + plot.w + fs * 0.5,
        y: y - fs * 0.75,
        w: nameW - fs * 0.5,
        h: fs * 1.5,
        text: s.name,
        fontSize: fs,
        bold: true,
        color,
        align: "left",
        valign: "middle",
        name: `bump-label-r-${si}`,
      });
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: xs,
      categoryWidth: data.categories.map(() => fs * 2),
      columnTop: data.categories.map(() => plot.y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}

/**
 * Horizontal "profile chart": line / area rotated 90° — categories run down
 * the left axis and values extend to the right (think-cell parity). Kept
 * separate from the vertical path so that stays byte-identical.
 */
function layoutLineHorizontal(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const area = cfg.kind === "area";
  const fs = style.fontSize;

  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const stackedPos = data.categories.map((_, c) => data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0));
  const stackedNeg = data.categories.map((_, c) => data.series.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0));
  const dataMax = area ? Math.max(0, ...stackedPos) : Math.max(0, ...all);
  const dataMin = area ? Math.min(0, ...stackedNeg) : Math.min(0, ...all);
  const fmt = resolveFormat(all, cfg.numberFormat);

  const frame = computeFrameHorizontal(cfg, style, decor);
  const scale = valueScale(frame, dataMin, dataMax, cfg.scale);
  const toX = (v: number) => frame.x + ((v - scale.min) / (scale.max - scale.min || 1)) * frame.w;
  const slotH = frame.h / Math.max(1, n);
  const centers = data.categories.map((_, c) => frame.y + slotH * (c + 0.5));
  const x0 = toX(0);

  const nodes: SceneNode[] = horizontalChrome(cfg, style, decor, frame, centers, scale, (v) => toX(v) - frame.x);
  const columnTop: number[] = data.categories.map(() => x0);

  // Legend chips (multi-series) in the reserved top strip.
  if (decor.seriesLabels && data.series.length > 1) {
    let lx = frame.x;
    const ly = (cfg.title ? fs * 1.6 + 6 : 0) + fs * 0.3;
    data.series.forEach((s, si) => {
      const chip = fs * 0.7;
      nodes.push(
        {
          kind: "rect",
          x: lx,
          y: ly,
          w: chip,
          h: chip,
          fill: seriesColor(style, si, s.color),
          name: `legend-chip-${si}`,
        },
        {
          kind: "text",
          x: lx + chip + 3,
          y: ly - fs * 0.3,
          w: textWidth(s.name, fs) + 6,
          h: fs * 1.4,
          text: s.name,
          fontSize: fs,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `legend-${si}`,
        },
      );
      lx += chip + 3 + textWidth(s.name, fs) + 12;
    });
  }

  if (area) {
    const posBase = data.categories.map(() => 0);
    const negBase = data.categories.map(() => 0);
    data.series.forEach((s, si) => {
      const fill = seriesColor(style, si, s.color);
      const lower: number[] = [];
      const upper: number[] = [];
      for (let c = 0; c < n; c++) {
        const v = s.values[c] ?? 0;
        if (v >= 0) {
          lower[c] = posBase[c];
          upper[c] = posBase[c] + v;
          posBase[c] += v;
        } else {
          upper[c] = negBase[c];
          lower[c] = negBase[c] + v;
          negBase[c] += v;
        }
      }
      for (let c = 0; c < n - 1; c++) {
        const xL0 = toX(lower[c]);
        const xL1 = toX(lower[c + 1]);
        const xU0 = toX(upper[c]);
        const xU1 = toX(upper[c + 1]);
        const span = centers[c + 1] - centers[c];
        const steps = slabSteps(span);
        const h = span / steps;
        for (let k = 0; k < steps; k++) {
          const t = (k + 0.5) / steps;
          const xL = xL0 + (xL1 - xL0) * t;
          const xU = xU0 + (xU1 - xU0) * t;
          nodes.push({
            kind: "rect",
            x: Math.min(xL, xU),
            y: centers[c] + k * h,
            w: Math.abs(xU - xL),
            h: h + 0.5,
            fill,
            name: `area-${si}-${c}-${k}`,
          });
        }
      }
    });
    for (let c = 0; c < n; c++) columnTop[c] = toX(posBase[c]);
  } else {
    data.series.forEach((s, si) => {
      const color = seriesColor(style, si, s.color);
      let prev: { x: number; y: number } | null = null;
      for (let c = 0; c < n; c++) {
        const v = s.values[c];
        if (v == null) {
          if (!decor.bridgeGaps) prev = null;
          continue;
        }
        const pt = { x: toX(v), y: centers[c] };
        columnTop[c] = Math.max(columnTop[c], pt.x);
        if (prev)
          nodes.push({
            kind: "line",
            x1: prev.x,
            y1: prev.y,
            x2: pt.x,
            y2: pt.y,
            stroke: color,
            strokeWidth: 2,
            name: `line-${si}-${c}`,
          });
        const r = 2.4;
        nodes.push({
          kind: "rect",
          x: pt.x - r,
          y: pt.y - r,
          w: r * 2,
          h: r * 2,
          fill: color,
          stroke: style.background,
          strokeWidth: 1,
          name: `marker-${si}-${c}`,
        });
        if (decor.segmentLabels) {
          nodes.push({
            kind: "text",
            x: pt.x + 4,
            y: pt.y - fs * 0.75,
            w: fs * 3,
            h: fs * 1.5,
            text: formatNumber(v, fmt),
            fontSize: fs,
            color: style.text,
            align: "left",
            valign: "middle",
            name: `label-${si}-${c}`,
          });
        }
        prev = pt;
      }
    });
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => slotH * 0.6),
      columnTop,
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: x0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
    },
  };
}
