import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { lerpColor } from "../color";
import { baselineNode, chromeNodes, computeFrame, computeFrameHorizontal, valueScale } from "./frame";
import { horizontalChrome, type LayoutResult } from "./column";

/** The five-number-summary datasheet rows (think-cell's own boxplot recipe). */
const SUMMARY_ROWS: [keyof FiveNum, RegExp][] = [
  ["min", /^min$/i],
  ["q1", /^q1$/i],
  ["median", /^median$/i],
  ["q3", /^q3$/i],
  ["max", /^max$/i],
];
const MEAN_ROW = /^mean$/i;
const OUTLIER_ROW = /^outliers?\b/i;

interface FiveNum {
  min: number;
  q1: number;
  median: number;
  q3: number;
  max: number;
}

interface Box extends FiveNum {
  mean?: number;
  outliers: number[];
}

/** Excel-style exclusive/inclusive quartile on a sorted sample. */
function quartile(sorted: number[], p: number, method: "exclusive" | "inclusive"): number {
  const n = sorted.length;
  const rank = method === "exclusive" ? p * (n + 1) - 1 : p * (n - 1);
  const lo = Math.max(0, Math.min(n - 1, Math.floor(rank)));
  const hi = Math.max(0, Math.min(n - 1, Math.ceil(rank)));
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (rank - lo);
}

/**
 * Box-and-whisker chart, one box per category. Two datasheet modes:
 *  - Precomputed: rows named Min/Q1/Median/Q3/Max (+ optional Mean,
 *    Outlier …) carry the summary directly — whiskers run to Min/Max.
 *  - Raw samples: every row is an observation; quartiles are computed
 *    (exclusive method) and whiskers use Tukey 1.5×IQR fences with
 *    outliers drawn as dots.
 */
export function layoutBoxplot(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  const opts = cfg.boxplot ?? {};
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()));
  const precomputed = SUMMARY_ROWS.some(([, re]) => find(re));

  const boxes: (Box | null)[] = data.categories.map((_, c) => {
    if (precomputed) {
      const summary = {} as FiveNum;
      for (const [key, re] of SUMMARY_ROWS) {
        const v = find(re)?.values[c];
        if (v == null) return null;
        summary[key] = v;
      }
      const mean = find(MEAN_ROW)?.values[c] ?? undefined;
      const outliers = data.series
        .filter((s) => OUTLIER_ROW.test(s.name.trim()))
        .map((s) => s.values[c])
        .filter((v): v is number => v != null);
      return { ...summary, mean: mean ?? undefined, outliers };
    }
    // Raw samples: this category's column across all rows.
    const sample = data.series.map((s) => s.values[c]).filter((v): v is number => v != null).sort((a, b) => a - b);
    if (sample.length < 2) return null;
    const method = opts.quartileMethod ?? "exclusive";
    const q1 = quartile(sample, 0.25, method);
    const q3 = quartile(sample, 0.75, method);
    const median = quartile(sample, 0.5, "inclusive");
    const whiskers = opts.whiskers ?? "tukey";
    let lo = sample[0];
    let hi = sample[sample.length - 1];
    let outliers: number[] = [];
    if (whiskers === "tukey") {
      const k = opts.iqrMultiplier ?? 1.5;
      const loFence = q1 - k * (q3 - q1);
      const hiFence = q3 + k * (q3 - q1);
      const inside = sample.filter((v) => v >= loFence && v <= hiFence);
      lo = inside[0] ?? q1;
      hi = inside[inside.length - 1] ?? q3;
      outliers = sample.filter((v) => v < loFence || v > hiFence);
    }
    const mean = opts.showMean ? sample.reduce((a, b) => a + b, 0) / sample.length : undefined;
    return { min: lo, q1, median, q3, max: hi, mean, outliers };
  });

  const all = boxes.flatMap((b) => (b ? [b.min, b.max, ...b.outliers, ...(b.mean != null ? [b.mean] : [])] : []));
  const fmt = resolveFormat(all, cfg.numberFormat);
  const H = !!cfg.horizontal;
  const frame = H
    ? computeFrameHorizontal(cfg, style, { ...decor, seriesLabels: false, totals: false })
    : computeFrame(cfg, style, { ...decor, seriesLabels: false }, []).frame;
  // One shared value scale across every box — that is the point of putting
  // them in one chart.
  const scale = valueScale(frame, Math.min(0, ...all), Math.max(0, ...all), cfg.scale);
  // Value coordinate along the value axis (x when horizontal, y otherwise).
  const qOf = H
    ? (v: number) => frame.x + ((v - scale.min) / (scale.max - scale.min || 1)) * frame.w
    : scale.toY;

  const catStart = H ? frame.y : frame.x;
  const catLen = H ? frame.h : frame.w;
  const slotW = catLen / Math.max(1, n);
  const boxW = slotW * 0.55;
  const capW = boxW * 0.45;
  const centers = data.categories.map((_, c) => catStart + slotW * (c + 0.5));

  // Summary rows (Min/Q1/…) are anatomy, not series — never a legend.
  const nodes: SceneNode[] = H
    ? horizontalChrome(cfg, style, { ...decor, totals: false, seriesLabels: false }, frame, centers, scale, (v) => qOf(v) - frame.x)
    : chromeNodes(cfg, style, { ...decor, seriesLabels: false }, frame, centers, scale);
  const columnTop: number[] = [];

  /** Category-axis point p + value-axis point q → chart coordinates. */
  const pt = (p: number, q: number) => (H ? { x: q, y: p } : { x: p, y: q });
  const segLine = (p1: number, q1: number, p2: number, q2: number, weight: number, stroke: string, nm: string): SceneNode => {
    const a = pt(p1, q1);
    const b = pt(p2, q2);
    return { kind: "line", x1: a.x, y1: a.y, x2: b.x, y2: b.y, stroke, strokeWidth: weight, name: nm };
  };

  boxes.forEach((b, c) => {
    const p = centers[c];
    if (!b) {
      columnTop.push(frame.y + frame.h);
      return;
    }
    const fill = seriesColor(style, 0, data.series.find((s) => s.color)?.color);
    const boxFill = lerpColor("#ffffff", fill, 0.22);
    const qQ1 = qOf(b.q1);
    const qQ3 = qOf(b.q3);
    const qMed = qOf(b.median);
    columnTop.push(H ? p - boxW / 2 : Math.min(qOf(b.max), ...b.outliers.map(qOf)));
    const boxLo = Math.min(qQ1, qQ3);
    nodes.push(
      // Whiskers with caps.
      segLine(p, qOf(b.max), p, qQ3, 0.75, style.axis, `whisker-hi-${c}`),
      segLine(p, qQ1, p, qOf(b.min), 0.75, style.axis, `whisker-lo-${c}`),
      segLine(p - capW / 2, qOf(b.max), p + capW / 2, qOf(b.max), 0.75, style.axis, `cap-hi-${c}`),
      segLine(p - capW / 2, qOf(b.min), p + capW / 2, qOf(b.min), 0.75, style.axis, `cap-lo-${c}`),
      // Q1–Q3 box with a heavier median line.
      H
        ? { kind: "rect", x: boxLo, y: p - boxW / 2, w: Math.abs(qQ1 - qQ3), h: boxW, fill: boxFill, stroke: fill, strokeWidth: 1, name: `box-${c}` }
        : { kind: "rect", x: p - boxW / 2, y: boxLo, w: boxW, h: Math.abs(qQ1 - qQ3), fill: boxFill, stroke: fill, strokeWidth: 1, name: `box-${c}` },
      segLine(p - boxW / 2, qMed, p + boxW / 2, qMed, 1.75, fill, `median-${c}`),
    );
    if (b.mean != null) {
      const m = pt(p, qOf(b.mean));
      nodes.push({
        kind: "text", x: m.x - fs * 0.6, y: m.y - fs * 0.7, w: fs * 1.2, h: fs * 1.4,
        text: "×", fontSize: fs, bold: true, color: style.neutral,
        align: "center", valign: "middle", name: `mean-${c}`,
      });
    }
    b.outliers.forEach((v, i) => {
      const o = pt(p, qOf(v));
      nodes.push({ kind: "ellipse", cx: o.x, cy: o.y, rx: 2.2, ry: 2.2, fill, name: `outlier-${c}-${i}` });
    });
    if (decor.segmentLabels) {
      const label = formatNumber(b.median, fmt);
      if (boxW >= fs * 1.2 && (H || boxW >= textWidth(label, fs * 0.9) + 4)) {
        nodes.push(
          H
            ? {
                kind: "text", x: qMed - 30, y: p - boxW / 2 - fs * 1.35, w: 60, h: fs * 1.3,
                text: label, fontSize: fs * 0.9, color: style.text,
                align: "center", valign: "bottom", name: `median-label-${c}`,
              }
            : {
                kind: "text", x: p - boxW / 2, y: qMed - fs * 1.5, w: boxW, h: fs * 1.3,
                text: label, fontSize: fs * 0.9, color: style.text,
                align: "center", valign: "bottom", name: `median-label-${c}`,
              },
        );
      }
    }
  });

  if (H) {
    nodes.push({ kind: "line", x1: frame.x, y1: frame.y, x2: frame.x, y2: frame.y + frame.h, stroke: style.axis, strokeWidth: 1, name: "baseline" });
  } else {
    nodes.push(baselineNode(frame, frame.y + frame.h, style));
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => boxW),
      columnTop,
      columnValue: boxes.map((b) => b?.median ?? 0),
      baselineY: H ? frame.x : frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: H ? undefined : qOf,
    },
  };
}

/** Value extent for Same Scale, including whisker ends and outliers. */
export function boxplotExtent(cfg: ChartConfig): { min: number; max: number } | null {
  const all = cfg.data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  return all.length ? { min: Math.min(0, ...all), max: Math.max(0, ...all) } : null;
}
