import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { lerpColor } from "../color";
import { baselineNode, chromeNodes, computeFrame, valueScale } from "./frame";
import type { LayoutResult } from "./column";

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
  const { frame } = computeFrame(cfg, style, { ...decor, seriesLabels: false }, []);
  const scale = valueScale(frame, Math.min(0, ...all), Math.max(0, ...all), cfg.scale);
  const toY = scale.toY;

  const slotW = frame.w / Math.max(1, n);
  const boxW = slotW * 0.55;
  const capW = boxW * 0.45;
  const centers = data.categories.map((_, c) => frame.x + slotW * (c + 0.5));

  const nodes: SceneNode[] = chromeNodes(cfg, style, decor, frame, centers, scale);
  const columnTop: number[] = [];

  boxes.forEach((b, c) => {
    const x = centers[c];
    if (!b) {
      columnTop.push(frame.y + frame.h);
      return;
    }
    const fill = seriesColor(style, 0, data.series.find((s) => s.color)?.color);
    const boxFill = lerpColor("#ffffff", fill, 0.22);
    const yQ1 = toY(b.q1);
    const yQ3 = toY(b.q3);
    const yMed = toY(b.median);
    columnTop.push(Math.min(toY(b.max), ...b.outliers.map(toY)));
    nodes.push(
      // Whiskers with caps.
      { kind: "line", x1: x, y1: toY(b.max), x2: x, y2: yQ3, stroke: style.axis, strokeWidth: 0.75, name: `whisker-hi-${c}` },
      { kind: "line", x1: x, y1: yQ1, x2: x, y2: toY(b.min), stroke: style.axis, strokeWidth: 0.75, name: `whisker-lo-${c}` },
      { kind: "line", x1: x - capW / 2, y1: toY(b.max), x2: x + capW / 2, y2: toY(b.max), stroke: style.axis, strokeWidth: 0.75, name: `cap-hi-${c}` },
      { kind: "line", x1: x - capW / 2, y1: toY(b.min), x2: x + capW / 2, y2: toY(b.min), stroke: style.axis, strokeWidth: 0.75, name: `cap-lo-${c}` },
      // Q1–Q3 box with a heavier median line.
      { kind: "rect", x: x - boxW / 2, y: Math.min(yQ1, yQ3), w: boxW, h: Math.abs(yQ1 - yQ3), fill: boxFill, stroke: fill, strokeWidth: 1, name: `box-${c}` },
      { kind: "line", x1: x - boxW / 2, y1: yMed, x2: x + boxW / 2, y2: yMed, stroke: fill, strokeWidth: 1.75, name: `median-${c}` },
    );
    if (b.mean != null) {
      nodes.push({
        kind: "text", x: x - fs * 0.6, y: toY(b.mean) - fs * 0.7, w: fs * 1.2, h: fs * 1.4,
        text: "×", fontSize: fs, bold: true, color: style.neutral,
        align: "center", valign: "middle", name: `mean-${c}`,
      });
    }
    b.outliers.forEach((v, i) => {
      nodes.push({ kind: "ellipse", cx: x, cy: toY(v), rx: 2.2, ry: 2.2, fill, name: `outlier-${c}-${i}` });
    });
    if (decor.segmentLabels) {
      const label = formatNumber(b.median, fmt);
      if (boxW >= textWidth(label, fs * 0.9) + 4) {
        nodes.push({
          kind: "text", x: x - boxW / 2, y: yMed - fs * 1.5, w: boxW, h: fs * 1.3,
          text: label, fontSize: fs * 0.9, color: style.text,
          align: "center", valign: "bottom", name: `median-label-${c}`,
        });
      }
    }
  });

  nodes.push(baselineNode(frame, frame.y + frame.h, style));

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => boxW),
      columnTop,
      columnValue: boxes.map((b) => b?.median ?? 0),
      baselineY: frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: toY,
    },
  };
}

/** Value extent for Same Scale, including whisker ends and outliers. */
export function boxplotExtent(cfg: ChartConfig): { min: number; max: number } | null {
  const all = cfg.data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  return all.length ? { min: Math.min(0, ...all), max: Math.max(0, ...all) } : null;
}
