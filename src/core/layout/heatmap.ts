import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, parseDateToken, resolveFormat } from "../format";
import { divergingScale, lerpColor, NO_DATA, sequentialScale } from "../color";
import { footnoteH } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Heatmap: series = rows, categories = columns, value → color on ONE global
 * scale (comparability across rows is the whole point). Sequential scale for
 * one-signed data; diverging through white, symmetric around zero, when the
 * data spans zero. Cell labels shown when they fit; compact gradient legend.
 */
export function layoutHeatmap(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const nCols = data.categories.length;
  const nRows = data.series.length;
  const fs = style.fontSize;
  const opts = cfg.heatmap ?? {};

  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const min = Math.min(...all);
  const max = Math.max(...all);
  const positive = opts.color ?? style.palette[0];
  const negative = opts.negativeColor ?? style.negative;
  const mode = opts.mode && opts.mode !== "auto" ? opts.mode : min < 0 && max > 0 ? "diverging" : "sequential";
  const constant = !all.length || min === max;
  const colorOf = constant
    ? () => lerpColor("#ffffff", positive, 0.5)
    : mode === "diverging"
      ? divergingScale(min, max, positive, negative)
      : sequentialScale(min, max, positive);
  const fmt = resolveFormat(all, cfg.numberFormat);
  const maxAbs = Math.max(1e-9, Math.abs(min), Math.abs(max));
  const sizeEncode = !!opts.sizeEncode;

  // Hierarchical row clustering: reorder rows by average-linkage similarity and
  // draw a dendrogram in a left gutter. Needs ≥3 rows and ≥2 columns.
  const clusterOn = !!opts.cluster && nRows >= 3 && nCols >= 2;
  const cl = clusterOn ? clusterRows(data.series.map((s) => data.categories.map((_, c) => s.values[c] ?? 0))) : null;
  const rows = cl ? cl.order.map((i) => data.series[i]) : data.series;

  // Calendar layout: a single daily series with date categories becomes a
  // weekday × week grid (the GitHub-contributions view).
  const calDays = data.categories.map((c) => parseDateToken(c));
  if (opts.calendar && data.series.length >= 1 && calDays.every((d): d is number => d != null)) {
    return calendarLayout(cfg, style, decor, calDays as number[], data.series[0].values, colorOf, min, max, constant);
  }

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const headerH = decor.categoryAxis !== false ? fs * 1.5 : 2;
  const rowLabelW = Math.min(cfg.width * 0.28, Math.max(fs, ...data.series.map((s) => textWidth(s.name, fs))) + 8);
  const legendH = fs * 3;
  // Marginal totals reserve a strip on the right (row sums) / bottom (column sums).
  const wantRowTotals = opts.totals === "row" || opts.totals === "both";
  const wantColTotals = opts.totals === "column" || opts.totals === "both";
  const totalsW = wantRowTotals ? fs * 4 : 0;
  const totalsH = wantColTotals ? fs * 1.9 : 0;
  const dendroW = clusterOn ? fs * 4 : 0; // left gutter for the row dendrogram
  const plot = {
    x: dendroW + rowLabelW,
    y: titleH + headerH,
    w: cfg.width - dendroW - rowLabelW - 2 - totalsW,
    h: cfg.height - titleH - headerH - legendH - totalsH - footnoteH(cfg, style, decor),
  };
  const cw = plot.w / Math.max(1, nCols);
  const ch = plot.h / Math.max(1, nRows);

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }
  if (decor.categoryAxis !== false) {
    data.categories.forEach((cat, c) => {
      nodes.push({
        kind: "text", x: plot.x + c * cw, y: titleH, w: cw, h: headerH,
        text: cat, fontSize: fs, color: style.text, align: "center", valign: "middle", name: `col-${c}`,
      });
    });
  }

  rows.forEach((s, ri) => {
    nodes.push({
      kind: "text", x: dendroW, y: plot.y + ri * ch, w: rowLabelW - 4, h: ch,
      text: s.name, fontSize: fs, color: style.text, align: "right", valign: "middle", name: `row-${ri}`,
    });
    data.categories.forEach((_, c) => {
      const v = s.values[c];
      const fill = v == null ? NO_DATA : colorOf(v);
      const x = plot.x + c * cw;
      const y = plot.y + ri * ch;
      if (sizeEncode) {
        // Centred square whose side ∝ √(|v|/maxAbs) so its area tracks magnitude.
        const frac = v == null ? 0 : Math.sqrt(Math.abs(v) / maxAbs);
        const side = Math.max(0, Math.min(cw - 1, ch - 1) * frac);
        nodes.push({ kind: "rect", x: x + (cw - 1 - side) / 2, y: y + (ch - 1 - side) / 2, w: side, h: side, fill, name: `cell-${ri}-${c}` });
      } else {
        nodes.push({ kind: "rect", x, y, w: cw - 1, h: ch - 1, fill, name: `cell-${ri}-${c}` });
      }
      if (v != null && decor.segmentLabels && !sizeEncode) {
        const label = formatNumber(v, fmt);
        if (cw >= textWidth(label, fs) + 4 && ch >= fs * 1.3) {
          nodes.push({
            kind: "text", x, y, w: cw - 1, h: ch - 1, text: label, fontSize: fs,
            color: contrastInk(fill), align: "center", valign: "middle", name: `cell-label-${ri}-${c}`,
          });
        }
      }
    });
  });

  // Marginal totals: neutral sum strips outside the color scale.
  const sum = (vals: (number | null)[]) => vals.reduce((a: number, v) => a + (v ?? 0), 0);
  if (wantRowTotals) {
    rows.forEach((s, ri) => {
      const y = plot.y + ri * ch;
      nodes.push(
        { kind: "rect", x: plot.x + plot.w + 2, y, w: totalsW - 4, h: ch - 1, fill: "#f0efec", name: `row-total-bg-${ri}` },
        {
          kind: "text", x: plot.x + plot.w + 2, y, w: totalsW - 6, h: ch - 1, text: formatNumber(sum(s.values), fmt),
          fontSize: fs * 0.95, bold: true, color: style.text, align: "center", valign: "middle", name: `row-total-${ri}`,
        },
      );
    });
  }
  if (wantColTotals) {
    data.categories.forEach((_, c) => {
      const x = plot.x + c * cw;
      const total = sum(data.series.map((s) => s.values[c]));
      nodes.push(
        { kind: "rect", x, y: plot.y + plot.h + 2, w: cw - 1, h: totalsH - 4, fill: "#f0efec", name: `col-total-bg-${c}` },
        {
          kind: "text", x, y: plot.y + plot.h + 2, w: cw - 1, h: totalsH - 4, text: formatNumber(total, fmt),
          fontSize: fs * 0.95, bold: true, color: style.text, align: "center", valign: "middle", name: `col-total-${c}`,
        },
      );
    });
  }

  // Gradient legend: a strip of small steps with min/max (and 0) labels.
  const ly = plot.y + plot.h + totalsH + fs * 0.6;
  if (constant) {
    nodes.push(
      { kind: "rect", x: plot.x, y: ly, w: fs * 1.6, h: fs * 0.9, fill: colorOf(min), name: "legend-swatch" },
      {
        kind: "text", x: plot.x + fs * 1.9, y: ly - fs * 0.25, w: fs * 8, h: fs * 1.4,
        text: all.length ? formatNumber(min, fmt) : "no data", fontSize: fs * 0.9,
        color: style.mutedText, align: "left", valign: "middle", name: "legend-min",
      },
    );
  } else {
    const lw = Math.min(plot.w * 0.5, fs * 14);
    const steps = 24;
    for (let i = 0; i < steps; i++) {
      const v = min + ((max - min) * i) / (steps - 1);
      nodes.push({ kind: "rect", x: plot.x + (lw / steps) * i, y: ly, w: lw / steps + 0.5, h: fs * 0.9, fill: colorOf(v), name: `legend-step-${i}` });
    }
    nodes.push(
      {
        kind: "text", x: plot.x, y: ly + fs * 0.95, w: lw / 2, h: fs * 1.2,
        text: formatNumber(min, fmt), fontSize: fs * 0.85, color: style.mutedText, align: "left", valign: "top", name: "legend-min",
      },
      {
        kind: "text", x: plot.x + lw / 2, y: ly + fs * 0.95, w: lw / 2, h: fs * 1.2,
        text: formatNumber(max, fmt), fontSize: fs * 0.85, color: style.mutedText, align: "right", valign: "top", name: "legend-max",
      },
    );
    if (mode === "diverging") {
      const zx = plot.x + ((0 - min) / (max - min)) * lw;
      nodes.push({
        kind: "text", x: zx - fs, y: ly + fs * 0.95, w: fs * 2, h: fs * 1.2,
        text: "0", fontSize: fs * 0.85, color: style.mutedText, align: "center", valign: "top", name: "legend-zero",
      });
    }
  }

  // Row dendrogram in the left gutter (leaves next to the labels, root at left).
  if (cl) {
    const maxH = cl.maxH || 1;
    const pos = new Map<number, number>();
    cl.order.forEach((leaf, idx) => pos.set(leaf, idx));
    const leafX = dendroW - 3;
    const nodeX = (h: number) => Math.max(2, leafX - (leafX - 2) * (h / maxH));
    const draw = (nd: ClusterNode): number => {
      if (nd.leaf != null) return plot.y + (pos.get(nd.leaf)! + 0.5) * ch;
      const yl = draw(nd.left!);
      const yr = draw(nd.right!);
      const x = nodeX(nd.height);
      const xl = nd.left!.leaf != null ? leafX : nodeX(nd.left!.height);
      const xr = nd.right!.leaf != null ? leafX : nodeX(nd.right!.height);
      nodes.push(
        { kind: "line", x1: x, y1: yl, x2: xl, y2: yl, stroke: style.mutedText, strokeWidth: 0.75, name: "dendro-h" },
        { kind: "line", x1: x, y1: yr, x2: xr, y2: yr, stroke: style.mutedText, strokeWidth: 0.75, name: "dendro-h" },
        { kind: "line", x1: x, y1: yl, x2: x, y2: yr, stroke: style.mutedText, strokeWidth: 0.75, name: "dendro-v" },
      );
      return (yl + yr) / 2;
    };
    draw(cl.root);
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.x + (c + 0.5) * cw),
      categoryWidth: data.categories.map(() => cw),
      columnTop: data.categories.map(() => plot.y),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}

/** A node in the row-clustering tree: a leaf (row index) or an internal merge. */
interface ClusterNode {
  height: number;
  members: number[];
  leaf?: number;
  left?: ClusterNode;
  right?: ClusterNode;
}

/**
 * Agglomerative average-linkage clustering of row vectors (Euclidean distance).
 * Returns the leaf order (for reordering rows), the merge tree, and the tallest
 * merge height (for scaling the dendrogram).
 */
function clusterRows(vecs: number[][]): { order: number[]; root: ClusterNode; maxH: number } {
  const dist = (a: number[], b: number[]) => Math.sqrt(a.reduce((s, v, k) => s + (v - b[k]) ** 2, 0));
  let clusters: ClusterNode[] = vecs.map((_, i) => ({ height: 0, members: [i], leaf: i }));
  const linkage = (A: ClusterNode, B: ClusterNode) => {
    let s = 0;
    for (const i of A.members) for (const j of B.members) s += dist(vecs[i], vecs[j]);
    return s / (A.members.length * B.members.length);
  };
  let maxH = 0;
  while (clusters.length > 1) {
    let bi = 0;
    let bj = 1;
    let bd = Infinity;
    for (let i = 0; i < clusters.length; i++) {
      for (let j = i + 1; j < clusters.length; j++) {
        const d = linkage(clusters[i], clusters[j]);
        if (d < bd) {
          bd = d;
          bi = i;
          bj = j;
        }
      }
    }
    const merged: ClusterNode = {
      height: bd,
      members: [...clusters[bi].members, ...clusters[bj].members],
      left: clusters[bi],
      right: clusters[bj],
    };
    maxH = Math.max(maxH, bd);
    clusters = clusters.filter((_, k) => k !== bi && k !== bj);
    clusters.push(merged);
  }
  const order: number[] = [];
  const walk = (n: ClusterNode) => {
    if (n.leaf != null) order.push(n.leaf);
    else {
      walk(n.left!);
      walk(n.right!);
    }
  };
  walk(clusters[0]);
  return { order, root: clusters[0], maxH };
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

/**
 * Calendar heatmap: a weekday (row) × week (column) grid. Days-since-epoch
 * map to a weekday (Mon=0 … Sun=6, since the epoch was a Thursday) and a week
 * index; month labels sit above the week where each month first appears.
 */
function calendarLayout(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  days: number[],
  values: (number | null)[],
  colorOf: (v: number) => string,
  min: number,
  max: number,
  constant: boolean,
): LayoutResult {
  const fs = style.fontSize;
  const weekdayOf = (d: number) => (((d + 3) % 7) + 7) % 7; // Mon=0 … Sun=6
  const minDay = Math.min(...days);
  const maxDay = Math.max(...days);
  const weekStart = minDay - weekdayOf(minDay); // Monday of the first week
  const weekOf = (d: number) => Math.floor((d - weekStart) / 7);
  const nWeeks = weekOf(maxDay) + 1;

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const wdLabelW = fs * 2.2;
  const monthH = fs * 1.4;
  const legendH = fs * 2.2;
  const gridX = wdLabelW;
  const gridY = titleH + monthH + 2;
  const availW = cfg.width - gridX - 4;
  const availH = cfg.height - gridY - legendH - footnoteH(cfg, style, decor);
  const cell = Math.max(4, Math.min(availW / Math.max(1, nWeeks), availH / 7));

  const nodes: SceneNode[] = [];
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }
  // Weekday labels (every other row to reduce clutter).
  for (let r = 0; r < 7; r++) {
    if (r % 2 === 0) {
      nodes.push({
        kind: "text", x: 0, y: gridY + r * cell, w: wdLabelW - 3, h: cell,
        text: WEEKDAYS[r], fontSize: fs * 0.8, color: style.mutedText, align: "right", valign: "middle", name: `weekday-${r}`,
      });
    }
  }
  // Month labels above the week where each month first appears.
  let lastMonth = -1;
  for (let w = 0; w < nWeeks; w++) {
    const mon = new Date((weekStart + w * 7) * 86400000).getUTCMonth();
    if (mon !== lastMonth) {
      nodes.push({
        kind: "text", x: gridX + w * cell, y: titleH, w: cell * 5, h: monthH,
        text: MONTHS[mon], fontSize: fs * 0.85, color: style.mutedText, align: "left", valign: "middle", name: `month-${w}`,
      });
      lastMonth = mon;
    }
  }
  // Day cells.
  days.forEach((d, i) => {
    const v = values[i];
    const x = gridX + weekOf(d) * cell;
    const y = gridY + weekdayOf(d) * cell;
    nodes.push({ kind: "rect", x, y, w: cell - 1.5, h: cell - 1.5, fill: v == null ? NO_DATA : colorOf(v), name: `cell-${i}` });
  });
  // Compact gradient legend (Less → More).
  const ly = gridY + 7 * cell + fs * 0.5;
  if (!constant) {
    const steps = 12;
    const sw = cell * 0.8;
    nodes.push({ kind: "text", x: gridX - fs * 3, y: ly, w: fs * 2.6, h: sw, text: "Less", fontSize: fs * 0.8, color: style.mutedText, align: "right", valign: "middle", name: "legend-less" });
    for (let i = 0; i < steps; i++) {
      nodes.push({ kind: "rect", x: gridX + i * (sw + 1), y: ly, w: sw, h: sw, fill: colorOf(min + ((max - min) * i) / (steps - 1)), name: `legend-step-${i}` });
    }
    nodes.push({ kind: "text", x: gridX + steps * (sw + 1) + 2, y: ly, w: fs * 3, h: sw, text: "More", fontSize: fs * 0.8, color: style.mutedText, align: "left", valign: "middle", name: "legend-more" });
  }

  return {
    nodes,
    anchors: {
      categoryX: days.map((d) => gridX + weekOf(d) * cell + cell / 2),
      categoryWidth: days.map(() => cell),
      columnTop: days.map(() => gridY),
      columnValue: values.map((v) => v ?? 0),
      baselineY: gridY + 7 * cell,
      plot: { x: gridX, y: gridY, w: nWeeks * cell, h: 7 * cell },
    },
  };
}
