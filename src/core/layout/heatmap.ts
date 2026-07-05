import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
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

  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  const headerH = decor.categoryAxis !== false ? fs * 1.5 : 2;
  const rowLabelW = Math.min(cfg.width * 0.28, Math.max(fs, ...data.series.map((s) => textWidth(s.name, fs))) + 8);
  const legendH = fs * 3;
  // Marginal totals reserve a strip on the right (row sums) / bottom (column sums).
  const wantRowTotals = opts.totals === "row" || opts.totals === "both";
  const wantColTotals = opts.totals === "column" || opts.totals === "both";
  const totalsW = wantRowTotals ? fs * 4 : 0;
  const totalsH = wantColTotals ? fs * 1.9 : 0;
  const plot = {
    x: rowLabelW,
    y: titleH + headerH,
    w: cfg.width - rowLabelW - 2 - totalsW,
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

  data.series.forEach((s, ri) => {
    nodes.push({
      kind: "text", x: 0, y: plot.y + ri * ch, w: rowLabelW - 4, h: ch,
      text: s.name, fontSize: fs, color: style.text, align: "right", valign: "middle", name: `row-${ri}`,
    });
    data.categories.forEach((_, c) => {
      const v = s.values[c];
      const fill = v == null ? NO_DATA : colorOf(v);
      const x = plot.x + c * cw;
      const y = plot.y + ri * ch;
      nodes.push({ kind: "rect", x, y, w: cw - 1, h: ch - 1, fill, name: `cell-${ri}-${c}` });
      if (v != null && decor.segmentLabels) {
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
    data.series.forEach((s, ri) => {
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
