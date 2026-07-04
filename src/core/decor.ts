import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors } from "./types";
import type { SceneNode } from "./scene";
import { cagr, formatNumber, formatPercent } from "./format";

/**
 * think-cell's signature annotations, computed from layout anchors so they
 * work identically across chart types.
 */
export function decorationNodes(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  a: LayoutAnchors,
): SceneNode[] {
  const nodes: SceneNode[] = [];
  const fs = style.fontSize;

  // --- CAGR arrow: diagonal arrow between two column tops with "+x.x% p.a." ---
  if (decor.cagr) {
    const { from, to } = clampPair(decor.cagr, a);
    const rate = cagr(a.columnValue[from], a.columnValue[to], to - from);
    // Clear the column totals row when it is shown.
    const lift = fs * 1.6 + (decor.totals ? fs * 1.5 : 0);
    const x1 = a.categoryX[from];
    const y1 = a.columnTop[from] - lift;
    const x2 = a.categoryX[to];
    const y2 = a.columnTop[to] - lift;
    const angle = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
    nodes.push(
      { kind: "line", x1, y1, x2, y2, stroke: style.text, strokeWidth: 1.25, name: "cagr-line" },
      { kind: "arrowhead", x: x2, y: y2, angle, size: 5, fill: style.text, name: "cagr-head" },
      {
        kind: "text",
        x: (x1 + x2) / 2 - 45,
        y: Math.min(y1, y2) - fs * 1.6,
        w: 90,
        h: fs * 1.4,
        text: rate == null ? "CAGR n/a" : `${formatPercent(rate, 1, true)} p.a.`,
        fontSize: fs,
        bold: true,
        color: style.text,
        align: "center",
        valign: "bottom",
        name: "cagr-label",
      },
    );
  }

  // --- Difference arrow: dashed level line + vertical arrow ---
  // Total (column totals) by default; a level difference arrow when `series`
  // names a series — it compares the cumulative stack level at that series.
  if (decor.difference) {
    const { from, to } = clampPair(decor.difference, a);
    const si = decor.difference.series;
    const useLevel = si != null && a.seriesLevels != null && a.valueToY != null && si >= 0 && si < (a.seriesLevels[0]?.length ?? 0);
    const vFrom = useLevel ? a.seriesLevels![from][si!] : a.columnValue[from];
    const vTo = useLevel ? a.seriesLevels![to][si!] : a.columnValue[to];
    const yFrom = useLevel ? a.valueToY!(vFrom) : a.columnTop[from];
    const yTo = useLevel ? a.valueToY!(vTo) : a.columnTop[to];
    const x = a.categoryX[to] + a.categoryWidth[to] / 2 + 10;
    nodes.push(
      {
        kind: "line",
        x1: a.categoryX[from] + a.categoryWidth[from] / 2 + 2,
        y1: yFrom,
        x2: x + 4,
        y2: yFrom,
        stroke: style.mutedText,
        strokeWidth: 0.75,
        dash: [2, 2],
        name: "diff-level",
      },
      { kind: "line", x1: x, y1: yFrom, x2: x, y2: yTo, stroke: style.text, strokeWidth: 1.25, name: "diff-line" },
      { kind: "arrowhead", x, y: yTo, angle: yTo < yFrom ? -90 : 90, size: 5, fill: style.text, name: "diff-head" },
    );
    const usePct = decor.difference.percent ?? true;
    const label =
      usePct && vFrom !== 0
        ? formatPercent(vTo / vFrom - 1, 0, true)
        : formatNumber(vTo - vFrom, { ...cfg.numberFormat, forceSign: true });
    nodes.push({
      kind: "text",
      x: x + 3,
      y: (yFrom + yTo) / 2 - fs * 0.75,
      w: Math.max(30, cfg.width - x - 3),
      h: fs * 1.5,
      text: label,
      fontSize: fs,
      bold: true,
      color: style.text,
      align: "left",
      valign: "middle",
      name: "diff-label",
    });
  }

  // --- Value lines: dashed horizontals at fixed values or the mean of totals ---
  const valueLines = decor.valueLines ?? (decor.valueLine ? [decor.valueLine] : []);
  if (valueLines.length && a.valueToY) {
    valueLines.forEach((vl, i) => {
      const value =
        vl.mode === "mean"
          ? a.columnValue.reduce((s, v) => s + v, 0) / (a.columnValue.length || 1)
          : vl.value;
      const y = a.valueToY!(value);
      nodes.push(
        {
          kind: "line",
          x1: a.plot.x,
          y1: y,
          x2: a.plot.x + a.plot.w,
          y2: y,
          stroke: style.mutedText,
          strokeWidth: 1,
          dash: [3, 2],
          name: `value-line-${i}`,
        },
        {
          kind: "text",
          x: a.plot.x + 2,
          y: y - fs * 1.5,
          w: 100,
          h: fs * 1.4,
          text: (vl.mode === "mean" ? "Ø " : "") + formatNumber(value, cfg.numberFormat),
          fontSize: fs * 0.95,
          color: style.mutedText,
          align: "left",
          valign: "bottom",
          name: `value-line-label-${i}`,
        },
      );
    });
  }

  return nodes;
}

function clampPair(p: { from: number; to: number }, a: LayoutAnchors): { from: number; to: number } {
  const n = a.categoryX.length;
  const from = Math.max(0, Math.min(n - 1, p.from));
  const to = Math.max(0, Math.min(n - 1, p.to));
  return from <= to ? { from, to } : { from: to, to: from };
}
