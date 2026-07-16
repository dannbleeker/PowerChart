import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { chromeNodes, computeFrame, computeFrameHorizontal } from "./frame";
import { legendRow, seriesLabelNodes, type LayoutResult } from "./column";

/**
 * Mekko (Marimekko) chart, think-cell style. Two variants:
 * - %-axis (default): column widths proportional to column totals, columns
 *   normalized to full height, so segment area ∝ absolute value.
 * - "Mekko with units": explicit column widths from the datasheet's
 *   `X extent` row; column heights represent absolute totals on a value scale.
 * cfg.horizontal rotates the chart: categories become rows.
 */
export function layoutMekko(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  const H = !!cfg.horizontal;
  const units = !!data.xExtent?.some((v) => v != null && v > 0);

  const totals = data.categories.map((_, c) =>
    data.series.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0),
  );
  const extents = units
    ? data.categories.map((_, c) => Math.max(0, data.xExtent?.[c] ?? 0))
    : totals;
  const maxTotal = Math.max(1e-9, ...totals);
  const fmt = resolveFormat(
    [...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)), ...totals],
    cfg.numberFormat,
  );

  const decorFull = { ...decor, totals: true };
  let frame = H
    ? computeFrameHorizontal(cfg, style, { ...decorFull, totals: true })
    : computeFrame(cfg, style, decorFull, decor.seriesLabels ? data.series.map((s) => s.name) : []).frame;
  if (H && !units && decor.categoryAxis) {
    // Row labels carry a share suffix ("EMEA (32%)") the generic frame
    // reservation doesn't know about — widen the left gutter for it.
    const extra = textWidth(" (00%)", fs);
    frame = { ...frame, x: frame.x + extra, w: frame.w - extra };
  }
  const grand = extents.reduce((a, b) => a + b, 0) || 1;
  const gap = 2;

  const nodes: SceneNode[] = [];
  const centers: number[] = [];
  const widths: number[] = [];
  const columnTop: number[] = [];
  const lastSegMid: (number | null)[] = data.series.map(() => null);

  // Category extent runs along x (vertical) or y (horizontal).
  const catLen = (H ? frame.h : frame.w) - gap * (n - 1);
  let pos = H ? frame.y : frame.x;

  for (let c = 0; c < n; c++) {
    const ext = (extents[c] / grand) * catLen;
    centers.push(pos + ext / 2);
    widths.push(ext);
    // %-variant: every column fills the plot; units variant: length ∝ total.
    const colLen = units ? (totals[c] / maxTotal) * (H ? frame.w : frame.h) : H ? frame.w : frame.h;

    let acc = H ? frame.x : frame.y + frame.h;
    data.series.forEach((s, si) => {
      const v = Math.max(0, s.values[c] ?? 0);
      if (v === 0 || totals[c] === 0) return;
      const segLen = (v / totals[c]) * colLen;
      const r = H
        ? { x: acc, y: pos, w: segLen, h: ext }
        : { x: pos, y: acc - segLen, w: ext, h: segLen };
      acc = H ? acc + segLen : acc - segLen;
      const fill = seriesColor(style, si, s.color);
      nodes.push({ kind: "rect", ...r, fill, stroke: style.background, strokeWidth: 0.75, name: `seg-${si}-${c}` });
      if (c === n - 1) lastSegMid[si] = H ? r.x + r.w / 2 : r.y + r.h / 2;
      // The label is centred in r.h and spans r.w in BOTH orientations, so the
      // vertical fit is r.h either way. (Gating it on r.w for horizontal mekko
      // measured the value-axis length, not the row thickness, so labels
      // rendered in rows thinner than the font.)
      //
      // Horizontal segments also need a length floor. The fit check below allows
      // 2pt of bleed — the text box is deliberately 4pt wider than the segment —
      // which is harmless for a vertical mekko's wide columns but lets a row of
      // hairline-thin horizontal segments print their labels on top of each
      // other. Vertical keeps exactly the room it always had.
      const roomAcross = r.h >= fs * 1.25;
      const roomAlong = !H || r.w >= fs * 1.25;
      if (decor.segmentLabels && roomAcross && roomAlong) {
        const label = formatNumber(v, fmt);
        if (textWidth(label, fs) <= r.w + 2) {
          nodes.push({
            kind: "text",
            x: r.x - 2,
            y: r.y + r.h / 2 - fs * 0.75,
            w: r.w + 4,
            h: fs * 1.5,
            text: label,
            fontSize: fs,
            color: contrastInk(fill),
            align: "center",
            valign: "middle",
            name: `label-${si}-${c}`,
          });
        }
      }
    });
    columnTop.push(H ? frame.x + colLen : frame.y + frame.h - colLen);

    // Column total at the column's end — the Mekko signature.
    if (H) {
      nodes.push({
        kind: "text",
        x: frame.x + colLen + 3,
        y: centers[c] - fs * 0.75,
        w: cfg.width - (frame.x + colLen) - 3,
        h: fs * 1.5,
        text: formatNumber(totals[c], fmt),
        fontSize: fs,
        bold: true,
        color: style.text,
        align: "left",
        valign: "middle",
        name: `total-${c}`,
      });
    } else {
      nodes.push({
        kind: "text",
        x: pos - 4,
        y: frame.y + frame.h - colLen - fs * 1.45,
        w: ext + 8,
        h: fs * 1.4,
        text: formatNumber(totals[c], fmt),
        fontSize: fs,
        bold: true,
        color: style.text,
        align: "center",
        valign: "bottom",
        name: `total-${c}`,
      });
    }

    pos += ext + gap;
  }

  // Chrome: title only via chromeNodes, custom category labels per orientation.
  nodes.push(
    ...chromeNodes(cfg, style, { ...decorFull, categoryAxis: false, valueAxis: false, gridlines: false }, frame, centers),
  );
  if (decor.categoryAxis) {
    for (let c = 0; c < n; c++) {
      const label = units
        ? data.categories[c]
        : `${data.categories[c]} (${formatPercent(extents[c] / grand)})`;
      if (H) {
        nodes.push({
          kind: "text",
          x: 0,
          y: centers[c] - fs * 0.75,
          w: frame.x - 4,
          h: fs * 1.5,
          text: label,
          fontSize: fs,
          color: style.text,
          align: "right",
          valign: "middle",
          name: `category-${c}`,
        });
      } else {
        nodes.push({
          kind: "text",
          x: centers[c] - widths[c] / 2 - 4,
          y: frame.y + frame.h + 3,
          w: widths[c] + 8,
          h: fs * 1.4,
          text: label,
          fontSize: fs,
          color: style.text,
          align: "center",
          valign: "top",
          name: `category-${c}`,
        });
      }
    }
  }
  if (H) {
    nodes.push({ kind: "line", x1: frame.x, y1: frame.y, x2: frame.x, y2: frame.y + frame.h, stroke: style.axis, strokeWidth: 1, name: "baseline" });
    if (decor.seriesLabels && data.series.length > 1) {
      nodes.push(...legendRow(cfg, style, frame.x, (cfg.title ? fs * 1.6 + 6 : 0) + 2));
    }
  } else {
    nodes.push({
      kind: "line",
      x1: frame.x,
      y1: frame.y + frame.h,
      x2: frame.x + frame.w,
      y2: frame.y + frame.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
    if (decor.seriesLabels) nodes.push(...seriesLabelNodes(cfg, style, frame, lastSegMid));
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: widths,
      columnTop,
      columnValue: totals,
      baselineY: H ? frame.x : frame.y + frame.h,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
    },
  };
}
