import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { fitPlot, footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Funnel chart: pipeline stages as horizontally centered bands whose width
 * is proportional to the value (the Power BI convention — plain rectangles,
 * so it renders natively in every renderer). Between stages, a muted
 * conversion label states the % of the previous stage. Order the values
 * ascending for a pyramid.
 */
export function layoutFunnel(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const n = data.categories.length;
  const values = data.categories.map((_, c) => Math.max(0, data.series[0]?.values[c] ?? 0));
  const vMax = Math.max(1, ...values);
  const fmt = resolveFormat(values, cfg.numberFormat);

  const titleH = titleHeight(cfg, style);
  // Left gutter for stage names when category labels are on (default).
  const catW = decor.categoryAxis
    ? Math.min(cfg.width * 0.28, Math.max(0, ...data.categories.map((c) => textWidth(c, fs))) + 10)
    : 2;
  const plot = fitPlot(cfg, {
    x: catW,
    y: titleH + 2,
    w: cfg.width - catW - 4,
    h: cfg.height - titleH - 2 - footnoteH(cfg, style, decor) - 4,
  });
  // Room for the conversion label between bands — but never more than the plot
  // can pay for while still giving every band ≥1pt. A fixed gap on a short frame
  // with many stages drove the cumulative pitch (bandH + gap per stage) past the
  // bottom of the plot, so the last bands rendered OFF-frame; flooring bandH alone
  // did not help because the gap was the overspend. Reserve 1pt per band, then
  // split the remainder as gaps.
  const gap = Math.max(0, Math.min(fs * 1.5, (plot.h - n) / Math.max(1, n - 1)));
  const bandH = Math.max(1, (plot.h - gap * (n - 1)) / Math.max(1, n));
  const cx = plot.x + plot.w / 2;

  // The BANDS are solved for the plot above, so they always fit. Their LABELS
  // were not, and each escaped the frame a different way once the font outgrew
  // the chart:
  //
  //   - a label is centred in its band and its ink is about `f` tall, so once
  //     the font passed the band height the bottom row's label hung below the
  //     plot — 4.4pt at a 28pt font, 15.1 at 36;
  //   - a category name sits right-aligned in a gutter capped at 28% of the
  //     width, and a name wider than that cap runs off the LEFT edge of the
  //     chart: "Negotiation" by 35.9pt at 28, and 83.4pt at 36.
  //
  // Both are the same answer the agenda and the process flow already use, and
  // the one this file's value labels now use for the horizontal case: shrink
  // until it fits, then clip the remainder. Shrunk TOGETHER so a row's name and
  // its number stay the same size, and last-resort — at any font that already
  // fits, `labelFs` is `fs` and nothing moves.
  const labelFs = (() => {
    const gutter = catW - 4;
    const tooTall = (f: number) => f * 1.15 > bandH;
    const tooWide = (f: number) => decor.categoryAxis && data.categories.some((c) => textWidth(c, f) > gutter);
    let f = fs;
    while (f > 6 && (tooTall(f) || tooWide(f))) f -= 0.5;
    return f;
  })();

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  const columnTop: number[] = [];
  values.forEach((v, c) => {
    const y = plot.y + c * (bandH + gap);
    columnTop.push(y);
    const w = Math.max(2, (v / vMax) * plot.w);
    const fill = data.series[0]?.colors?.[c] ?? style.palette[c % style.palette.length];
    nodes.push({ kind: "rect", x: cx - w / 2, y, w, h: bandH, fill, name: `stage-${c}` });

    if (decor.categoryAxis) {
      nodes.push({
        kind: "text",
        x: 0,
        y: y + bandH / 2 - labelFs * 0.75,
        w: catW - 4,
        h: labelFs * 1.5,
        text: clipToWidth(data.categories[c], labelFs, catW - 4),
        fontSize: labelFs,
        color: style.text,
        align: "right",
        valign: "middle",
        name: `category-${c}`,
      });
    }
    if (decor.segmentLabels) {
      const label = formatNumber(v, fmt);
      const labelW = textWidth(label, labelFs) + 6;
      // Outside means to the RIGHT of the band, and the widest band already
      // reaches the edge of the plot — so on a large font there is no room out
      // there and the label was drawn past the frame. `stage-value-0` landed at
      // x = 480.0 on a 480pt frame: the top stage of the funnel showed no value
      // at all while every stage below it did, because the one band big enough
      // to matter is the one with nothing to its right.
      //
      // So "outside" has to be a placement the frame can hold. When it is not,
      // the label goes inside after all: the band that cannot fit a label beside
      // it is the WIDEST one, which is exactly the band with the most room in
      // it. A cramped label on the bar beats a missing number.
      const roomOutside = cx + w / 2 + 4 + labelW <= cfg.width;
      const fitsInside = w >= textWidth(label, labelFs) + 8 && bandH >= labelFs * 1.3;
      const inside = fitsInside || !roomOutside;
      nodes.push({
        kind: "text",
        x: inside ? cx - w / 2 : cx + w / 2 + 4,
        y: y + bandH / 2 - labelFs * 0.75,
        w: inside ? w : labelW,
        h: labelFs * 1.5,
        text: label,
        fontSize: labelFs,
        bold: true,
        color: inside ? contrastInk(fill) : style.text,
        align: inside ? "center" : "left",
        valign: "middle",
        name: `stage-value-${c}`,
      });
    }
    // Conversion rate vs the previous stage, in the gap between bands. The
    // marker follows the direction: a fixed ▾ contradicted itself on the
    // ascending (pyramid) ordering this file recommends, printing "▾ 500.0%".
    if (c > 0 && values[c - 1] > 0) {
      const marker = v > values[c - 1] ? "▴ " : v < values[c - 1] ? "▾ " : "";
      nodes.push({
        kind: "text",
        x: cx - 40,
        y: y - gap,
        w: 80,
        h: gap,
        text: `${marker}${formatPercent(v / values[c - 1], 1, false, cfg.numberFormat?.locale)}`,
        // Off the row font, not the chart font: the conversion rate is muted
        // secondary text and must stay SMALLER than the stage names beside it.
        // Shrinking only the names inverted that — at a 28pt font the names
        // came down to 14 and these stayed at 23.8, so the percentages became
        // the loudest thing on the chart. At any font that fits, labelFs is fs
        // and this is the value it always had.
        fontSize: labelFs * 0.85,
        color: style.mutedText,
        align: "center",
        valign: "middle",
        name: `conversion-${c}`,
      });
    }
  });

  return {
    nodes,
    anchors: {
      categoryX: values.map(() => cx),
      categoryWidth: values.map((v) => Math.max(2, (v / vMax) * plot.w)),
      columnTop,
      columnValue: values,
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
