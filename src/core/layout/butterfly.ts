import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { niceTicks } from "../format";
import { legendRow, type LayoutResult, type LegendEntry } from "./column";
import { fitPlot, legendRowCount, titleHeight, titleNode } from "./frame";

/**
 * Butterfly (tornado) chart: think-cell models this as two bar charts placed
 * back-to-back, one rotated 180°, sharing the same scale. Uses the first two
 * series; category labels sit in the center gutter.
 */
export function layoutButterfly(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const n = data.categories.length;
  const fs = style.fontSize;
  // Flanks: with butterfly.split, the first `split` series stack on the left
  // and the rest on the right; otherwise the classic series 0 / series 1.
  const withIdx = data.series.map((s, si) => ({ s, si }));
  const split = cfg.butterfly?.split;
  const leftSeries = split != null ? withIdx.slice(0, split) : withIdx.slice(0, 1);
  const rightSeries = split != null ? withIdx.slice(split) : withIdx.slice(1, 2);
  const stacked = leftSeries.length > 1 || rightSeries.length > 1;
  const stackSum = (series: typeof withIdx, c: number) =>
    series.reduce((a, { s }) => a + Math.abs(s.values[c] ?? 0), 0);
  const signedSum = (series: typeof withIdx, c: number) => series.reduce((a, { s }) => a + (s.values[c] ?? 0), 0);

  const titleH = titleHeight(cfg, style);
  // Room for outside value labels on each flank. Scaled down below, with the
  // category gutter, when the two together would leave the BARS nothing.
  const valueW0 = fs * 3.4;
  const gutterW0 = Math.min(cfg.width * 0.3, Math.max(0, ...data.categories.map((c) => textWidth(c, fs))) + 12);
  // The chrome — a value strip on each flank and the category gutter down the
  // middle — is priced in font sizes and label widths, so on a small frame it
  // takes the chart: 84 of 120 points at a thumbnail, leaving both sets of bars
  // 36 between them and the longest one 18pt. The bars are what a butterfly is
  // read by, so the chrome is scaled to fit inside HALF the width, the same way
  // the gantt scales its three text gutters into a budget. At any size where it
  // already fits, the scale is 1 and nothing moves.
  const chromeBudget = cfg.width * 0.5;
  const chrome = valueW0 * 2 + gutterW0;
  const chromeScale = chrome > chromeBudget ? chromeBudget / chrome : 1;
  const valueW = valueW0 * chromeScale;
  const gutterW = gutterW0 * chromeScale;
  // Stacked flanks legend the whole series set across the top; it wraps (shared
  // walk), so the header band has to be as tall as the rows it will occupy.
  const legendEntries: LegendEntry[] = stacked
    ? [...leftSeries, ...rightSeries].map(({ s, si }) => ({
        label: s.name,
        color: seriesColor(style, si, s.color),
        name: `legend-${si}`,
      }))
    : [];
  const headerH =
    fs *
    1.6 *
    Math.max(
      1,
      legendRowCount(
        legendEntries.map((e) => e.label),
        fs,
        valueW,
        cfg.width - 4,
      ),
    );
  // A value axis reserves a strip at the bottom for tick labels on both flanks.
  const axisH = decor.valueAxis ? fs * 1.5 : 0;
  const plot = fitPlot(cfg, {
    x: valueW,
    y: titleH + headerH + 2,
    w: Math.max(0, cfg.width - valueW * 2),
    // Floor at 0 for the same reason halfW is floored below: a chart too short
    // for its title/header/axis chrome would otherwise give every bar rect a
    // NEGATIVE height, which SVG drops and PowerPoint clamps to a sliver.
    h: Math.max(0, cfg.height - titleH - headerH - 6 - axisH),
  });
  // Floor at 0: a very narrow frame can drive plot.w below the gutter, which
  // would give the header texts and bar rects negative widths.
  const halfW = Math.max(0, (plot.w - gutterW) / 2);
  const leftEdge = plot.x + halfW; // right edge of the left half
  const rightEdge = leftEdge + gutterW; // left edge of the right half

  const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null)).map((v) => Math.abs(v));
  const sums = data.categories.flatMap((_, c) => [stackSum(leftSeries, c), stackSum(rightSeries, c)]);
  const ticks = niceTicks(0, Math.max(1, ...sums), 4);
  const max = ticks[ticks.length - 1];
  const fmt = resolveFormat(all, cfg.numberFormat);
  const qOf = (v: number) => (Math.abs(v) / max) * halfW;

  const slotH = plot.h / Math.max(1, n);
  // A value label is centred on its row in a box `fs * 1.5` tall, so once the
  // font outgrows the row pitch the labels overlap each OTHER and the last one
  // leaves the plot — 10.7pt past a 200x150 frame at a 32pt font, and colliding
  // at any frame size once the font is big enough for the row count. Bound by
  // the row it labels, which is the same thing that stops both. Last resort: at
  // any font that already fits its row this is `fs` and nothing moves.
  const rowFs = Math.min(fs, slotH / 1.5);
  const barH = slotH * (2 / 3);

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  if (!stacked) {
    // Series headers above each half (classic two-series butterfly).
    (
      [
        [leftSeries[0], plot.x, leftEdge],
        [rightSeries[0], rightEdge, plot.x + plot.w],
      ] as const
    ).forEach(([entry, x0, x1], i) => {
      nodes.push({
        kind: "text",
        x: x0,
        y: titleH,
        w: Math.max(0, x1 - x0),
        h: headerH,
        text: entry?.s.name ?? "",
        fontSize: fs,
        bold: true,
        color: seriesColor(style, entry?.si ?? i, entry?.s.color),
        align: "center",
        valign: "middle",
        name: `header-${i}`,
      });
    });
  } else {
    // Stacked flanks: one legend of all series across the top, via the shared
    // wrapping row — the forked single-row walk marched chips off the canvas.
    nodes.push(...legendRow(cfg, style, plot.x, titleH, { maxX: cfg.width - 4, entries: legendEntries }));
  }

  // Value gridlines mirrored on both flanks, drawn behind the bars.
  if (decor.gridlines) {
    for (const tk of ticks) {
      if (tk <= 0) continue;
      const q = qOf(tk);
      for (const x of [leftEdge - q, rightEdge + q]) {
        nodes.push({
          kind: "line",
          x1: x,
          y1: plot.y,
          x2: x,
          y2: plot.y + plot.h,
          stroke: style.gridline,
          strokeWidth: 1,
          name: `gridline-${tk}`,
        });
      }
    }
  }

  // The gutter holds the category names, and it may have been scaled down above
  // to keep the bars a chart — so the names have to come with it or they spill
  // onto the bars they are naming (they are CENTRED in the gutter, so a name
  // wider than it overhangs both flanks). Shrunk together so every row reads at
  // one size, then clipped; at any size where the names already fit, this is
  // `fs` and nothing moves.
  const catFs = (() => {
    let f = fs;
    while (f > 5 && data.categories.some((c) => textWidth(c, f) > gutterW)) f -= 0.5;
    return f;
  })();

  // A category name is bounded by BOTH the gutter it is centred in and the row it
  // names: `catFs` is the width, `rowFs` the height. Fitting only the width left
  // the names overlapping each other vertically at a big font, which is the same
  // defect as the value labels beside them and wants the same bound.
  const nameFs = Math.min(catFs, rowFs);

  const columnTop: number[] = [];
  for (let c = 0; c < n; c++) {
    const cy = plot.y + slotH * (c + 0.5);
    columnTop.push(cy - barH / 2);
    // Category label in the center gutter.
    nodes.push({
      kind: "text",
      x: leftEdge,
      y: cy - nameFs * 0.75,
      w: gutterW,
      h: nameFs * 1.5,
      text: clipToWidth(data.categories[c], nameFs, gutterW),
      fontSize: nameFs,
      color: style.text,
      align: "center",
      valign: "middle",
      name: `category-${c}`,
    });
    const drawSide = (series: typeof withIdx, dir: -1 | 1, edge: number) => {
      let offset = 0;
      for (const { s, si } of series) {
        const v = s.values[c];
        if (v == null) continue;
        const len = qOf(v);
        const x = dir < 0 ? edge - offset - len : edge + offset;
        const fill = seriesColor(style, si, s.color);
        nodes.push({ kind: "rect", x, y: cy - barH / 2, w: len, h: barH, fill, name: `seg-${si}-${c}` });
        if (decor.segmentLabels) {
          const label = formatNumber(v, fmt);
          const single = series.length === 1;
          const inside = len >= textWidth(label, fs) + 4 && barH >= fs * 1.25;
          // Stacked segments only label when the value fits inside; single
          // flanks fall back to an outside label (classic behaviour).
          if (inside || single) {
            nodes.push({
              kind: "text",
              x: inside ? x : dir < 0 ? x - fs * 3.4 - 2 : x + len + 2,
              y: cy - rowFs * 0.75,
              w: inside ? len : fs * 3.4,
              h: rowFs * 1.5,
              text: label,
              fontSize: rowFs,
              color: inside ? contrastInk(fill) : style.text,
              align: inside ? "center" : dir < 0 ? "right" : "left",
              valign: "middle",
              name: `label-${si}-${c}`,
            });
          }
        }
        offset += len;
      }
    };
    drawSide(leftSeries, -1, leftEdge);
    drawSide(rightSeries, 1, rightEdge);
  }

  // Center axis lines flanking the gutter.
  for (const x of [leftEdge, rightEdge]) {
    nodes.push({
      kind: "line",
      x1: x,
      y1: plot.y,
      x2: x,
      y2: plot.y + plot.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
  }

  // Value tick labels on both flanks, in the reserved bottom strip.
  if (decor.valueAxis) {
    const ty = plot.y + plot.h + 1;
    for (const tk of ticks) {
      const q = qOf(tk);
      const label = formatNumber(tk, fmt);
      // 0 sits at the inner edges (the gutter sides); other ticks mirror outward.
      const xs = tk === 0 ? [leftEdge, rightEdge] : [leftEdge - q, rightEdge + q];
      xs.forEach((x, side) => {
        nodes.push({
          kind: "text",
          x: x - valueW / 2,
          y: ty,
          w: valueW,
          h: axisH,
          text: label,
          fontSize: fs * 0.85,
          color: style.mutedText,
          align: "center",
          valign: "middle",
          name: `tick-${tk}-${side === 0 ? "l" : "r"}`,
        });
      });
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map((_, c) => plot.y + slotH * (c + 0.5)),
      categoryWidth: data.categories.map(() => barH),
      columnTop,
      columnValue: data.categories.map((_, c) => signedSum(rightSeries, c) - signedSum(leftSeries, c)),
      baselineY: leftEdge,
      plot,
    },
  };
}
