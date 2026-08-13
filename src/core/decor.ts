import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors } from "./types";
import { textWidth, type SceneNode } from "./scene";
import { cagr, formatNumber, formatPercent } from "./format";
import { MIN_LABEL_FS, titleNode } from "./layout/frame";

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

  // Every decoration here anchors to the category geometry (categoryX / columnTop
  // / seriesLevels). A chart that drew NO segments — a single-0 pie or sunburst,
  // say — has empty anchor arrays, so any anchored read comes back undefined and
  // the decoration would emit NaN geometry. There is nothing to annotate.
  if (!a.categoryX.length) return nodes;

  // --- CAGR arrow: diagonal arrow between two column tops with "+x.x% p.a." ---
  if (decor.cagr) {
    const { from, to } = clampPair(decor.cagr, a);
    // Per-series CAGR when requested (think-cell computes on totals by default).
    const si = decor.cagr.series;
    const sVals = si != null ? cfg.data.series[si]?.values : undefined;
    const vFrom = sVals ? (sVals[from] ?? 0) : a.columnValue[from];
    const vTo = sVals ? (sVals[to] ?? 0) : a.columnValue[to];
    const rate = cagr(vFrom, vTo, to - from);
    // Anchored on the value it PRINTS, for the reason spelled out on the
    // difference arrow below — this block carries the identical mismatch, and
    // when `decor.cagr.series` names a series it is worse: the rate is that
    // series' and the arrow sat on the drawn top of whichever series happened
    // to be highest.
    const cagrY = (c: number, v: number) => (a.valueToY ? a.valueToY(sVals ? v : a.columnValue[c]) : a.columnTop[c]);
    /** How far the arrowhead reaches back from its tip — see the SVG renderer. */
    const ARROW = 5;
    // Clear the column totals row and difference arrows when shown — but never
    // by more than the chart above the columns can pay for. The lift is
    // legibility, not data: what the arrow claims is its two category anchors and
    // the DIFFERENCE between its ends, and shortening the lift moves both ends by
    // the same amount, so the slope and the anchors both survive. Letting it
    // stand put the arrowhead 4pt above the top of a 60pt-tall stacked column.
    const wantLift = fs * 1.6 + (decor.totals ? fs * 1.5 : 0) + (decor.difference ? fs * 1.2 : 0);
    const lift = Math.max(0, Math.min(wantLift, Math.min(cagrY(from, vFrom), cagrY(to, vTo)) - ARROW * 1.8));
    const x1 = a.categoryX[from];
    const y1 = cagrY(from, vFrom) - lift;
    const x2 = a.categoryX[to];
    const y2 = cagrY(to, vTo) - lift;
    const angle = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
    nodes.push(
      { kind: "line", x1, y1, x2, y2, stroke: style.text, strokeWidth: 1.25, name: "cagr-line" },
      { kind: "arrowhead", x: x2, y: y2, angle, size: ARROW, fill: style.text, name: "cagr-head" },
    );
    // The arrow is lifted clear of the column tops and the caption sits above the
    // arrow, so on a frame with no headroom the caption meets the TITLE. Its y is
    // decorative — it captions the arrow, where the arrow's endpoints are anchored
    // data — so it may be shrunk and dropped where they may not.
    //
    // CLAMPING it to the title's bottom was tried and measured and is still
    // refused: it turned five `title x cagr-label` overlaps into EIGHT against
    // the column totals, because a clamp moves a label whether or not the
    // destination is free. What that measurement did not try is making the
    // caption SMALLER, and it was the last label in the engine still drawn at
    // the full chart font in a FIXED 90pt box — wider than an 80pt chart.
    //
    // So: fit it to the gap between the title and the arrow, and past the floor
    // drop the caption while KEEPING the arrow. The arrow still shows the growth
    // and its two anchors; a rate drawn through the title is not readable
    // anyway. Same answer the ring, radar, tilemap and pie reservations give.
    const capText = rate == null ? "CAGR n/a" : `${formatPercent(rate, 1, true, cfg.numberFormat?.locale)} p.a.`;
    // Measured against the title's INK, not against `titleHeight`. The reserved
    // band is taller than the text in it, and the caption has always been free to
    // sit inside that band — bounding by the reservation dropped the caption on
    // 120x90 and 160x120 charts where nothing was overlapping at all.
    const tNode = titleNode(cfg, style) as { fontSize: number; text: string } | null;
    // A valign-top title's ink runs from its baseline (y + fontSize) down by the
    // descender, so its bottom is `fontSize * 1.21`.
    const titleInkBottom = tNode ? tNode.fontSize * 1.21 : 0;
    const titleInkRight = tNode ? textWidth(tNode.text, tNode.fontSize, true) : 0;
    // The caption is valign-bottom in a `cf * 1.4` box sitting `cf * 1.6` above
    // the arrow, so its ink top is `min(y1, y2) - cf * 1.25`. Clearing the title
    // means keeping that at or below the title's ink bottom.
    const capFits = (f: number) => Math.min(y1, y2) - f * 1.25 >= titleInkBottom;
    let cf = fs;
    while (cf > MIN_LABEL_FS && !capFits(cf)) cf -= 0.5;
    const capW0 = Math.min(90, cfg.width);
    const capX0 = Math.max(0, Math.min((x1 + x2) / 2 - capW0 / 2, cfg.width - capW0));
    // A caption sitting entirely to the RIGHT of the title's ink cannot collide
    // with it however tall it is, so it keeps its full size.
    const clearOfTitleX = capX0 + (capW0 - textWidth(capText, fs, true)) / 2 >= titleInkRight;
    if (clearOfTitleX) cf = fs;
    if (clearOfTitleX || capFits(cf)) {
      // Bounded by the frame, not a fixed 90pt box, and re-centred on the arrow
      // at whatever width it actually gets.
      const capW = capW0;
      nodes.push({
        kind: "text",
        x: capX0,
        // The box and the font move together, so a caption that needed no
        // shrinking lands exactly where it always did.
        y: Math.max(0, Math.min(y1, y2) - cf * 1.6),
        w: capW,
        h: cf * 1.4,
        text: capText,
        fontSize: cf,
        bold: true,
        color: style.text,
        align: "center",
        valign: "bottom",
        name: "cagr-label",
      });
    }
  }

  // --- Difference arrow: dashed level line + vertical arrow ---
  // Total (column totals) by default; a level difference arrow when `series`
  // names a series — it compares the cumulative stack level at that series.
  if (decor.difference) {
    const { from, to } = clampPair(decor.difference, a);
    const si = decor.difference.series;
    const useLevel =
      si != null && a.seriesLevels != null && a.valueToY != null && si >= 0 && si < (a.seriesLevels[0]?.length ?? 0);
    // The pixel anchor and the number the arrow prints have to describe the same
    // mark, so take BOTH from `columnValue` wherever the layout publishes a
    // value→y map. `columnTop` is the drawn top, and the two part company more
    // often than they look:
    //
    //  - a multi-series LINE chart publishes the topmost point as `columnTop`
    //    and the FIRST series' value as `columnValue`, so the arrow spanned one
    //    series' points and read the other's growth (+100% over a span that is
    //    +200%);
    //  - a STACKED column with a negative segment puts `columnTop` at the
    //    positive total and `columnValue` at the net, so +10/−4 → +10/−8 drew a
    //    ZERO-LENGTH arrow labelled −67%, and +10/−8 → +5/0 drew one pointing
    //    DOWN labelled +150%.
    //
    // Anchoring on the printed value fixes both without touching what any
    // layout publishes — and it is an identity wherever the two already agree,
    // which is every chart in the showcase.
    const anchorY = (c: number) => (a.valueToY ? a.valueToY(a.columnValue[c]) : a.columnTop[c]);
    let vFrom = useLevel ? a.seriesLevels![from][si!] : a.columnValue[from];
    let yFrom = useLevel ? a.valueToY!(vFrom) : anchorY(from);
    const vTo = useLevel ? a.seriesLevels![to][si!] : a.columnValue[to];
    const yTo = useLevel ? a.valueToY!(vTo) : anchorY(to);
    // Anchor the arrow's start at a value line instead of a column.
    const vlIdx = decor.difference.fromValueLine;
    const vls = decor.valueLines ?? (decor.valueLine ? [decor.valueLine] : []);
    if (vlIdx != null && vls[vlIdx] && a.valueToY) {
      const vl = vls[vlIdx];
      vFrom = vl.mode === "mean" ? a.columnValue.reduce((s, v) => s + v, 0) / (a.columnValue.length || 1) : vl.value;
      yFrom = a.valueToY(vFrom);
    }
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
        ? formatPercent(vTo / vFrom - 1, 0, true, cfg.numberFormat?.locale)
        : formatNumber(vTo - vFrom, { ...cfg.numberFormat, forceSign: true });
    // The label reads to the RIGHT of the arrow, in a margin `computeFrame`
    // reserves for it — but only the cartesian frame reserves one. A line chart
    // puts its last category hard against the plot edge, so on an ordinary
    // 400x300 two-point line the label started 8pt past the chart and the
    // frame-clip at the end of `buildChart` cut "+100%" down to "+…". Flip it to
    // the other side of the arrow when the right has no room: the ARROW is what
    // anchors the claim, and the caption says the same thing from either side.
    const labelW = textWidth(label, fs, true) + 2;
    const flip = x + 3 + labelW > cfg.width;
    nodes.push({
      kind: "text",
      x: flip ? Math.max(0, x - 3 - labelW) : x + 3,
      y: (yFrom + yTo) / 2 - fs * 0.75,
      w: flip ? labelW : Math.max(30, cfg.width - x - 3),
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
        vl.mode === "mean" ? a.columnValue.reduce((s, v) => s + v, 0) / (a.columnValue.length || 1) : vl.value;
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

  // --- Speech-bubble callouts: a comment anchored to a column or level ---
  decor.callouts?.forEach((co, i) => {
    const c = Math.max(0, Math.min(a.categoryX.length - 1, co.category));
    const ax = a.categoryX[c];
    const useLevel =
      co.series != null &&
      a.seriesLevels != null &&
      a.valueToY != null &&
      co.series >= 0 &&
      co.series < (a.seriesLevels[c]?.length ?? 0);
    const ay = useLevel ? a.valueToY!(a.seriesLevels![c][co.series!]) : a.columnTop[c];
    const w = textWidth(co.text, fs) + fs * 1.2;
    const h = fs * 1.9;
    // Bubble center defaults to hovering above the anchor — but "above" is only
    // available when the anchor has room above it. A 100% or Mekko chart exposes
    // no valueToY, so EVERY callout falls back to columnTop, which on those kinds
    // is the plot ceiling for every column by construction; the same happens on
    // any chart when the callout names the tallest column. Lifting 4.2 font sizes
    // from there put the box, its text and its tail entirely off the top of the
    // canvas. Keep the bubble on the canvas — a bubble overlapping its column
    // still reads; one at y = -30 is simply lost.
    const clamp = (v: number, lo: number, hi: number) => (lo > hi ? (lo + hi) / 2 : Math.min(hi, Math.max(lo, v)));
    const bx = clamp(ax + (co.dx ?? 0), w / 2, cfg.width - w / 2);
    const by = clamp(ay - fs * 4.2 + (co.dy ?? 0), h / 2, cfg.height - h / 2);
    nodes.push(
      {
        kind: "line",
        x1: bx,
        y1: by + h / 2 - 1,
        x2: ax,
        y2: ay - 2,
        stroke: style.text,
        strokeWidth: 0.75,
        name: `callout-tail-${i}`,
      },
      {
        kind: "rect",
        x: bx - w / 2,
        y: by - h / 2,
        w,
        h,
        fill: style.background,
        stroke: style.text,
        strokeWidth: 1,
        name: `callout-box-${i}`,
      },
      {
        kind: "text",
        x: bx - w / 2,
        y: by - h / 2,
        w,
        h,
        text: co.text,
        fontSize: fs,
        color: style.text,
        align: "center",
        valign: "middle",
        name: `callout-text-${i}`,
      },
    );
  });

  return nodes;
}

/**
 * Shaded background bands highlighting an axis region — drawn BEHIND the
 * data (the caller prepends these to the scene). axis "y" spans a value
 * range; axis "x" spans category indices.
 */
export function bandNodes(cfg: ChartConfig, style: ChartStyle, decor: Decorations, a: LayoutAnchors): SceneNode[] {
  const nodes: SceneNode[] = [];
  const fs = style.fontSize;
  decor.bands?.forEach((band, i) => {
    const fill = band.color ?? "#f2f1ec";
    let r: { x: number; y: number; w: number; h: number } | null = null;
    // `axis` names the DATA axis the band spans, not a screen direction: a "y"
    // band is a range of VALUES and a "x" band is a range of CATEGORIES,
    // whichever way the chart is turned. Reading `valueToY` alone meant the
    // range zones of a bullet chart — conventionally horizontal — were dropped
    // silently along with its target tick.
    const H = !!cfg.horizontal;
    const valueToPos = H ? a.valueToX : a.valueToY;
    if (band.axis === "y" && valueToPos) {
      // Clip to the plot: band.from/to are data values that may fall outside the
      // value domain, and the map extrapolates past the axis, so an unclamped
      // band renders off-frame. (The category branch below clamps its indices;
      // a band entirely outside the plot collapses to w/h <= 0 and is dropped.)
      const v1 = valueToPos(band.from);
      const v2 = valueToPos(band.to);
      if (H) {
        const left = Math.max(a.plot.x, Math.min(v1, v2));
        const right = Math.min(a.plot.x + a.plot.w, Math.max(v1, v2));
        r = { x: left, y: a.plot.y, w: right - left, h: a.plot.h };
      } else {
        const top = Math.max(a.plot.y, Math.min(v1, v2));
        const bot = Math.min(a.plot.y + a.plot.h, Math.max(v1, v2));
        r = { x: a.plot.x, y: top, w: a.plot.w, h: bot - top };
      }
    } else if (band.axis === "x" && a.categoryX.length) {
      const c1 = Math.max(0, Math.min(a.categoryX.length - 1, Math.min(band.from, band.to)));
      const c2 = Math.max(0, Math.min(a.categoryX.length - 1, Math.max(band.from, band.to)));
      // `categoryX` is the category's centre on the CROSS axis — x on a column
      // chart, y on a bar.
      const m1 = a.categoryX[c1] - a.categoryWidth[c1] * 0.75;
      const m2 = a.categoryX[c2] + a.categoryWidth[c2] * 0.75;
      r = H ? { x: a.plot.x, y: m1, w: a.plot.w, h: m2 - m1 } : { x: m1, y: a.plot.y, w: m2 - m1, h: a.plot.h };
    }
    if (!r || r.w <= 0 || r.h <= 0) return;
    nodes.push({ kind: "rect", ...r, fill, name: `band-${i}` });
    if (band.label) {
      nodes.push({
        kind: "text",
        x: r.x + 3,
        y: r.y + 1,
        w: Math.max(20, r.w - 6),
        h: fs * 1.3,
        text: band.label,
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: "left",
        valign: "top",
        name: `band-label-${i}`,
      });
    }
  });
  return nodes;
}

function clampPair(p: { from: number; to: number }, a: LayoutAnchors): { from: number; to: number } {
  const n = a.categoryX.length;
  const from = Math.max(0, Math.min(n - 1, p.from));
  const to = Math.max(0, Math.min(n - 1, p.to));
  return from <= to ? { from, to } : { from: to, to: from };
}
