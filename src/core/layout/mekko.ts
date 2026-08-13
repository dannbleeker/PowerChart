import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, formatPercent, resolveFormat } from "../format";
import { seriesColor } from "../style";
import { bandFontSize, chromeNodes, computeFrame, computeFrameHorizontal, fitPlot, titleHeight } from "./frame";
import { legendRow, seriesLabelNodes, type LayoutResult } from "./column";
import { columnPositiveTotal } from "./totals";

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

  const totals = data.categories.map((_, c) => columnPositiveTotal(data.series, c));
  const extents = units ? data.categories.map((_, c) => Math.max(0, data.xExtent?.[c] ?? 0)) : totals;
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
    //
    // Back through `fitPlot`, because this widening happens AFTER the frame was
    // floored and takes a fixed number of points however narrow the chart is:
    // on a 14pt-wide mekko it left `w: -12.2`, and a negative-width segment is
    // not a thin one — SVG drops it and PowerPoint clamps it to a sliver, so
    // every row loses its bars while the totals beside them still print. Any
    // frame that can afford the gutter is returned untouched.
    const extra = textWidth(" (00%)", fs);
    frame = fitPlot(cfg, { ...frame, x: frame.x + extra, w: frame.w - extra });
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
      const r = H ? { x: acc, y: pos, w: segLen, h: ext } : { x: pos, y: acc - segLen, w: ext, h: segLen };
      acc = H ? acc + segLen : acc - segLen;
      const fill = seriesColor(style, si, s.color);
      // Transparent "no-fill" segment: it still occupies the column (`acc` has
      // moved past it) but draws nothing, floating the segments above — same
      // contract as the stacked column (column.ts). Emitting the rect made both
      // PowerPoint sinks paint it: the pptx maps a bare word to mid grey and
      // Office.js hands "transparent" to setSolidColor, which it rejects.
      if (fill === "transparent") return;
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
      // Bound by the row it ends, not by the chart's font: a Mekko's rows are
      // proportional to their totals, so a small column gets a thin row, and a
      // total centred in a box `fs * 1.5` tall then runs into its neighbours.
      // Per-row rather than one size for the axis, because the rows genuinely
      // differ in height here. Last resort: a row that can afford `fs` keeps it.
      //
      // A row can be ZERO here — the rows are shares of the total, so a chart
      // whose data is all zeroes has no rows at all — and `Math.min` answered a
      // font of zero, which OOXML rejects outright. `bandFontSize` is the floor.
      const totalFs = bandFontSize(fs, ext, 1.5);
      if (totalFs > 0)
        nodes.push({
          kind: "text",
          x: frame.x + colLen + 3,
          y: centers[c] - totalFs * 0.75,
          w: cfg.width - (frame.x + colLen) - 3,
          h: totalFs * 1.5,
          text: formatNumber(totals[c], fmt),
          fontSize: totalFs,
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
    ...chromeNodes(
      cfg,
      style,
      { ...decorFull, categoryAxis: false, valueAxis: false, gridlines: false },
      frame,
      centers,
    ),
  );
  if (decor.categoryAxis) {
    const catLabels = data.categories.map((cat, c) =>
      units ? cat : `${cat} (${formatPercent(extents[c] / grand, 0, false, cfg.numberFormat?.locale)})`,
    );
    // A Mekko's category label sits under its own COLUMN and is as wide as it
    // needs to be, so a label wider than its column overflows the box
    // symmetrically — into the neighbouring label, and off the left edge of the
    // chart for the first one. At an 22pt font the three-category sample had two
    // overlapping pairs and `category-0` starting at x = -2, which reads as one
    // run-on string: "EMEA (32%)Americas (42%)APAC (27%)".
    //
    // Shrunk together until the widest fits its own column, then clipped —
    // together, because category labels that differ in size read as a hierarchy
    // that is not there. Same two-step the agenda and the process flow already
    // use, and `catFs` stays at `fs` whenever they fit, so an ordinary chart is
    // untouched.
    let catFs = fs;
    // The room is the column plus the gap it shares with its neighbours — half
    // on each side, so adjacent boxes ABUT rather than overlap. It used to be
    // `+ 8` against a 2pt gap, which let every label bleed 4pt into each
    // neighbour: "Americas (42%)" and "APAC (27%)" collided by 6pt on an
    // ordinary 200x150 mekko, at the default font.
    const catRoom = (c: number) => widths[c] + gap;
    // Rotating the chart rotates which side of the label is crowded, and this
    // predicate used to answer `false` outright when horizontal — so the whole
    // fit above was skipped and the names were drawn at `fs`, unclipped, in a
    // gutter `computeFrameHorizontal` caps at 30% of the width. Sideways the
    // room is that gutter, and it is the same for every row, so the axis still
    // shrinks together and still reads at one size.
    const gutter = Math.max(1, frame.x - 4);
    const overflows = (f: number) =>
      H ? catLabels.some((l) => textWidth(l, f) > gutter) : catLabels.some((l, c) => textWidth(l, f) > catRoom(c));
    while (catFs > 6 && overflows(catFs)) catFs -= 0.5;
    for (let c = 0; c < n; c++) {
      const label = catLabels[c];
      if (H) {
        // Bounded by the gutter (width, shared) AND by the row it names
        // (height, per-row — a Mekko's rows are proportional to their totals,
        // so a small column gets a thin one). The smaller wins, the same way
        // the butterfly's category names take the smaller of their gutter fit
        // and their row fit.
        const nameFs = bandFontSize(catFs, catRoom(c), 1.5);
        if (nameFs > 0)
          nodes.push({
            kind: "text",
            x: 0,
            y: centers[c] - nameFs * 0.75,
            w: gutter,
            h: nameFs * 1.5,
            text: clipToWidth(label, nameFs, gutter),
            fontSize: nameFs,
            color: style.text,
            align: "right",
            valign: "middle",
            name: `category-${c}`,
          });
      } else {
        nodes.push({
          kind: "text",
          x: centers[c] - catRoom(c) / 2,
          y: frame.y + frame.h + 3,
          w: catRoom(c),
          h: catFs * 1.4,
          text: clipToWidth(label, catFs, catRoom(c)),
          fontSize: catFs,
          color: style.text,
          align: "center",
          valign: "top",
          name: `category-${c}`,
        });
      }
    }
  }
  if (H) {
    nodes.push({
      kind: "line",
      x1: frame.x,
      y1: frame.y,
      x2: frame.x,
      y2: frame.y + frame.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
    if (decor.seriesLabels && data.series.length > 1) {
      nodes.push(...legendRow(cfg, style, frame.x, titleHeight(cfg, style) + 2, { maxX: cfg.width - 4 }));
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
