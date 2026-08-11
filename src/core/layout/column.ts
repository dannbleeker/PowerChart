import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors, Series } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, niceTicks, resolveFormat, segmentLabel, axisTickLabel } from "../format";
import { seriesColor } from "../style";
import { lerpColor } from "../color";
import {
  baselineNode,
  breakMarkerNodes,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  legendWrapWalk,
  seriesLegendLabels,
  logFloor,
  valueScale,
  type Frame,
  type ValueScale,
} from "./frame";
// Combo base kinds. These modules import back from column (LayoutResult /
// horizontalChrome), but the calls happen at runtime so the ESM cycle resolves.
import { layoutWaterfall, detailParents } from "./waterfall";
import { layoutMekko } from "./mekko";
import { layoutLine } from "./line";
import { columnNegativeTotal, columnPositiveTotal, columnSignedTotal } from "./totals";
import { maxOf, minOf } from "../agg";

export interface LayoutResult {
  nodes: SceneNode[];
  anchors: LayoutAnchors;
}

/** Minimum segment thickness (relative to font size) before its label is hidden. */
const LABEL_FIT = 1.25;

/** How one series' rectangles are painted, once IBCS notation is applied. */
interface MarkPaint {
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  pattern?: Series["pattern"];
}

/**
 * IBCS scenario notation restyles a series by the data's nature: PY a lighter
 * solid, PL/BU an outlined/hollow bar, FC a hatch. AC/none stay solid, carrying
 * only the series' own `pattern`.
 *
 * Shared by the plotted segment and its legend chip: a hollow PL bar keyed by a
 * solid block — or two same-coloured series told apart only by a hatch, keyed by
 * two identical squares — hides exactly the redundant encoding the notation
 * exists to provide for greyscale and colour-blind readers.
 */
function markPaint(
  style: ChartStyle,
  fill: string,
  s: Pick<Series, "scenario" | "pattern">,
  baseStroke?: string,
  baseStrokeWidth?: number,
): MarkPaint {
  const paint: MarkPaint = { fill, stroke: baseStroke, strokeWidth: baseStrokeWidth, pattern: s.pattern };
  if (s.scenario === "PY") {
    paint.fill = lerpColor(style.background, fill, 0.5);
  } else if (s.scenario === "PL" || s.scenario === "BU") {
    paint.fill = "none";
    paint.stroke = fill;
    paint.strokeWidth = 1.5;
  } else if (s.scenario === "FC") {
    // SVG keeps the true IBCS hatch, but rect.pattern is SVG-only — both
    // PowerPoint renderers drop it, which made FC pixel-identical to AC in
    // the actual deliverable. Back the hatch with a fill/border encoding all
    // three renderers can express: a light tint (distinct from PY's heavier
    // 0.5) plus a series-coloured edge (distinct from PL/BU's hollow bar).
    paint.pattern = "diagonal";
    paint.fill = lerpColor(style.background, fill, 0.75);
    paint.stroke = fill;
    paint.strokeWidth = 1;
  }
  return paint;
}

/**
 * The value a CLUSTERED category publishes to every decoration that reads one.
 *
 * This was `Math.max(0, ...values)`, and the zero seed is the bug: a category
 * whose bars are all negative published 0. `columnValue` is what mean lines,
 * difference arrows and the rest print, so a cash-flow chart of
 * [-40, -20, -10, 60] drew its mean line ABOVE the baseline and labelled it
 * "Ø 15" — the mean of [0, 0, 0, 60] — while the same data as `stacked`, `line`
 * or `area` correctly said "Ø -2.5". A difference arrow between two all-negative
 * categories came out zero-length and labelled "0", sitting beside the very
 * totals that contradict it.
 *
 * Zero is right for the drawn column's TOP — a bar hanging below the baseline
 * tops out at the baseline, and `columnTop` still clamps that way. It is wrong
 * for the column's VALUE, which is what a reader is being told the category is
 * worth. So: the extreme in the direction the column actually goes — the
 * highest bar when anything rises, the lowest when nothing does.
 */
export function clusteredTopValue(values: number[]): number {
  const finite = values.filter((v) => Number.isFinite(v));
  if (!finite.length) return 0;
  const hi = Math.max(...finite);
  return hi > 0 ? hi : Math.min(...finite);
}

/**
 * Stacked / clustered / 100% column charts — and, following think-cell's
 * "a bar chart is a rotated column chart" model, the same layouts in
 * horizontal orientation when cfg.horizontal is set.
 */
export function layoutColumns(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data, kind } = cfg;
  const n = data.categories.length;
  const stacked = kind !== "clustered";
  const pct = kind === "stacked100";
  const H = !!cfg.horizontal;
  // Clustered-stacked: series carry stack indices (blank datasheet rows).
  const stackIds = [...new Set(data.series.map((s) => s.stack ?? 0))].sort((a, b) => a - b);
  const nStacks = stacked && !pct ? stackIds.length : 1;
  const stackPos = new Map(stackIds.map((id, i) => [id, i]));

  const vframe = H ? null : computeFrame(cfg, style, decor, decor.seriesLabels ? data.series.map((s) => s.name) : []);
  const frame = H ? computeFrameHorizontal(cfg, style, decor) : vframe!.frame;
  const fs = style.fontSize;

  // Category slots run along x (vertical) or y (horizontal).
  const catStart = H ? frame.y : frame.x;
  const catLen = H ? frame.h : frame.w;
  const slotLen = catLen / Math.max(1, n);
  // Excel-style gap width: gap between columns as a % of column width.
  // Default 50 reproduces think-cell's 2/3-of-slot columns.
  const gapWidth = Math.max(0, Math.min(500, cfg.gapWidth ?? 50));
  const colThick = slotLen / (1 + gapWidth / 100);
  // Excel-style clustered overlap (−100…100): fraction each bar overlaps its
  // neighbour. 0 = edge to edge (the historical default).
  const overlapFrac = Math.max(-100, Math.min(100, cfg.overlap ?? 0)) / 100;
  const centers = Array.from({ length: n }, (_, i) => catStart + slotLen * (i + 0.5));

  const posTotals = data.categories.map((_, c) => columnPositiveTotal(data.series, c));
  const negTotals = data.categories.map((_, c) => columnNegativeTotal(data.series, c));
  const signedTotals = data.categories.map((_, c) => columnSignedTotal(data.series, c));
  // Per-category denominator for 100% charts (think-cell's "100%=" row).
  const denominators = data.categories.map((_, c) => {
    const d = data.hundredPercent?.[c];
    if (d != null && d > 0) return d;
    // An all-negative category has no positive total; normalise against the
    // negative magnitude so its segments fill down to -100% instead of
    // collapsing to zero (v / 0).
    return posTotals[c] > 0 ? posTotals[c] : -negTotals[c];
  });
  const fmt = resolveFormat(
    [...data.series.flatMap((s) => s.values.filter((v): v is number => v != null)), ...signedTotals],
    cfg.numberFormat,
  );

  // Per-stack totals (clustered-stacked scales to the tallest single stack).
  const stackPosTotals = (c: number, id: number) =>
    data.series.reduce((a, s) => a + ((s.stack ?? 0) === id ? Math.max(0, s.values[c] ?? 0) : 0), 0);
  const stackNegTotals = (c: number, id: number) =>
    data.series.reduce((a, s) => a + ((s.stack ?? 0) === id ? Math.min(0, s.values[c] ?? 0) : 0), 0);

  const allValues = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
  const logOn = !stacked && !H && !!cfg.logScale;
  let dataMin: number, dataMax: number;
  if (pct) {
    dataMin = 0;
    dataMax = 1;
  } else if (stacked && nStacks > 1) {
    // Folded, not spread: this array is categories x stacks (up to 4096 x 256),
    // which overflows the argument list and throws RangeError. The sibling
    // branches below spread at most MAX_CATEGORIES or MAX_SERIES values, so they
    // stay safe as-is.
    dataMin = minOf(
      data.categories.flatMap((_, c) => stackIds.map((id) => stackNegTotals(c, id))),
      0,
    );
    dataMax = maxOf(
      data.categories.flatMap((_, c) => stackIds.map((id) => stackPosTotals(c, id))),
      0,
    );
  } else if (stacked) {
    dataMin = Math.min(0, ...negTotals);
    dataMax = Math.max(0, ...posTotals);
  } else {
    dataMin = minOf(allValues, 0);
    dataMax = maxOf(allValues, 0);
  }
  // 100% charts: negatives are shares below the zero line. The axis drops to
  // the most-negative column share (0 when all data is positive → unchanged).
  const pctNegMin = pct
    ? Math.min(0, ...data.categories.map((_, c) => (denominators[c] > 0 ? negTotals[c] / denominators[c] : 0)))
    : 0;
  // …and the ceiling is not always 100%. An authored "100% =" denominator can be
  // SMALLER than the column it normalises — a multi-select survey ("100% = 500
  // respondents", answers totalling 180%), or a typo. Pinning max at 1 painted
  // those segments outside the plot and outside the chart frame; `waffle` already
  // handles the same input. Left at exactly 1 whenever no column overflows, so
  // every well-formed 100% chart is untouched.
  const pctPosMax = pct
    ? Math.max(1, ...data.categories.map((_, c) => (denominators[c] > 0 ? posTotals[c] / denominators[c] : 0)))
    : 1;
  const scale: ValueScale = pct
    ? {
        min: pctNegMin,
        max: pctPosMax,
        percent: true,
        // Clamped to the domain, the way valueScale already filters its own:
        // niceTicks rounds OUTWARD, so it returned ticks below pctNegMin, and
        // toY mapped those past the plot — a gridline at y=301.6 on a 300pt
        // canvas, with its label sitting on the category row.
        ticks:
          pctNegMin < 0 || pctPosMax > 1
            ? niceTicks(pctNegMin, pctPosMax, 5).filter((t) => t >= pctNegMin - 1e-9 && t <= pctPosMax + 1e-9)
            : [0, 0.25, 0.5, 0.75, 1],
        toY: (v: number) => frame.y + frame.h - ((v - pctNegMin) / (pctPosMax - pctNegMin)) * frame.h,
      }
    : valueScale(
        frame,
        logOn ? logFloor(allValues, dataMin) : dataMin,
        dataMax,
        cfg.scale,
        H ? undefined : cfg.axisBreak,
        logOn,
      );

  // Value coordinate: distance along the value axis from the scale minimum.
  // Vertical charts route through toY so axis breaks apply; horizontal stays linear.
  //
  // CLAMPED either way. `toY` clips to the plot and says why (frame.ts: "every
  // charting tool draws a value above the axis maximum as a bar reaching the top
  // of the plot"), so the vertical branch inherited that and the horizontal one
  // — its own linear map — did not. Pin the axis to 0–100 on revenue data, which
  // the pane invites since "Axis scale min / max" are free-text boxes, and the
  // same config that draws a full-height bar upright drew a 30,128pt one
  // sideways: off the slide, and past the OOXML coordinate limit that makes
  // PowerPoint offer to repair the file.
  const valLen = H ? frame.w : frame.h;
  const qOf = H
    ? (v: number) => Math.max(0, Math.min(valLen, ((v - scale.min) / (scale.max - scale.min || 1)) * valLen))
    : (v: number) => frame.y + frame.h - scale.toY(v);
  /** Rect spanning [v0, v1] on the value axis at category position/thickness. */
  const segRect = (catPos: number, thick: number, v0: number, v1: number) => {
    const q0 = Math.min(qOf(v0), qOf(v1));
    const q1 = Math.max(qOf(v0), qOf(v1));
    return H
      ? { x: frame.x + q0, y: catPos - thick / 2, w: q1 - q0, h: thick }
      : { x: catPos - thick / 2, y: frame.y + frame.h - q1, w: thick, h: q1 - q0 };
  };

  const nodes: SceneNode[] = H
    ? horizontalChrome(cfg, style, decor, frame, centers, scale, qOf)
    : chromeNodes(cfg, style, decor, frame, centers, scale);
  const zeroQ = qOf(0);
  const y0 = H ? frame.x + zeroQ : frame.y + frame.h - zeroQ;
  const columnTop: number[] = [];
  /**
   * The DATA value `columnTop` is drawn at, per category.
   *
   * The pixel anchor and the number a decoration prints have to describe the
   * same mark, and on a clustered chart they did not: `columnValue` published
   * the SUM of every series while `columnTop` is the tallest single bar, which
   * is drawn nowhere near it. Its sibling `seriesLevels` was fixed for exactly
   * this ("clustered, each bar stands on the baseline") and this one was not
   * swept with it.
   */
  const drawnTopValue: number[] = [];
  const seriesLevels: number[][] = [];
  /** Segment mid-position of the last category per series, for series labels. */
  const lastSegMid: (number | null)[] = data.series.map(() => null);
  /** Cumulative segment boundaries per column (value units), for connectors. */
  const posBounds: number[][] = [];
  const negBounds: number[][] = [];

  for (let c = 0; c < n; c++) {
    // Running positive/negative levels per stack group (value units).
    const ups = stackIds.map(() => 0);
    const downs = stackIds.map(() => 0);
    const levels: number[] = data.series.map(() => 0);
    const stackThick = colThick / nStacks;
    // Clustered bars fill the column; overlap widens each bar and shrinks the
    // stride so they overlap (or gap). At overlap 0 this is colThick / nBars.
    const nBars = Math.max(1, data.series.length);
    const barW = stacked ? stackThick : colThick / (1 + (nBars - 1) * (1 - overlapFrac));
    const barStep = stacked ? stackThick : barW * (1 - overlapFrac);
    const barThick = barW;
    // think-cell's Segment Order: stacking order within this column.
    const order = data.series.map((_, i) => i);
    if (cfg.segmentOrder === "reverse") order.reverse();
    else if (cfg.segmentOrder === "ascending" || cfg.segmentOrder === "descending") {
      const sign = cfg.segmentOrder === "ascending" ? 1 : -1;
      order.sort((a, b) => sign * ((data.series[a].values[c] ?? 0) - (data.series[b].values[c] ?? 0)));
    }

    order.forEach((si, position) => {
      const s = data.series[si];
      const raw = s.values[c];
      let v = raw ?? 0;
      if (pct) v = denominators[c] > 0 ? v / denominators[c] : 0;
      let r: { x: number; y: number; w: number; h: number } | null = null;
      const fill = seriesColor(style, si, s.colors?.[c] ?? s.color);

      const sp = stackPos.get(s.stack ?? 0) ?? 0;
      // The `!H` that used to be here forced every clustered chart back to
      // plain rectangles the moment the rotation toggle went on — and a
      // Cleveland dot plot and a dumbbell range chart are NORMALLY horizontal,
      // which is the whole point of them: long category labels run down the
      // left. `barStyle: "range"` sideways gave two full-length bars from zero
      // instead of dot–line–dot, so the connector carrying the entire meaning
      // was gone and the endpoints read as two independent magnitudes. Stacked
      // is still excluded, because a stack has no single point to mark.
      const barStyle = stacked ? "bar" : (decor.barStyle ?? "bar");
      if (raw != null && v !== 0) {
        if (stacked) {
          const catPos = nStacks > 1 ? centers[c] - colThick / 2 + (sp + 0.5) * stackThick : centers[c];
          const thick = nStacks > 1 ? stackThick - 1 : colThick;
          if (v >= 0) {
            r = segRect(catPos, thick, ups[sp], ups[sp] + v);
            ups[sp] += v;
            // Key the boundary by SERIES, not push order: a zero/null segment
            // pushes nothing, which used to shift every later boundary down an
            // index and join mismatched series between columns.
            if (nStacks === 1) (posBounds[c] ??= [])[si] = ups[sp];
          } else {
            r = segRect(catPos, thick, downs[sp] + v, downs[sp]);
            downs[sp] += v;
            if (nStacks === 1) (negBounds[c] ??= [])[si] = downs[sp];
          }
        } else {
          const pos = centers[c] - colThick / 2 + barW / 2 + position * barStep;
          r = segRect(pos, barThick - 1, 0, v);
        }
      }
      // Where THIS series' mark tops out, in value units — what a level-anchored
      // decoration (difference arrow, callout) points at. Stacked, that is the
      // cumulative level; clustered, each bar stands on the baseline, so it is
      // the bar's own value. Publishing the cumulative sum for clustered too left
      // it 0 for every series (ups/downs only advance in the stacked branch), so
      // the arrow silently collapsed onto the axis instead of the bar top.
      levels[si] = stacked ? ups[sp] + downs[sp] : v;
      if (!r) return;

      if (barStyle !== "bar") {
        // Lollipop / dot / dumbbell-range rendering for clustered charts:
        // the value point is a dot; lollipops add a stem from the baseline;
        // range connects the two series' dots with a line (drawn once).
        // Two coordinates rather than an x and a y: `cross` is the position on
        // the CATEGORY axis (a bar's x on a column chart, its y on a bar) and
        // `val` is the position on the VALUE axis. The rest of the block places
        // marks in those terms and maps them once, which is what lets the same
        // dot, stem and connector serve both orientations.
        const cross = barStyle === "range" ? centers[c] : H ? r.y + r.h / 2 : r.x + r.w / 2;
        const val = H ? (v >= 0 ? r.x + r.w : r.x) : v >= 0 ? r.y : r.y + r.h;
        const dotR = 4;
        /** A line along the value axis, at `cross` on the category axis. */
        const alongValue = (v1: number, v2: number) =>
          H ? { x1: v1, y1: cross, x2: v2, y2: cross } : { x1: cross, y1: v1, x2: cross, y2: v2 };
        if (barStyle === "lollipop") {
          nodes.push({
            kind: "line",
            // `y0` is the baseline in whichever direction the value axis runs.
            ...alongValue(y0, val),
            stroke: fill,
            strokeWidth: 1.5,
            name: `stem-${si}-${c}`,
          });
        }
        if (barStyle === "range" && si === 1 && data.series[0].values[c] != null) {
          const other = H
            ? frame.x + qOf(data.series[0].values[c]!)
            : frame.y + frame.h - qOf(data.series[0].values[c]!);
          nodes.push({
            kind: "line",
            ...alongValue(other, val),
            stroke: style.mutedText,
            strokeWidth: 1.5,
            name: `range-${c}`,
          });
        }
        nodes.push({
          kind: "ellipse",
          cx: H ? val : cross,
          cy: H ? cross : val,
          rx: dotR,
          ry: dotR,
          fill,
          stroke: style.background,
          strokeWidth: 1,
          name: `seg-${si}-${c}`,
        });
        if (c === n - 1) lastSegMid[si] = H ? cross : val;
        if (decor.segmentLabels) {
          const label = formatNumber(raw!, fmt);
          nodes.push({
            kind: "text",
            // Clear of the dot along the value axis, centred on it across.
            x: (H ? val : cross) + dotR + 2,
            y: (H ? cross : val) - fs * 0.7,
            w: textWidth(label, fs) + 4,
            h: fs * 1.4,
            text: label,
            fontSize: fs,
            color: style.text,
            align: "left",
            valign: "middle",
            name: `label-${si}-${c}`,
          });
        }
        return;
      }

      // Transparent "no-fill" segment: it still occupies the stack (the level
      // was already advanced) but draws nothing, floating the segments above.
      if (fill === "transparent") return;
      const seg = markPaint(style, fill, s, style.background, stacked ? 0.75 : 0);
      nodes.push({
        kind: "rect",
        ...r,
        fill: seg.fill,
        stroke: seg.stroke,
        strokeWidth: seg.strokeWidth,
        pattern: seg.pattern,
        name: `seg-${si}-${c}`,
      });
      if (c === n - 1) lastSegMid[si] = H ? r.x + r.w : r.y + r.h / 2;

      if (decor.segmentLabels) {
        // think-cell's label-content dropdown: value / % / series / category.
        const label = segmentLabel(decor.labelContent ?? (pct ? ["percent"] : ["value"]), {
          value: raw!,
          fraction: pct ? v : posTotals[c] > 0 ? Math.max(0, raw!) / posTotals[c] : null,
          series: s.name,
          category: data.categories[c],
          fmt,
        });
        const along = H ? r.w : r.h; // extent along the value axis
        const across = H ? r.h : r.w;
        const fits = H
          ? along >= textWidth(label, fs) + 2 && across >= fs * LABEL_FIT
          : along >= fs * LABEL_FIT && textWidth(label, fs) <= across + 2;
        if (fits) {
          nodes.push({
            kind: "text",
            x: r.x - 4,
            y: r.y + r.h / 2 - fs * 0.75,
            w: r.w + 8,
            h: fs * 1.5,
            text: label,
            fontSize: fs,
            // Against the painted fill, not the original series colour: a hollow
            // PL/BU bar shows the canvas, and a PY/FC tint is far lighter than `fill`.
            color: contrastInk(seg.fill === "none" ? style.background : seg.fill),
            align: "center",
            valign: "middle",
            name: `label-${si}-${c}`,
          });
        }
      }
    });
    seriesLevels.push(levels);

    const topV = pct
      ? denominators[c] > 0
        ? posTotals[c] / denominators[c]
        : 0
      : stacked
        ? Math.max(...ups)
        : clusteredTopValue(data.series.map((s) => s.values[c] ?? 0));
    const topQ = qOf(Math.max(0, topV));
    columnTop.push(H ? frame.x + topQ : frame.y + frame.h - topQ);
    drawnTopValue.push(topV);

    // Clustered-stacked: one total per stack sub-column (vertical only).
    if (decor.totals && !pct && nStacks > 1 && !H) {
      stackIds.forEach((id, sp) => {
        const subX = centers[c] - colThick / 2 + sp * stackThick;
        const subTopQ = qOf(Math.max(0, ups[sp]));
        const total = data.series.reduce((a, s) => a + ((s.stack ?? 0) === id ? (s.values[c] ?? 0) : 0), 0);
        nodes.push({
          kind: "text",
          x: subX - 4,
          y: frame.y + frame.h - subTopQ - fs * 1.45,
          w: stackThick + 8,
          h: fs * 1.4,
          text: formatNumber(total, fmt),
          fontSize: fs * 0.95,
          bold: true,
          color: style.text,
          align: "center",
          valign: "bottom",
          name: `total-${c}-s${sp}`,
        });
      });
    } else if (decor.totals && !pct) {
      if (H) {
        nodes.push({
          kind: "text",
          x: frame.x + topQ + 3,
          y: centers[c] - fs * 0.75,
          w: cfg.width - (frame.x + topQ) - 3,
          h: fs * 1.5,
          text: formatNumber(signedTotals[c], fmt),
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
          x: centers[c] - slotLen / 2,
          y: frame.y + frame.h - topQ - fs * 1.45,
          w: slotLen,
          h: fs * 1.4,
          text: formatNumber(signedTotals[c], fmt),
          fontSize: fs,
          bold: true,
          color: style.text,
          align: "center",
          valign: "bottom",
          name: `total-${c}`,
        });
      }
    }
  }

  // Grand total (think-cell 14): one label at the top-right showing the sum of
  // every category total. A FIXED anchor in the de-collision pass, so a tall
  // right-hand column's own (movable) total nudges around it. Vertical only, and
  // never on a 100% chart, where every column totals the same 100%.
  if (decor.grandTotal && !pct && !H && n > 0) {
    const grand = signedTotals.reduce((a, b) => a + b, 0);
    const gtext = formatNumber(grand, fmt);
    const gw = Math.min(frame.w, textWidth(gtext, fs, true) + 8);
    nodes.push({
      kind: "text",
      x: frame.x + frame.w - gw,
      y: frame.y - fs * 1.5,
      w: gw,
      h: fs * 1.4,
      text: gtext,
      fontSize: fs,
      bold: true,
      color: style.text,
      align: "right",
      valign: "bottom",
      name: "grand-total",
    });
  }

  // IBCS variance tier: a strip below the columns showing an actual series'
  // deviation from a reference (plan / previous year) as signed bars from a zero
  // line. Vertical only; the band height was reserved in computeFrame.
  if (decor.variance && !H && vframe && vframe.res.varianceH > 0) {
    const { actual, reference, mode = "absolute", goodIsUp = true } = decor.variance;
    const av = data.series[actual]?.values ?? [];
    const rv = data.series[reference]?.values ?? [];
    const deltas = data.categories.map((_, c) => {
      const a = av[c];
      const r = rv[c];
      if (a == null || r == null) return null;
      return mode === "percent" ? (r !== 0 ? ((a - r) / Math.abs(r)) * 100 : null) : a - r;
    });
    const maxAbs = maxOf(
      deltas.filter((d): d is number => d != null).map((d) => Math.abs(d)),
      0,
    );
    const bandTop = frame.y + frame.h;
    const bandH = vframe.res.varianceH;
    const zeroY = bandTop + bandH * 0.5;
    const halfH = bandH * 0.24; // leaves room for the delta label at the bar tip
    const scale = maxAbs > 0 ? halfH / maxAbs : 0;
    nodes.push({
      kind: "line",
      x1: frame.x,
      y1: zeroY,
      x2: frame.x + frame.w,
      y2: zeroY,
      stroke: style.axis,
      strokeWidth: 1,
      name: "variance-zero",
    });
    const GOOD = "#0ca30c";
    deltas.forEach((d, c) => {
      if (d == null) return;
      const favorable = goodIsUp ? d >= 0 : d <= 0;
      const color = d === 0 ? style.mutedText : favorable ? GOOD : style.negative;
      const barLen = Math.abs(d) * scale;
      const barW = Math.min(colThick, slotLen * 0.5);
      const up = d >= 0;
      const barY = up ? zeroY - barLen : zeroY;
      nodes.push({
        kind: "rect",
        x: centers[c] - barW / 2,
        y: barY,
        w: barW,
        h: barLen,
        fill: color,
        name: `variance-bar-${c}`,
      });
      const text = (d >= 0 ? "+" : "") + (mode === "percent" ? `${Math.round(d)}%` : formatNumber(d, fmt));
      nodes.push({
        kind: "text",
        x: centers[c] - slotLen / 2,
        y: up ? barY - fs * 1.05 : barY + barLen + 1,
        w: slotLen,
        h: fs * 1.05,
        text,
        fontSize: fs * 0.85,
        bold: true,
        color,
        align: "center",
        valign: up ? "bottom" : "top",
        name: `variance-label-${c}`,
      });
    });
  }

  // Connector lines between adjacent stacked columns: one per segment
  // boundary, so the development of each segment is easy to follow.
  if (decor.connectors && stacked && nStacks === 1) {
    const edge = (c: number, q: number, side: 1 | -1) =>
      H
        ? { x: frame.x + q, y: centers[c] + (side * colThick) / 2 }
        : { x: centers[c] + (side * colThick) / 2, y: frame.y + frame.h - q };
    for (let c = 0; c < n - 1; c++) {
      for (const bounds of [posBounds, negBounds]) {
        const a = bounds[c] ?? [];
        const b = bounds[c + 1] ?? [];
        // Sparse by series: only join a boundary that exists on both columns.
        for (let i = 0; i < Math.max(a.length, b.length); i++) {
          if (a[i] == null || b[i] == null) continue;
          const p1 = edge(c, qOf(a[i]), 1);
          const p2 = edge(c + 1, qOf(b[i]), -1);
          nodes.push({
            kind: "line",
            x1: p1.x,
            y1: p1.y,
            x2: p2.x,
            y2: p2.y,
            stroke: style.mutedText,
            strokeWidth: 0.75,
            name: `connector-${c}-${i}${bounds === negBounds ? "n" : ""}`,
          });
        }
      }
    }
  }

  if (!H) nodes.push(...breakMarkerNodes(frame, scale, style));

  // Zero baseline: horizontal line (vertical charts) or vertical line (bars).
  if (H) {
    nodes.push({
      kind: "line",
      x1: y0,
      y1: frame.y,
      x2: y0,
      y2: frame.y + frame.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    });
  } else {
    nodes.push(baselineNode(frame, y0, style));
  }

  if (decor.seriesLabels && !H) {
    nodes.push(...seriesLabelNodes(cfg, style, frame, lastSegMid as (number | null)[]));
  }

  return {
    nodes,
    anchors: {
      categoryX: centers,
      categoryWidth: data.categories.map(() => colThick),
      columnTop,
      // Stacked keeps the signed total, which IS the drawn column top there.
      // Clustered takes the value `columnTop` was placed at — the tallest bar —
      // because the sum of a clustered category is drawn nowhere on the chart.
      // Two consumers put ink at whatever this says: the Error-row whiskers
      // (`base = a.columnValue[c]`) and `decorations.valueLines: [{mode:"mean"}]`.
      // With the sum, a 3-category 2-series chart anchored its whiskers at
      // y = -3, -80 and -63 on a 300pt canvas, and its mean line at y = -141
      // labelled "Ø 46" against a value axis topping out near 30 — real shapes
      // on the slide ABOVE the chart, since the Office renderer applies no clamp.
      // `decorations.difference`/`cagr` were hit more quietly: their Y came from
      // `columnTop` and their NUMBER from here, so the arrow sat over a +25% rise
      // and read "+29%".
      columnValue: pct ? posTotals : stacked ? signedTotals : drawnTopValue,
      seriesLevels,
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: pct || H ? undefined : scale.toY,
      valueToX: H && !pct ? (v: number) => frame.x + qOf(v) : undefined,
    },
  };
}

/**
 * Combo chart, think-cell style: stacked columns plus line series drawn over
 * them on the same value axis. Series with `type: "line"` become lines; if
 * none is marked, the last series does.
 */
export function layoutCombo(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  /** Drawn over the columns rather than as one: a line, or bare markers. */
  const isOverlay = (s: (typeof cfg.data.series)[number]) => s.type === "line" || s.type === "marker";
  const marked = cfg.data.series.some(isOverlay);
  const nSeries = cfg.data.series.length;
  // Unmarked combo: the last series is the line, the rest are columns. A lone
  // series is a plain column, not both a column *and* a line (which would
  // double-render it).
  const lines = marked ? cfg.data.series.filter(isOverlay) : nSeries > 1 ? cfg.data.series.slice(-1) : [];
  const cols = marked
    ? cfg.data.series.filter((s) => !isOverlay(s))
    : nSeries > 1
      ? cfg.data.series.slice(0, nSeries - 1)
      : cfg.data.series;

  // Column mode: stacked (default), clustered, 100%, waterfall, or mekko.
  const columnsKind = cfg.combo?.columns ?? "stacked";
  // Independent axes: each line series gets its own scale (labelled, no shared
  // secondary-axis ticks) — for dashboards mixing unlike units.
  const independent = cfg.combo?.lineAxes === "independent" && lines.length >= 1;
  // Mekko (% normalised) and 100% columns expose no value→y map, so a line
  // needs its own axis; waterfall and column bases carry a shared scale.
  const noPrimaryAxis = columnsKind === "stacked100" || columnsKind === "mekko";
  // One shared scale: whichever of column extent / line values reaches higher.
  const stackMax =
    columnsKind === "clustered"
      ? maxOf(
          cols.flatMap((s) => s.values.filter((v): v is number => v != null)),
          0,
        )
      : maxOf(
          cfg.data.categories.map((_, c) => cols.reduce((a, s) => a + Math.max(0, s.values[c] ?? 0), 0)),
          0,
        );
  const lineMax = maxOf(
    lines.flatMap((s) => s.values.filter((v): v is number => v != null)),
    0,
  );
  // A shared-axis line can also dip BELOW the column base — e.g. a negative
  // overlay over all-positive or all-zero bars. The column scale floors at its
  // own data only, so without this the line plots far off the bottom of the plot
  // (a fuzz-found overshoot). Mirror the max-overflow fix: drop the shared floor
  // to the line, but only when it actually reaches lower than the columns would,
  // so combos whose bars already run more negative are left untouched.
  const lineMin = minOf(
    lines.flatMap((s) => s.values.filter((v): v is number => v != null)),
    0,
  );
  const colMin =
    columnsKind === "clustered"
      ? minOf(
          cols.flatMap((s) => s.values.filter((v): v is number => v != null)),
          0,
        )
      : minOf(
          cfg.data.categories.map((_, c) => cols.reduce((a, s) => a + Math.min(0, s.values[c] ?? 0), 0)),
          0,
        );
  const sharedFloor = lineMin < colMin ? niceTicks(lineMin, Math.max(stackMax, lineMax, 1))[0] : undefined;
  // Waterfall columns reach their running cumulative total, not the per-category
  // positive sum, so `stackMax` understates them; track the cumulative peak so a
  // shared-axis line taller than it isn't clipped off the top of the plot.
  const wf = (() => {
    const totals = new Set(cfg.waterfall?.totalIndices ?? []);
    const spacers = new Set(cfg.waterfall?.spacerIndices ?? []);
    const detailOf = detailParents(cfg);
    let running = 0;
    let max = 0;
    let min = 0;
    cfg.data.categories.forEach((_, c) => {
      // Match waterfallChain exactly: totals redraw the running level (no advance),
      // spacers carry it across unchanged, and detail-group columns decompose their
      // parent IN PLACE — none advance the chain. Counting spacer/detail columns
      // here overstated the peak, so the line-overflow stretch never fired and a
      // taller shared-axis line clipped off the top. Every column series
      // contributes to a real step (layoutWaterfall stacks them all).
      if (!totals.has(c) && !spacers.has(c) && !detailOf.has(c)) {
        running += cols.reduce((a, s) => a + (s.values[c] ?? 0), 0);
      }
      max = Math.max(max, running);
      min = Math.min(min, running);
    });
    return { max, min };
  })();
  // Secondary axis: line series get their own right-hand scale. A 100% / mekko
  // base forces it. Independent axes replace the shared secondary axis.
  const secondary = (!!cfg.secondaryAxis || noPrimaryAxis) && !independent;
  // A shared-axis line that overflows the waterfall's cumulative peak needs the
  // column scale stretched to fit it (only then — otherwise leave the
  // waterfall's own auto scale untouched to preserve existing layouts).
  const waterfallLineOverflow =
    columnsKind === "waterfall" && !secondary && !independent && cfg.scale?.max == null && lineMax > wf.max;
  // The mirror case #157 missed: a shared-axis line dipping BELOW the waterfall's
  // running trough (or below 0) needs the floor stretched down, or its low points
  // plot off the bottom of the plot.
  const waterfallLineUnderflow =
    columnsKind === "waterfall" && !secondary && !independent && cfg.scale?.min == null && lineMin < wf.min;
  // Overflow keeps its exact prior form (snapshot-neutral); underflow only adds a
  // min. When both fire, one niceTicks over the combined range keeps ticks aligned.
  const wfScale =
    waterfallLineOverflow && waterfallLineUnderflow
      ? (() => {
          const t = niceTicks(Math.min(lineMin, wf.min), Math.max(lineMax, wf.max, 1));
          return { ...cfg.scale, min: t[0], max: t[t.length - 1] };
        })()
      : waterfallLineOverflow
        ? { ...cfg.scale, max: niceTicks(0, Math.max(lineMax, wf.max, 1)).pop() }
        : waterfallLineUnderflow
          ? { ...cfg.scale, min: niceTicks(Math.min(lineMin, wf.min), Math.max(lineMax, wf.max, 1))[0] }
          : cfg.scale;
  const colCfg: ChartConfig = {
    ...cfg,
    kind: columnsKind,
    data: { ...cfg.data, series: cols },
    scale:
      columnsKind === "stacked100"
        ? undefined
        : columnsKind === "mekko"
          ? cfg.scale
          : columnsKind === "waterfall"
            ? wfScale
            : cfg.scale?.max != null || secondary
              ? cfg.scale
              : {
                  ...cfg.scale,
                  min: cfg.scale?.min ?? sharedFloor,
                  max: niceTicks(0, Math.max(stackMax, lineMax, 1)).pop(),
                },
  };
  const result =
    columnsKind === "waterfall"
      ? layoutWaterfall(colCfg, style, decor)
      : columnsKind === "mekko"
        ? layoutMekko(colCfg, style, decor)
        : columnsKind === "area"
          ? layoutLine(colCfg, style, decor)
          : layoutColumns(colCfg, style, decor);
  const { anchors, nodes } = result;
  // Horizontal (bar) base: the value axis runs left-to-right, and `categoryX`
  // holds each category's Y centre. The overlay is the same line drawn against
  // the other axis — before this it was drawn with the category's Y in the X
  // slot, so a horizontal Pareto (`pareto: true` keeps `horizontal`) rendered
  // its bars correctly and its cumulative-% line at ninety degrees to them.
  const H = !!cfg.horizontal;
  // No value→q map on the base and no line axis at all → nothing to overlay.
  // Read the map that matches the ORIENTATION: `valueToY` is undefined on every
  // horizontal base, which used to drop each overlay series silently, while the
  // chart's accessible description still listed it.
  const baseMap = H ? anchors.valueToX : anchors.valueToY;
  if (!baseMap && !secondary && !independent) return result;

  const fs = style.fontSize;
  /** Value → coordinate along the value axis (x when horizontal, y when not). */
  let lineToY = baseMap ?? ((v: number) => (H ? anchors.plot.x + v : anchors.plot.y + anchors.plot.h - v));
  /** Point for a value at a category, in the base chart's orientation. */
  const pointAt = (q: number, c: number) => (H ? { x: q, y: anchors.categoryX[c] } : { x: anchors.categoryX[c], y: q });
  if (secondary) {
    // Span down to any negative line value, not just up from 0 — anchoring the
    // secondary axis at 0 mapped a negative point below the plot floor. For an
    // all-positive line min(0, lineMin) is 0, so the domain and mapping are
    // unchanged (byte-identical); only a negative overlay is rescued.
    const ticks2 = niceTicks(Math.min(0, lineMin), Math.max(1, lineMax), 5);
    const min2 = ticks2[0];
    const max2 = ticks2[ticks2.length - 1];
    const plot = anchors.plot;
    lineToY = H
      ? (v: number) => plot.x + ((v - min2) / (max2 - min2 || 1)) * plot.w
      : (v: number) => plot.y + plot.h - ((v - min2) / (max2 - min2 || 1)) * plot.h;
    const fmt2 = resolveFormat(ticks2, cfg.numberFormat);
    for (const t of ticks2) {
      // The base chart's own axis sits below (horizontal) or left (vertical), so
      // the secondary strip goes on the opposite side: above the plot for bars,
      // right of it for columns. Drawn at plot right on a bar chart it ran off
      // the canvas edge. On a bar chart it is also pinned to the canvas top when
      // the chrome left no room — an untitled one plots from y≈0, and
      // plot.y − 1.5em put the whole strip off the canvas.
      const q = lineToY(t);
      nodes.push({
        kind: "text",
        x: H ? Math.max(0, Math.min(q - fs * 1.7, cfg.width - fs * 3.4)) : plot.x + plot.w + 2,
        y: H ? Math.max(0, plot.y - fs * 1.5) : q - fs * 0.7,
        w: fs * 3.4,
        h: fs * 1.4,
        text: formatNumber(t, fmt2),
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: H ? "center" : "left",
        valign: "middle",
        name: "secondary-axis",
      });
    }
  }
  const fmt = resolveFormat(
    lines.flatMap((s) => s.values.filter((v): v is number => v != null)),
    cfg.numberFormat,
  );
  lines.forEach((s, li) => {
    const color = seriesColor(style, cols.length + li, s.color);
    // Independent axis: zoom this line to its own value range (a nice-ticked
    // [min,max]) so its shape is visible whatever its units; the point labels
    // carry the real values since there is no shared numeric axis to read.
    const nums = s.values.filter((v): v is number => v != null);
    const ownTicks = independent && nums.length ? niceTicks(Math.min(...nums), Math.max(...nums)) : [0, 1];
    const lo = ownTicks[0];
    const hi = ownTicks[ownTicks.length - 1];
    const toY = independent
      ? H
        ? (v: number) => anchors.plot.x + ((v - lo) / (hi - lo || 1)) * anchors.plot.w
        : (v: number) => anchors.plot.y + anchors.plot.h - ((v - lo) / (hi - lo || 1)) * anchors.plot.h
      : lineToY;
    const labelOn = decor.segmentLabels || independent;
    // A marker series is this same overlay minus the connecting segments: the
    // values are per-category facts (a benchmark, a target, a consensus), and a
    // line between them would claim they interpolate. The mark is a little
    // larger, since it has no line to carry it.
    const markersOnly = s.type === "marker";
    let prev: { x: number; y: number } | null = null;
    let last: { x: number; y: number } | null = null;
    s.values.forEach((v, c) => {
      if (v == null || c >= anchors.categoryX.length) {
        prev = null;
        return;
      }
      const pt = pointAt(toY(v), c);
      if (prev && !markersOnly)
        nodes.push({
          kind: "line",
          x1: prev.x,
          y1: prev.y,
          x2: pt.x,
          y2: pt.y,
          stroke: color,
          strokeWidth: 2,
          name: `combo-line-${li}-${c}`,
        });
      const r = markersOnly ? 3.2 : 2.4;
      nodes.push({
        kind: "rect",
        x: pt.x - r,
        y: pt.y - r,
        w: r * 2,
        h: r * 2,
        fill: color,
        stroke: style.background,
        strokeWidth: 1,
        name: `combo-marker-${li}-${c}`,
      });
      if (labelOn) {
        // Categories run down a bar chart, so a label ABOVE its point would sit
        // on the neighbouring category's row; put it beside the mark instead.
        nodes.push(
          H
            ? {
                kind: "text",
                x: Math.min(pt.x + r + 2, cfg.width - 60),
                y: pt.y - fs * 0.7,
                w: 60,
                h: fs * 1.4,
                text: formatNumber(v, fmt),
                fontSize: fs,
                color: independent ? color : style.text,
                align: "left",
                valign: "middle",
                name: `combo-label-${li}-${c}`,
              }
            : {
                kind: "text",
                x: pt.x - 30,
                y: pt.y - fs * 1.65,
                w: 60,
                h: fs * 1.4,
                text: formatNumber(v, fmt),
                fontSize: fs,
                color: independent ? color : style.text,
                align: "center",
                valign: "bottom",
                name: `combo-label-${li}-${c}`,
              },
        );
      }
      prev = pt;
      last = pt;
    });
    if (decor.seriesLabels && last != null) {
      const end: { x: number; y: number } = last;
      nodes.push({
        kind: "text",
        x: H ? Math.max(0, Math.min(end.x - 40, cfg.width - 80)) : anchors.plot.x + anchors.plot.w + 4,
        y: H ? Math.max(0, end.y + fs * 0.9) : end.y - fs * 1.6,
        w: H ? 80 : cfg.width - (anchors.plot.x + anchors.plot.w) - 4,
        h: fs * 1.4,
        text: s.name,
        fontSize: fs,
        color: style.text,
        align: H ? "center" : "left",
        valign: "middle",
        name: `combo-series-label-${li}`,
      });
    }
  });
  return result;
}

/** Chrome for horizontal (bar) orientation: title, legend, left category labels, bottom axis. */
export function horizontalChrome(
  cfg: ChartConfig,
  style: ChartStyle,
  decor: Decorations,
  frame: Frame,
  centers: number[],
  scale: ValueScale,
  qOf: (v: number) => number,
): SceneNode[] {
  const nodes: SceneNode[] = chromeNodes(
    cfg,
    style,
    { ...decor, categoryAxis: false, valueAxis: false, gridlines: false },
    frame,
    centers,
  );
  const fs = style.fontSize;
  if (decor.gridlines) {
    for (const t of scale.ticks) {
      if (t === 0) continue;
      const x = frame.x + qOf(t);
      nodes.push({
        kind: "line",
        x1: x,
        y1: frame.y,
        x2: x,
        y2: frame.y + frame.h,
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: "gridline",
      });
    }
  }
  if (decor.valueAxis) {
    // Same labeller as the vertical chrome: a bar chart's value axis is the same
    // tick strip rotated, so it earns the same step-derived precision and the
    // same unitless share branch (see axisTickLabel).
    const axisLabel = axisTickLabel(scale.ticks, scale.percent, cfg.numberFormat);
    for (const t of scale.ticks) {
      const x = frame.x + qOf(t);
      nodes.push({
        kind: "text",
        x: x - 24,
        y: frame.y + frame.h + 2,
        w: 48,
        h: fs * 1.4,
        text: axisLabel(t),
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: "center",
        valign: "top",
        name: "value-axis",
      });
    }
  }
  if (decor.categoryAxis) {
    cfg.data.categories.forEach((cat, i) => {
      nodes.push({
        kind: "text",
        x: 0,
        y: centers[i] - fs * 0.75,
        w: frame.x - 4,
        h: fs * 1.5,
        text: cat,
        fontSize: fs,
        color: style.text,
        align: "right",
        valign: "middle",
        name: `category-${i}`,
      });
    });
  }
  if (decor.seriesLabels && cfg.data.series.length > 1) {
    nodes.push(...legendRow(cfg, style, frame.x, (cfg.title ? fs * 1.6 + 6 : 0) + 2, { maxX: cfg.width - 4 }));
  }
  return nodes;
}

/** Horizontal legend row: color chip + series name, left to right. */
/** One legend entry: a coloured chip and a label. */
export interface LegendEntry {
  label: string;
  /** Chip fill, as the mark it labels is painted ("none" for a hollow bar). */
  color: string;
  /** Chip outline / hatch, so the chip is a miniature of that mark. */
  stroke?: string;
  strokeWidth?: number;
  pattern?: Series["pattern"];
  /** Node name for the text (defaults to `legend-${index}`). */
  name?: string;
}

/**
 * Horizontal legend of coloured chips, wrapping to new rows so a chart with many
 * series/groups never marches its chips off the right edge (`opts.maxX`). Custom
 * entries (group names, a "Peer range" swatch) come via `opts.entries`; without
 * them it legends the multi-series set from `seriesLegendLabels`. Called with no
 * opts it is byte-identical to the old single-row version (maxX defaults to no
 * wrap).
 */
export function legendRow(
  cfg: ChartConfig,
  style: ChartStyle,
  x0: number,
  y: number,
  opts: { maxX?: number; entries?: LegendEntry[] } = {},
): SceneNode[] {
  const fs = style.fontSize;
  const nodes: SceneNode[] = [];
  const chip = fs * 0.7;
  const rowH = fs * 1.6;
  const maxX = opts.maxX ?? Infinity;
  const entries: LegendEntry[] =
    opts.entries ??
    seriesLegendLabels(cfg).map((label, si) => {
      // Key the chip with the paint the segments actually get, not the raw
      // series colour — pattern and IBCS scenario are half the encoding.
      const paint = markPaint(style, seriesColor(style, si, cfg.data.series[si].color), cfg.data.series[si]);
      return {
        label,
        color: paint.fill,
        stroke: paint.stroke,
        strokeWidth: paint.strokeWidth,
        pattern: paint.pattern,
        name: `legend-${si}`,
      };
    });
  // Shared wrap walk (frame.ts) so the row count here matches what the frame
  // reserved via legendRowCount.
  const slots = legendWrapWalk(
    entries.map((e) => e.label),
    fs,
    x0,
    maxX,
  );
  entries.forEach((e, si) => {
    const wLabel = textWidth(e.label, fs);
    const { x, row } = slots[si];
    const ry = y + row * rowH;
    nodes.push(
      {
        kind: "rect",
        x,
        y: ry + fs * 0.35,
        w: chip,
        h: chip,
        fill: e.color,
        stroke: e.stroke,
        strokeWidth: e.strokeWidth,
        pattern: e.pattern,
        name: `legend-chip-${si}`,
      },
      {
        kind: "text",
        x: x + chip + 3,
        y: ry,
        w: wLabel + 6,
        h: fs * 1.4,
        text: e.label,
        fontSize: fs,
        color: style.text,
        align: "left",
        valign: "middle",
        name: e.name ?? `legend-${si}`,
      },
    );
  });
  return nodes;
}

/**
 * Right-hand series labels at the last column's segment midpoints,
 * greedily pushed apart so they never overlap (think-cell placement).
 */
export function seriesLabelNodes(
  cfg: ChartConfig,
  style: ChartStyle,
  frame: { x: number; y: number; w: number; h: number },
  midYs: (number | null)[],
): SceneNode[] {
  const fs = style.fontSize;
  const lineH = fs * 1.35;
  /** Font the labels are actually drawn at — reduced only if they must be spread. */
  let labelFs = fs;
  const entries = cfg.data.series
    .map((s, i) => ({
      name: s.scenario ? `${s.name} (${s.scenario})` : s.name,
      color: seriesColor(style, i, s.color),
      y: midYs[i],
    }))
    .filter((e): e is { name: string; color: string; y: number } => e.y != null)
    .sort((a, b) => a.y - b.y);
  // Push overlapping labels apart, then clamp back into the frame.
  for (let i = 1; i < entries.length; i++) {
    if (entries[i].y - entries[i - 1].y < lineH) entries[i].y = entries[i - 1].y + lineH;
  }
  const overflow = entries.length ? entries[entries.length - 1].y + lineH / 2 - (frame.y + frame.h) : 0;
  if (overflow > 0) {
    for (const e of entries) e.y -= overflow;
    for (let i = entries.length - 2; i >= 0; i--) {
      if (entries[i + 1].y - entries[i].y < lineH) entries[i].y = entries[i + 1].y - lineH;
    }
    // Only the BOTTOM was clamped, so the upward propagation walked straight
    // past the canvas top: twelve series on a 240×160 chart put two labels at
    // negative y, and thirty on a 400×300 put the topmost at −123. `collide.ts`
    // refuses to nudge a label off the top for the reason that applies here too
    // — an overlapping label still reads, an off-canvas one is lost — but it
    // only moves labels UP, so it cannot rescue one already emitted above the
    // canvas. Spread evenly over the band when the gap cannot be honoured,
    // which is what `layoutSlope.place()` does after the same discovery.
    const top = lineH / 2;
    if (entries[0].y < top) {
      const bottom = Math.max(top, Math.min(frame.y + frame.h, cfg.height - lineH / 2));
      const step = entries.length > 1 ? (bottom - top) / (entries.length - 1) : 0;
      entries.forEach((e, i) => (e.y = top + i * step));
      // Spread AND shrunk, because a spread alone trades one defect for a worse
      // one. The step here is by definition below the gap the labels wanted, so
      // every neighbouring pair overlaps — and `series-label-` is in
      // `collide.ts`'s MOVABLE list, whose nudge only goes UP. It therefore
      // pushed each label past the one above it and returned them REORDERED:
      // every label on the canvas, each naming the wrong line. An unreadably
      // small label is a bad chart; a legible label naming someone else's line
      // is a wrong one. Shrink to the step so nothing overlaps and the
      // de-collision pass has nothing to do.
      labelFs = Math.max(5, Math.min(fs, step / 1.25));
    }
  }
  const x = frame.x + frame.w + 4;
  return entries.map((e, i) => ({
    kind: "text" as const,
    x,
    y: e.y - lineH / 2,
    w: cfg.width - x,
    h: lineH,
    text: e.name,
    fontSize: labelFs,
    color: style.text,
    align: "left" as const,
    valign: "middle" as const,
    name: `series-label-${i}`,
  }));
}
