import type { ChartConfig, ChartStyle, Decorations, LayoutAnchors } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, niceTicks, resolveFormat, segmentLabel } from "../format";
import { seriesColor } from "../style";
import {
  baselineNode,
  breakMarkerNodes,
  chromeNodes,
  computeFrame,
  computeFrameHorizontal,
  valueScale,
  type Frame,
  type ValueScale,
} from "./frame";
// Combo base kinds. These modules import back from column (LayoutResult /
// horizontalChrome), but the calls happen at runtime so the ESM cycle resolves.
import { layoutWaterfall } from "./waterfall";
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

  const frame = H
    ? computeFrameHorizontal(cfg, style, decor)
    : computeFrame(cfg, style, decor, decor.seriesLabels ? data.series.map((s) => s.name) : []).frame;
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

  let dataMin: number, dataMax: number;
  if (pct) {
    dataMin = 0;
    dataMax = 1;
  } else if (stacked && nStacks > 1) {
    dataMin = Math.min(0, ...data.categories.flatMap((_, c) => stackIds.map((id) => stackNegTotals(c, id))));
    dataMax = Math.max(0, ...data.categories.flatMap((_, c) => stackIds.map((id) => stackPosTotals(c, id))));
  } else if (stacked) {
    dataMin = Math.min(0, ...negTotals);
    dataMax = Math.max(0, ...posTotals);
  } else {
    const all = data.series.flatMap((s) => s.values.filter((v): v is number => v != null));
    dataMin = minOf(all, 0);
    dataMax = maxOf(all, 0);
  }
  // 100% charts: negatives are shares below the zero line. The axis drops to
  // the most-negative column share (0 when all data is positive → unchanged).
  const pctNegMin = pct
    ? Math.min(0, ...data.categories.map((_, c) => (denominators[c] > 0 ? negTotals[c] / denominators[c] : 0)))
    : 0;
  const scale: ValueScale = pct
    ? {
        min: pctNegMin,
        max: 1,
        ticks: pctNegMin < 0 ? niceTicks(pctNegMin, 1, 5) : [0, 0.25, 0.5, 0.75, 1],
        toY: (v: number) => frame.y + frame.h - ((v - pctNegMin) / (1 - pctNegMin)) * frame.h,
      }
    : valueScale(frame, dataMin, dataMax, cfg.scale, H ? undefined : cfg.axisBreak, !stacked && !H && cfg.logScale);

  // Value coordinate: distance along the value axis from the scale minimum.
  // Vertical charts route through toY so axis breaks apply; horizontal stays linear.
  const valLen = H ? frame.w : frame.h;
  const qOf = H
    ? (v: number) => ((v - scale.min) / (scale.max - scale.min || 1)) * valLen
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
      const barStyle = !stacked && !H ? (decor.barStyle ?? "bar") : "bar";
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
      levels[si] = ups[sp] + downs[sp];
      if (!r) return;

      if (barStyle !== "bar") {
        // Lollipop / dot / dumbbell-range rendering for clustered charts:
        // the value point is a dot; lollipops add a stem from the baseline;
        // range connects the two series' dots with a line (drawn once).
        const dotX = barStyle === "range" ? centers[c] : r.x + r.w / 2;
        const dotY = v >= 0 ? r.y : r.y + r.h;
        const dotR = 4;
        if (barStyle === "lollipop") {
          nodes.push({
            kind: "line",
            x1: dotX,
            y1: frame.y + frame.h - qOf(0),
            x2: dotX,
            y2: dotY,
            stroke: fill,
            strokeWidth: 1.5,
            name: `stem-${si}-${c}`,
          });
        }
        if (barStyle === "range" && si === 1 && data.series[0].values[c] != null) {
          const y0r = frame.y + frame.h - qOf(data.series[0].values[c]!);
          nodes.push({
            kind: "line",
            x1: dotX,
            y1: y0r,
            x2: dotX,
            y2: dotY,
            stroke: style.mutedText,
            strokeWidth: 1.5,
            name: `range-${c}`,
          });
        }
        nodes.push({
          kind: "ellipse",
          cx: dotX,
          cy: dotY,
          rx: dotR,
          ry: dotR,
          fill,
          stroke: style.background,
          strokeWidth: 1,
          name: `seg-${si}-${c}`,
        });
        if (c === n - 1) lastSegMid[si] = dotY;
        if (decor.segmentLabels) {
          const label = formatNumber(raw!, fmt);
          nodes.push({
            kind: "text",
            x: dotX + dotR + 2,
            y: dotY - fs * 0.7,
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
      nodes.push({
        kind: "rect",
        ...r,
        fill,
        stroke: style.background,
        strokeWidth: stacked ? 0.75 : 0,
        pattern: s.pattern,
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
            color: contrastInk(fill),
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
        : Math.max(0, ...data.series.map((s) => s.values[c] ?? 0));
    const topQ = qOf(Math.max(0, topV));
    columnTop.push(H ? frame.x + topQ : frame.y + frame.h - topQ);

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
      columnValue: pct ? posTotals : signedTotals,
      seriesLevels,
      baselineY: y0,
      plot: { x: frame.x, y: frame.y, w: frame.w, h: frame.h },
      valueToY: pct || H ? undefined : scale.toY,
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
  const wfMax = (() => {
    const totals = new Set(cfg.waterfall?.totalIndices ?? []);
    let running = 0;
    let max = 0;
    cfg.data.categories.forEach((_, c) => {
      // Every column series contributes to the running total (layoutWaterfall
      // stacks them all) — summing only cols[0] understated the peak and pushed
      // a multi-series waterfall combo off the top of the plot.
      if (!totals.has(c)) running += cols.reduce((a, s) => a + (s.values[c] ?? 0), 0);
      max = Math.max(max, running);
    });
    return max;
  })();
  // Secondary axis: line series get their own right-hand scale. A 100% / mekko
  // base forces it. Independent axes replace the shared secondary axis.
  const secondary = (!!cfg.secondaryAxis || noPrimaryAxis) && !independent;
  // A shared-axis line that overflows the waterfall's cumulative peak needs the
  // column scale stretched to fit it (only then — otherwise leave the
  // waterfall's own auto scale untouched to preserve existing layouts).
  const waterfallLineOverflow =
    columnsKind === "waterfall" && !secondary && !independent && cfg.scale?.max == null && lineMax > wfMax;
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
            ? waterfallLineOverflow
              ? { ...cfg.scale, max: niceTicks(0, Math.max(lineMax, wfMax, 1)).pop() }
              : cfg.scale
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
  // No value→y map and no line axis at all → nothing to overlay.
  if (!anchors.valueToY && !secondary && !independent) return result;

  const fs = style.fontSize;
  let lineToY = anchors.valueToY ?? ((v: number) => anchors.plot.y + anchors.plot.h - v);
  if (secondary) {
    const ticks2 = niceTicks(0, Math.max(1, lineMax), 5);
    const max2 = ticks2[ticks2.length - 1];
    const plot = anchors.plot;
    lineToY = (v: number) => plot.y + plot.h - (v / max2) * plot.h;
    const fmt2 = resolveFormat(ticks2, cfg.numberFormat);
    for (const t of ticks2) {
      nodes.push({
        kind: "text",
        x: plot.x + plot.w + 2,
        y: lineToY(t) - fs * 0.7,
        w: fs * 3.4,
        h: fs * 1.4,
        text: formatNumber(t, fmt2),
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: "left",
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
      ? (v: number) => anchors.plot.y + anchors.plot.h - ((v - lo) / (hi - lo || 1)) * anchors.plot.h
      : lineToY;
    const labelOn = decor.segmentLabels || independent;
    // A marker series is this same overlay minus the connecting segments: the
    // values are per-category facts (a benchmark, a target, a consensus), and a
    // line between them would claim they interpolate. The mark is a little
    // larger, since it has no line to carry it.
    const markersOnly = s.type === "marker";
    let prev: { x: number; y: number } | null = null;
    let lastY: number | null = null;
    s.values.forEach((v, c) => {
      if (v == null || c >= anchors.categoryX.length) {
        prev = null;
        return;
      }
      const pt = { x: anchors.categoryX[c], y: toY(v) };
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
        nodes.push({
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
        });
      }
      prev = pt;
      lastY = pt.y;
    });
    if (decor.seriesLabels && lastY != null) {
      nodes.push({
        kind: "text",
        x: anchors.plot.x + anchors.plot.w + 4,
        y: lastY - fs * 1.6,
        w: cfg.width - (anchors.plot.x + anchors.plot.w) - 4,
        h: fs * 1.4,
        text: s.name,
        fontSize: fs,
        color: style.text,
        align: "left",
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
    const axisFmt = resolveFormat(scale.ticks, cfg.numberFormat);
    for (const t of scale.ticks) {
      const x = frame.x + qOf(t);
      nodes.push({
        kind: "text",
        x: x - 24,
        y: frame.y + frame.h + 2,
        w: 48,
        h: fs * 1.4,
        text: formatNumber(t, axisFmt),
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
  color: string;
  /** Node name for the text (defaults to `legend-${index}`). */
  name?: string;
}

/**
 * Horizontal legend of coloured chips, wrapping to new rows so a chart with many
 * series/groups never marches its chips off the right edge (`opts.maxX`). Custom
 * entries (group names, a "Peer range" swatch) come via `opts.entries`; without
 * them it legends `cfg.data.series`. Called with no opts it is byte-identical to
 * the old single-row version (maxX defaults to no wrap).
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
    cfg.data.series.map((s, si) => ({ label: s.name, color: seriesColor(style, si, s.color), name: `legend-${si}` }));
  let x = x0;
  let row = 0;
  entries.forEach((e, si) => {
    const wLabel = textWidth(e.label, fs);
    // Wrap before an entry that would cross maxX — never on the first of a row.
    if (x > x0 && x + chip + 3 + wLabel > maxX) {
      x = x0;
      row++;
    }
    const ry = y + row * rowH;
    nodes.push(
      { kind: "rect", x, y: ry + fs * 0.35, w: chip, h: chip, fill: e.color, name: `legend-chip-${si}` },
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
    x += chip + 3 + wLabel + 12;
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
  const entries = cfg.data.series
    .map((s, i) => ({ name: s.name, color: seriesColor(style, i, s.color), y: midYs[i] }))
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
  }
  const x = frame.x + frame.w + 4;
  return entries.map((e, i) => ({
    kind: "text" as const,
    x,
    y: e.y - lineH / 2,
    w: cfg.width - x,
    h: lineH,
    text: e.name,
    fontSize: fs,
    color: style.text,
    align: "left" as const,
    valign: "middle" as const,
    name: `series-label-${i}`,
  }));
}
