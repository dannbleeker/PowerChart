import type { ChartConfig, ChartStyle, Decorations, MarkerSymbol } from "../types";
import { markerScale } from "../geometry";
import { textWidth, type SceneNode } from "../scene";
import { formatNumber, formatP, histogramBins, niceTicks, resolveFormat, trendStats } from "../format";
import { placeLabels, type Box, type LabelRequest } from "../labels";
import { spreadAlongAxis } from "../spread";
import { PALETTE } from "../style";
import { lerpColor, sequentialScale } from "../color";
import { footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Scatter / bubble chart, think-cell datasheet convention: categories are
 * point labels; rows named X, Y (and Size for bubbles) carry coordinates.
 * A row named Group (values 1..k) colors points by group.
 * Point labels are placed by the greedy collision-avoiding placer and hidden
 * when the chart gets too dense.
 */
/**
 * The overlap-relief cap, in DATA UNITS of the spread axis — or null when
 * spread is off. Exported so the footnote quotes exactly the number the layout
 * enforces: a disclosure that drifts from the cap is worse than no disclosure.
 */
export function spreadCap(cfg: ChartConfig): { axis: "x" | "y"; limit: number } | null {
  const axis = cfg.scatter?.spread;
  if (!axis || (cfg.kind !== "scatter" && cfg.kind !== "bubble")) return null;
  // Quadrants make a categorical claim — which box a point is in. A nudge, however
  // small and however well disclosed, could carry a marker across the crossing
  // line and change that claim. Overlap is the lesser problem, so spread yields.
  // Returning null here also suppresses the footnote, so the chart never
  // promises an approximation it did not make.
  if (cfg.decorations?.quadrants) return null;
  // Measure the cap against the SAME nice-ticked domain the plot maps, so the
  // disclosed number matches what the viewer sees. (This used to force a zero
  // baseline of its own, which — once the plot became data-driven — let the cap
  // exceed the whole visible range on a tight cluster.)
  const ticks = niceTicks(...scatterDomain(cfg, axis), 5);
  const range = ticks[ticks.length - 1] - ticks[0];
  if (!(range > 0)) return null;
  const limit = cfg.scatter?.spreadLimit ?? range * 0.02;
  return { axis, limit: Math.max(0, Math.min(limit, range * 0.1)) };
}

/**
 * The value domain the plot maps for one axis of a scatter/bubble: the plotted
 * points' coordinates plus the decorations that must stay on-plot (partition
 * lines, a quadrant crossing, axis bands). Shared by the layout and spreadCap.
 *
 * The domain follows the DATA and is NOT anchored to zero — a scatter is read by
 * point position, not by length-from-zero the way a bar chart is, so forcing a
 * zero baseline (as this once did) collapsed a tight cluster like x∈[1000,1050]
 * into a sliver.
 */
export function scatterDomain(cfg: ChartConfig, axis: "x" | "y"): [number, number] {
  const find = (re: RegExp) => cfg.data.series.find((s) => re.test(s.name.trim()));
  const xs = find(/^x$/i)?.values ?? [];
  const ys = find(/^y$/i)?.values ?? [];
  // Only points carrying BOTH coordinates are plotted; a half-specified point
  // contributes no extent.
  const vals: number[] = [];
  for (let i = 0; i < cfg.data.categories.length; i++) {
    const x = xs[i];
    const y = ys[i];
    if (x == null || y == null) continue;
    vals.push(axis === "x" ? x : y);
  }
  const extra = (find(axis === "x" ? /^x\s*line$/i : /^y\s*line$/i)?.values ?? []).filter(
    (v): v is number => v != null,
  );
  const q = cfg.decorations?.quadrants;
  if (q) extra.push(axis === "x" ? q.x : q.y);
  for (const b of cfg.decorations?.bands ?? []) if (b.axis === axis) extra.push(b.from, b.to);
  const all = [...vals, ...extra].filter((v) => Number.isFinite(v));
  if (!all.length) return [0, 1];
  let lo = Math.min(...all);
  let hi = Math.max(...all);
  if (lo === hi) {
    const pad = Math.abs(lo) > 1e-9 ? Math.abs(lo) * 0.1 : 1;
    lo -= pad;
    hi += pad;
  }
  return [lo, hi];
}

/** Width of the continuous-color gradient bar; the group legend clears it. */
const COLOR_BAR_W = 90;

/**
 * A point's group id, normalised to a whole number ≥ 1.
 *
 * Group ids come out of a datasheet row, so they are whatever a cell can hold —
 * blank, 0, negative, fractional, NaN. Every consumer indexes something with
 * them (the palette, the marker shape, the legend), and each used to re-derive
 * its own index from the raw value: `palette[NaN]` quietly yields `fill=
 * "undefined"`, and `markers[NaN]` yields no shape at all, which throws in the
 * renderer. Normalising once, here, makes all of them safe by construction.
 * Anything not a usable id reads as group 1 — the same bucket a point with no
 * Group row already falls into.
 */
function groupOf(v: number | null | undefined): number {
  const g = Math.round(Number(v ?? 1));
  return Number.isFinite(g) && g >= 1 ? g : 1;
}

/**
 * A point mark of the given shape, carrying the same ink as a circle of radius
 * `r`.
 *
 * `r` is the DATA radius — "area ∝ size" is the bubble's central claim — so
 * every shape is grown or shrunk by markerScale to match the circle's area.
 * Without that, a group drawn as stars would read as a quarter the magnitude of
 * an identical value drawn as squares.
 *
 * "circle" and "square" resolve to the ellipse and rect the scene already has,
 * so the default scatter emits exactly the node it always did (markerScale is
 * 1 for a circle) and its output cannot move; only the shapes that need preset
 * geometry become SymbolNodes.
 */
function markerNode(
  shape: MarkerSymbol,
  cx: number,
  cy: number,
  r: number,
  fill: string,
  stroke: string,
  strokeWidth: number,
  name: string,
): SceneNode {
  const m = markerExtent(shape, r);
  if (shape === "circle") return { kind: "ellipse", cx, cy, rx: m, ry: m, fill, stroke, strokeWidth, name };
  if (shape === "square")
    return { kind: "rect", x: cx - m, y: cy - m, w: m * 2, h: m * 2, fill, stroke, strokeWidth, name };
  return { kind: "symbol", shape, cx, cy, size: m, fill, stroke, strokeWidth, name };
}

/** Drawn half-extent of a marker whose data radius is `r`. */
const markerExtent = (shape: MarkerSymbol, r: number) => r * markerScale(shape);

export function layoutScatter(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const find = (re: RegExp) => data.series.find((s) => re.test(s.name.trim()));
  const xs = find(/^x$/i)?.values ?? [];
  const ys = find(/^y$/i)?.values ?? [];
  const sizes = cfg.kind === "bubble" ? (find(/^size$/i)?.values ?? []) : [];
  const groups = find(/^group$/i)?.values ?? [];
  // A numeric "Color" row encodes a third/fourth variable on a sequential ramp.
  const colorVals = find(/^colou?r$/i)?.values ?? [];
  // Partition lines and trend line, think-cell's scatter decorations.
  const xLines = (find(/^x\s*line$/i)?.values ?? []).filter((v): v is number => v != null);
  const yLines = (find(/^y\s*line$/i)?.values ?? []).filter((v): v is number => v != null);
  const wantTrend = (find(/^trend$/i)?.values ?? []).some((v) => v != null);

  const pts = data.categories
    .map((label, i) => ({
      label,
      x: xs[i],
      y: ys[i],
      size: sizes[i] ?? null,
      group: groupOf(groups[i]),
      color: colorVals[i] ?? null,
    }))
    .filter((p): p is typeof p & { x: number; y: number } => p.x != null && p.y != null);

  // Shape per group, cycled like the palette. Off => every point a circle,
  // which is the ellipse the layout has always emitted.
  const markers = cfg.scatter?.markers?.length ? cfg.scatter.markers : null;
  const markerFor = (group: number): MarkerSymbol => (markers ? markers[(group - 1) % markers.length] : "circle");

  // Continuous color scale (a "Color" row): maps each point onto a sequential
  // ramp; supersedes group coloring and swaps the chip legend for a gradient.
  const colorNums = pts.map((p) => p.color).filter((v): v is number => v != null);
  const colorScale =
    colorNums.length > 0
      ? {
          min: Math.min(...colorNums),
          max: Math.max(...colorNums),
          of: sequentialScale(Math.min(...colorNums), Math.max(...colorNums), (cfg.style?.palette ?? PALETTE)[0]),
        }
      : null;

  const titleH = titleHeight(cfg, style);
  const axisW = 34;
  const multiGroup = !colorScale && new Set(pts.map((p) => p.group)).size > 1;
  const legendH = multiGroup || colorScale ? fs * 1.8 : 0;
  /** Where the plot starts before any marginal gutter — the legends' anchor. */
  const chromeTop = titleH + 6 + legendH;
  // Marginal gutters, in font-size units like the heatmap's already-shipped
  // marginal totals, so they scale with the style instead of being a magic 34.
  const GUT = fs * 3.4;
  const wantMx = decor.marginals === "x" || decor.marginals === "both";
  const wantMy = decor.marginals === "y" || decor.marginals === "both";
  // NB: bodyW/bodyH are today's w/h expressions character-for-character. Do not
  // "simplify" bodyH in terms of chromeTop — float subtraction is not
  // associative, and (H-t)-6-l differs from H-(t+6+l) for many font sizes. It
  // happens to agree at the default fs=10, which every showcase config uses, so
  // the deck's byte-identity gate would NOT catch the difference.
  const bodyW = cfg.width - axisW - 8;
  const bodyH = cfg.height - titleH - 6 - legendH - fs * 1.6 - footnoteH(cfg, style, decor);
  // A gutter costs real space. If what's left would stop being a chart, drop
  // the marginals rather than the plot.
  const MIN_PLOT = 60;
  const mTop = wantMx && bodyH - GUT >= MIN_PLOT ? GUT : 0;
  const mRight = wantMy && bodyW - GUT >= MIN_PLOT ? GUT : 0;
  const plot = {
    x: axisW,
    y: chromeTop + mTop,
    w: bodyW - mRight,
    h: bodyH - mTop,
  };

  // Data-driven axis domain (no forced zero baseline) shared with spreadCap;
  // folds in partition lines, a quadrant crossing and x/y bands so those
  // decorations never fall outside the plot. See scatterDomain.
  const xTicks = niceTicks(...scatterDomain(cfg, "x"), 5);
  const yTicks = niceTicks(...scatterDomain(cfg, "y"), 5);
  const x0 = xTicks[0];
  const x1 = xTicks[xTicks.length - 1];
  const y0 = yTicks[0];
  const y1 = yTicks[yTicks.length - 1];
  const toX = (v: number) => plot.x + ((v - x0) / (x1 - x0 || 1)) * plot.w;
  const toY = (v: number) => plot.y + plot.h - ((v - y0) / (y1 - y0 || 1)) * plot.h;

  const xFmt = resolveFormat(xTicks, cfg.numberFormat);
  const yFmt = resolveFormat(yTicks, cfg.numberFormat);

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);
  // Quadrant preset: one X/Y crossing → four tinted zones with corner
  // labels and the crossing lines — BCG-matrix framing in one step.
  if (decor.quadrants) {
    const { x: qx, y: qy, labels } = decor.quadrants;
    const cx = Math.max(plot.x, Math.min(plot.x + plot.w, toX(qx)));
    const cy = Math.max(plot.y, Math.min(plot.y + plot.h, toY(qy)));
    const zones: { x: number; y: number; w: number; h: number }[] = [
      { x: plot.x, y: plot.y, w: cx - plot.x, h: cy - plot.y }, // TL
      { x: cx, y: plot.y, w: plot.x + plot.w - cx, h: cy - plot.y }, // TR
      { x: plot.x, y: cy, w: cx - plot.x, h: plot.y + plot.h - cy }, // BL
      { x: cx, y: cy, w: plot.x + plot.w - cx, h: plot.y + plot.h - cy }, // BR
    ];
    zones.forEach((z, i) => {
      if (z.w <= 0 || z.h <= 0) return;
      // Checkerboard tint so adjacent zones read as distinct regions.
      nodes.push({ kind: "rect", ...z, fill: i === 0 || i === 3 ? "#f2f1ec" : "#faf9f6", name: `quadrant-${i}` });
      const label = labels?.[i];
      if (label) {
        nodes.push({
          kind: "text",
          x: z.x + 4,
          y: z.y + 2,
          w: Math.max(20, z.w - 8),
          h: fs * 1.3,
          text: label,
          fontSize: fs * 0.9,
          bold: true,
          color: style.mutedText,
          align: i === 1 || i === 3 ? "right" : "left",
          valign: "top",
          name: `quadrant-label-${i}`,
        });
      }
    });
    nodes.push(
      {
        kind: "line",
        x1: cx,
        y1: plot.y,
        x2: cx,
        y2: plot.y + plot.h,
        stroke: style.mutedText,
        strokeWidth: 1,
        dash: [3, 2],
        name: "quadrant-x",
      },
      {
        kind: "line",
        x1: plot.x,
        y1: cy,
        x2: plot.x + plot.w,
        y2: cy,
        stroke: style.mutedText,
        strokeWidth: 1,
        dash: [3, 2],
        name: "quadrant-y",
      },
    );
  }

  // Background bands (both axes in value units), behind gridlines and points.
  decor.bands?.forEach((band, i) => {
    const clampX = (v: number) => Math.max(plot.x, Math.min(plot.x + plot.w, toX(v)));
    const clampY = (v: number) => Math.max(plot.y, Math.min(plot.y + plot.h, toY(v)));
    const r =
      band.axis === "x"
        ? {
            x: Math.min(clampX(band.from), clampX(band.to)),
            y: plot.y,
            w: Math.abs(clampX(band.to) - clampX(band.from)),
            h: plot.h,
          }
        : {
            x: plot.x,
            y: Math.min(clampY(band.from), clampY(band.to)),
            w: plot.w,
            h: Math.abs(clampY(band.to) - clampY(band.from)),
          };
    if (r.w <= 0 || r.h <= 0) return;
    nodes.push({ kind: "rect", ...r, fill: band.color ?? "#f2f1ec", name: `band-${i}` });
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

  // Gridlines + axis labels on both axes.
  for (const t of yTicks) {
    const y = toY(t);
    nodes.push(
      {
        kind: "line",
        x1: plot.x,
        y1: y,
        x2: plot.x + plot.w,
        y2: y,
        stroke: style.gridline,
        strokeWidth: 0.75,
        name: "gridline-y",
      },
      {
        kind: "text",
        x: 0,
        y: y - fs * 0.7,
        w: plot.x - 4,
        h: fs * 1.4,
        text: formatNumber(t, yFmt),
        fontSize: fs * 0.9,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: "y-axis",
      },
    );
  }
  for (const t of xTicks) {
    const x = toX(t);
    nodes.push({
      kind: "text",
      x: x - 24,
      y: plot.y + plot.h + 2,
      w: 48,
      h: fs * 1.4,
      text: formatNumber(t, xFmt),
      fontSize: fs * 0.9,
      color: style.mutedText,
      align: "center",
      valign: "top",
      name: "x-axis",
    });
  }
  nodes.push(
    {
      kind: "line",
      x1: plot.x,
      y1: plot.y + plot.h,
      x2: plot.x + plot.w,
      y2: plot.y + plot.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "baseline",
    },
    {
      kind: "line",
      x1: toX(0) >= plot.x ? toX(0) : plot.x,
      y1: plot.y,
      x2: toX(0) >= plot.x ? toX(0) : plot.x,
      y2: plot.y + plot.h,
      stroke: style.axis,
      strokeWidth: 1,
      name: "y-axis-line",
    },
  );

  // Partition lines (dashed) at fixed x / y values.
  for (const v of xLines) {
    const x = toX(v);
    nodes.push({
      kind: "line",
      x1: x,
      y1: plot.y,
      x2: x,
      y2: plot.y + plot.h,
      stroke: style.mutedText,
      strokeWidth: 1,
      dash: [3, 2],
      name: "x-line",
    });
  }
  for (const v of yLines) {
    const y = toY(v);
    nodes.push({
      kind: "line",
      x1: plot.x,
      y1: y,
      x2: plot.x + plot.w,
      y2: y,
      stroke: style.mutedText,
      strokeWidth: 1,
      dash: [3, 2],
      name: "y-line",
    });
  }

  // OLS trend line across all points — always stating fit and significance.
  if (wantTrend && pts.length >= 2) {
    const mx = pts.reduce((s, p) => s + p.x, 0) / pts.length;
    const my = pts.reduce((s, p) => s + p.y, 0) / pts.length;
    const sxx = pts.reduce((s, p) => s + (p.x - mx) ** 2, 0);
    if (sxx > 0) {
      const slope = pts.reduce((s, p) => s + (p.x - mx) * (p.y - my), 0) / sxx;
      const at = (x: number) => my + slope * (x - mx);
      nodes.push({
        kind: "line",
        x1: toX(x0),
        y1: toY(at(x0)),
        x2: toX(x1),
        y2: toY(at(x1)),
        stroke: style.negative,
        strokeWidth: 1.25,
        dash: [4, 2],
        name: "trend",
      });
      const stats = trendStats(pts);
      if (stats) {
        const label = `R² = ${stats.r2.toFixed(2)}${stats.p != null ? `, p ${formatP(stats.p)}` : ""}`;
        const endY = toY(at(x1));
        nodes.push({
          kind: "text",
          x: plot.x + plot.w - textWidth(label, fs * 0.9) - 4,
          y: Math.max(plot.y, Math.min(plot.y + plot.h - fs * 1.3, endY + (slope >= 0 ? fs * 0.4 : -fs * 1.7))),
          w: textWidth(label, fs * 0.9) + 4,
          h: fs * 1.3,
          text: label,
          fontSize: fs * 0.9,
          color: style.negative,
          align: "right",
          valign: "middle",
          name: "trend-stats",
        });
      }
    }
  }

  // Continuous color legend: a discretized gradient bar with min/max labels
  // (renderer-safe — small rects, no SVG gradient).
  if (colorScale) {
    const cFmt = resolveFormat([colorScale.min, colorScale.max], cfg.numberFormat);
    const steps = 24;
    const barW = COLOR_BAR_W;
    const cell = barW / steps;
    const bx = plot.x;
    // The min/max labels hang BELOW the gradient bar, so with a top gutter the
    // legend has to sit a little higher or they land on the marginal bars.
    // Only when the gutter exists, so no existing output moves.
    const by = chromeTop - fs * (mTop > 0 ? 1.75 : 1.35);
    for (let i = 0; i < steps; i++) {
      const t = (i + 0.5) / steps;
      nodes.push({
        kind: "rect",
        x: bx + i * cell,
        y: by,
        w: cell + 0.5,
        h: fs * 0.7,
        fill: colorScale.of(colorScale.min + t * (colorScale.max - colorScale.min)),
        name: `color-legend-${i}`,
      });
    }
    const colorName = find(/^colou?r$/i)?.name ?? "Color";
    nodes.push(
      {
        kind: "text",
        x: bx - 40,
        y: by - fs * 0.15,
        w: 38,
        h: fs,
        text: colorName,
        fontSize: fs * 0.85,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: "color-legend-title",
      },
      {
        kind: "text",
        x: bx,
        y: by + fs * 0.75,
        w: barW,
        h: fs,
        text: formatNumber(colorScale.min, cFmt),
        fontSize: fs * 0.75,
        color: style.mutedText,
        align: "left",
        valign: "middle",
        name: "color-legend-min",
      },
      {
        kind: "text",
        x: bx,
        y: by + fs * 0.75,
        w: barW,
        h: fs,
        text: formatNumber(colorScale.max, cFmt),
        fontSize: fs * 0.75,
        color: style.mutedText,
        align: "right",
        valign: "middle",
        name: "color-legend-max",
      },
    );
  }

  // Group legend when points are colored by group. A color scale normally
  // supersedes group coloring and suppresses this — but markers put group on
  // the SHAPE channel, which the color legend says nothing about, so the legend
  // has to come back or the shapes stand unexplained.
  const groupIds = colorScale && !markers ? [] : [...new Set(pts.map((p) => p.group))].sort((a, b) => a - b);
  if (groupIds.length > 1) {
    // Clear the gradient bar when both legends are up — they share this row and
    // both anchor at plot.x, so without the offset the chips land on the ramp.
    let lx = plot.x + (colorScale ? COLOR_BAR_W + 16 : 0);
    for (const g of groupIds) {
      const chip = fs * 0.7;
      const label = `Group ${g}`;
      // Under a color scale the chip's color would be a lie (color means the
      // Color row there), so the shape carries the legend in neutral ink.
      const chipFill = colorScale ? style.mutedText : (cfg.style?.palette ?? PALETTE)[(g - 1) % 8];
      // The chip is drawn like the points it explains, area and all — so an
      // area-matched star is 1.67x wider than `chip` and would sit on its own
      // label. Advance by what was actually drawn. Without markers the drawn
      // width IS chip, so the default legend's spacing is untouched.
      const drawn = markers ? markerExtent(markerFor(g), chip / 2) * 2 : chip;
      nodes.push(
        markers
          ? markerNode(
              markerFor(g),
              lx + drawn / 2,
              chromeTop - fs * 1.2 + chip / 2,
              chip / 2,
              chipFill,
              style.background,
              0,
              `legend-chip-${g}`,
            )
          : {
              kind: "rect",
              x: lx,
              y: chromeTop - fs * 1.2,
              w: chip,
              h: chip,
              fill: chipFill,
              name: `legend-chip-${g}`,
            },
        {
          kind: "text",
          x: lx + drawn + 3,
          y: chromeTop - fs * 1.55,
          w: textWidth(label, fs) + 6,
          h: fs * 1.4,
          text: label,
          fontSize: fs,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `legend-${g}`,
        },
      );
      lx += drawn + 3 + textWidth(label, fs) + 14;
    }
  }

  // Bubble radius: area ∝ size, max radius 9% of the smaller plot dimension.
  const maxSize = Math.max(1e-9, ...pts.map((p) => Math.abs(p.size ?? 0)));
  const maxR = Math.min(plot.w, plot.h) * 0.09;
  const radius = (p: (typeof pts)[number]) =>
    cfg.kind === "bubble" && p.size != null ? Math.max(2.5, Math.sqrt(Math.abs(p.size) / maxSize) * maxR) : 3;

  // Overlap relief along one axis. The cross axis stays exact, the named one
  // moves by at most the disclosed cap, and every marker keeps its identity —
  // this only shifts where a marker is drawn, never which point it is.
  const cap = spreadCap(cfg);
  const spread = new Map<number, number>();
  if (cap) {
    const byX = cap.axis === "x";
    // The cap is quoted in data units; the relaxation works in px.
    const limitPx = byX ? (cap.limit / (x1 - x0 || 1)) * plot.w : (cap.limit / (y1 - y0 || 1)) * plot.h;
    const disp = spreadAlongAxis(
      pts.map((p) => ({ m: byX ? toX(p.x) : toY(p.y), c: byX ? toY(p.y) : toX(p.x), r: radius(p) })),
      byX
        ? { limit: limitPx, min: plot.x, max: plot.x + plot.w }
        : { limit: limitPx, min: plot.y, max: plot.y + plot.h },
    );
    disp.forEach((d, i) => spread.set(i, d));
  }
  /** Drawn position: the exact one, plus any disclosed nudge on the spread axis. */
  const px = (p: (typeof pts)[number], i: number) => toX(p.x) + (cap?.axis === "x" ? (spread.get(i) ?? 0) : 0);
  const py = (p: (typeof pts)[number], i: number) => toY(p.y) + (cap?.axis === "y" ? (spread.get(i) ?? 0) : 0);

  // Bubble size legend: without a key, bubble AREA is unreadable. Two
  // outline reference circles (a nice maximum and its half), top-right.
  const legendBoxes: Box[] = [];
  if (cfg.kind === "bubble" && pts.some((p) => p.size != null)) {
    const sizeFmt = resolveFormat(
      pts.map((p) => Math.abs(p.size ?? 0)),
      cfg.numberFormat,
    );
    const refMax = niceTicks(0, maxSize, 3).pop()!;
    const refs = [refMax, refMax / 2];
    let lx = plot.x + plot.w - 4;
    refs.forEach((v, i) => {
      const r = Math.max(2.5, Math.sqrt(v / maxSize) * maxR);
      const cx = lx - r;
      const cy = plot.y + maxR * 1.1 + (Math.sqrt(refMax / maxSize) * maxR - r); // bottom-aligned circles
      nodes.push(
        {
          kind: "ellipse",
          cx,
          cy,
          rx: r,
          ry: r,
          fill: "none",
          stroke: style.mutedText,
          strokeWidth: 1,
          name: `size-legend-${i}`,
        },
        {
          kind: "text",
          x: cx - r,
          y: cy - Math.sqrt(refMax / maxSize) * maxR - fs * 1.35,
          w: r * 2,
          h: fs * 1.2,
          text: formatNumber(v, sizeFmt),
          fontSize: fs * 0.8,
          color: style.mutedText,
          align: "center",
          valign: "bottom",
          name: `size-legend-label-${i}`,
        },
      );
      legendBoxes.push({ x: cx - r, y: cy - r - fs * 1.4, w: r * 2, h: r * 2 + fs * 1.4 });
      lx = cx - r - fs * 0.8;
    });
  }

  // Trajectory / trail: connect the points in datasheet (row) order with a
  // direction arrowhead at each segment midpoint, drawn behind the markers —
  // a Gapminder-style path of one entity through the X/Y space over time.
  if (decor.trajectory && pts.length > 1) {
    for (let i = 0; i < pts.length - 1; i++) {
      const ax = toX(pts[i].x);
      const ay = toY(pts[i].y);
      const bx = toX(pts[i + 1].x);
      const by = toY(pts[i + 1].y);
      nodes.push({
        kind: "line",
        x1: ax,
        y1: ay,
        x2: bx,
        y2: by,
        stroke: style.mutedText,
        strokeWidth: 1.5,
        name: `trajectory-${i}`,
      });
      const angle = (Math.atan2(by - ay, bx - ax) * 180) / Math.PI;
      nodes.push({
        kind: "arrowhead",
        x: (ax + bx) / 2,
        y: (ay + by) / 2,
        angle,
        size: 4,
        fill: style.mutedText,
        name: `trajectory-head-${i}`,
      });
    }
  }

  // Marginal distribution histograms in the reserved gutters. The bin count is
  // a multiple of the axis's own tick intervals, so every tick is a bin edge
  // and a bar reads straight against the scale beside it. A rule keyed off the
  // sample size alone (Sturges, Freedman-Diaconis) puts edges BETWEEN the
  // ticks, which is exactly what a chart-adjacent histogram must not do. The
  // multiplier is the only freedom and it stays bounded: past ~2 sub-bins per
  // interval the bars are a few points wide and read as noise, not shape.
  if (mTop > 0 || mRight > 0) {
    const binMult = pts.length >= 15 ? 2 : 1;
    const fill = lerpColor("#ffffff", (cfg.style?.palette ?? PALETTE)[0], 0.35);
    if (mTop > 0) {
      const counts = histogramBins(
        pts.map((p) => p.x),
        x0,
        x1,
        (xTicks.length - 1) * binMult,
      );
      const peak = Math.max(1, ...counts);
      const bw = plot.w / counts.length;
      counts.forEach((n, i) => {
        if (!n) return;
        const h = (n / peak) * (mTop - 5);
        nodes.push({
          kind: "rect",
          x: plot.x + i * bw,
          y: plot.y - 3 - h,
          w: Math.max(0.5, bw - 1),
          h,
          fill,
          name: `marginal-x-${i}`,
        });
      });
    }
    if (mRight > 0) {
      const counts = histogramBins(
        pts.map((p) => p.y),
        y0,
        y1,
        (yTicks.length - 1) * binMult,
      );
      const peak = Math.max(1, ...counts);
      const bh = plot.h / counts.length;
      counts.forEach((n, i) => {
        if (!n) return;
        const w = (n / peak) * (mRight - 5);
        // Bin 0 is the bottom of the y axis, so it is the LAST band down the plot.
        nodes.push({
          kind: "rect",
          x: plot.x + plot.w + 3,
          y: plot.y + plot.h - (i + 1) * bh,
          w,
          h: Math.max(0.5, bh - 1),
          fill,
          name: `marginal-y-${i}`,
        });
      });
    }
  }

  // Point labels treat the size legend as an obstacle.
  const markerBoxes: Box[] = [...legendBoxes];
  // Paint back-to-front. Emitting in datasheet order lets a large bubble drawn
  // early bury a smaller one drawn later — completely, in all three renderers,
  // with nothing to tell the reader a point is missing. Largest first puts the
  // big ones at the back. This is paint order only: every marker keeps its
  // datasheet index in its name and its exact position. Ties keep datasheet
  // order (Array.sort is stable), so the result stays deterministic. Labels
  // already run biggest-first below, for the same reason.
  const paintOrder = pts.map((_, i) => i).sort((a, b) => radius(pts[b]) - radius(pts[a]));
  for (const i of paintOrder) {
    const p = pts[i];
    const r = radius(p);
    const gi = p.group - 1;
    const fill =
      colorScale && p.color != null
        ? colorScale.of(p.color)
        : colorScale
          ? style.mutedText
          : (cfg.style?.palette ?? PALETTE)[gi % 8];
    nodes.push(markerNode(markerFor(p.group), px(p, i), py(p, i), r, fill, style.background, 1, `point-${i}`));
    // Keep labels off the mark as DRAWN: an area-matched star reaches ~1.67x
    // its data radius, and a keep-out box built from `r` would let a label sit
    // on its points.
    const mr = markerExtent(markerFor(p.group), r);
    markerBoxes.push({ x: px(p, i) - mr, y: py(p, i) - mr, w: mr * 2, h: mr * 2 });
  }

  // Greedy label placement, biggest bubbles first so important points win.
  if (decor.segmentLabels !== false) {
    const order = pts.map((_, i) => i).sort((a, b) => radius(pts[b]) - radius(pts[a]));
    // Point label content: category (default), optionally with "(x, y)".
    const pointLabel = (p: (typeof pts)[number]) => {
      const parts = decor.labelContent ?? ["category"];
      const fmt = resolveFormat([p.x, p.y], cfg.numberFormat);
      return parts
        .map((part) =>
          part === "category"
            ? p.label
            : part === "value"
              ? `(${formatNumber(p.x, fmt)}, ${formatNumber(p.y, fmt)})`
              : null,
        )
        .filter(Boolean)
        .join(" ");
    };
    const reqs: LabelRequest[] = order.map((i) => ({
      cx: toX(pts[i].x),
      cy: toY(pts[i].y),
      r: radius(pts[i]),
      w: textWidth(pointLabel(pts[i]), fs) + 2,
      h: fs * 1.3,
    }));
    for (const placed of placeLabels(reqs, { x: 0, y: plot.y, w: cfg.width, h: plot.h + fs * 1.5 }, markerBoxes)) {
      const p = pts[order[placed.index]];
      nodes.push({
        kind: "text",
        x: placed.box.x,
        y: placed.box.y,
        w: placed.box.w,
        h: placed.box.h,
        text: pointLabel(p),
        fontSize: fs,
        color: style.text,
        align: "left",
        valign: "middle",
        name: `label-${order[placed.index]}`,
      });
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: pts.map((p) => toX(p.x)),
      categoryWidth: pts.map((p) => radius(p) * 2),
      columnTop: pts.map((p) => toY(p.y)),
      columnValue: pts.map((p) => p.y),
      baselineY: plot.y + plot.h,
      plot,
    },
  };
}
