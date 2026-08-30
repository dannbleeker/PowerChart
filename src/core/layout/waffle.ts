import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { textWidth, type SceneNode } from "../scene";
import { formatPercent } from "../format";
import { noDataFill } from "../color";
import { MIN_LABEL_FS, firstSeriesValues, footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * Waffle chart: a 10×10 unit grid where each cell is 1% — the honest
 * part-to-whole for a single dominant share. Categories are the parts
 * (pie semantics, first series only), rounded to whole cells by largest
 * remainder; unassigned cells stay no-data gray. Cells fill from the
 * bottom-left, row by row upward.
 *
 * Denominator: `100%=` row value when present, else the sum of the values —
 * except a single category ≤100 with no denominator, which reads as a
 * literal percentage ("68" → 68 filled cells). A stated denominator below the
 * parts' sum is inconsistent and undrawable, so the sum wins.
 */
export function layoutWaffle(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const { plot: values, raw } = firstSeriesValues(data, data.categories.length);
  const sum = values.reduce((a, b) => a + b, 0);
  const single = data.categories.length === 1;
  const stated =
    data.hundredPercent?.[0] && data.hundredPercent[0] > 0
      ? data.hundredPercent[0]
      : single && sum <= 100
        ? 100
        : sum || 1;
  // A denominator smaller than the parts cannot be drawn: the quotas sum past
  // 100 cells, the largest-remainder pass is skipped and the 100-cell cap
  // truncated whichever categories came LAST in input order — two equal values
  // then got wildly different areas while the legend printed the same percent.
  // Fall back to the parts' own sum so the picture and the printed shares agree.
  const denom = Math.max(stated, sum);

  // Whole cells per category via largest remainder, so the filled count is
  // exactly round(100 · sum/denom).
  const quotas = values.map((v) => (v / denom) * 100);
  const cells = quotas.map(Math.floor);
  const target = Math.min(100, Math.round(quotas.reduce((a, b) => a + b, 0)));
  const byRemainder = quotas.map((q, i) => ({ r: q - Math.floor(q), i })).sort((a, b) => b.r - a.r);
  for (let k = 0; cells.reduce((a, b) => a + b, 0) < target && k < byRemainder.length; k++) {
    cells[byRemainder[k].i]++;
  }

  const titleH = titleHeight(cfg, style);
  const legendEntries = data.categories.map((name, c) => ({
    name,
    // A share DERIVED from an unknown is unknown. A blank cell used to read
    // "Signups  0%", which states a measurement that was never taken.
    pct:
      raw[c] === null
        ? ""
        : formatPercent(values[c] / denom, quotas[c] > 0 && quotas[c] < 1 ? 1 : 0, false, cfg.numberFormat?.locale),
    color: data.series[0]?.colors?.[c] ?? style.palette[c % style.palette.length],
  }));
  const legendW =
    decor.seriesLabels === false
      ? 0
      : Math.max(...legendEntries.map((e) => textWidth(`${e.name}  ${e.pct}`.trimEnd(), fs))) + fs * 2.2;

  const footH = footnoteH(cfg, style, decor);
  const availH = cfg.height - titleH - footH - 8;
  /**
   * Hold a legend block of height `blockH` beside the grid: centred on the
   * grid's middle where the frame can pay for it, otherwise as close to that as
   * the frame allows.
   *
   * The grid has a 20pt floor and the legend does not scale with it, so on a
   * short frame the block is TALLER than the square it is centred on and hangs
   * out of the chart — 3.2pt below a 60pt-tall waffle. Sliding it is safe in a
   * way clipping is not: these labels are beside the grid, not anchored to any
   * cell in it.
   */
  const legendBandY = (blockH: number) =>
    Math.max(0, Math.min(gy + gridSize / 2 - blockH / 2, cfg.height - footH - blockH));
  const availW = cfg.width - legendW - 8;
  const gridSize = Math.max(20, Math.min(availW, availH));
  const step = gridSize / 10;
  const cellSize = step * 0.86;
  const gx = 2;
  // Kept on the canvas. The grid has a 20pt floor so it never collapses, and on
  // a frame whose title and footnote leave less than that the centring put it
  // partly below the foot of the chart — 8.3pt at a 32pt font on a 60pt-tall
  // waffle. The grid IS the chart, so it is clamped rather than shrunk further
  // or dropped.
  const gy = Math.max(0, Math.min(titleH + 4 + (availH - gridSize) / 2, cfg.height - gridSize));

  const nodes: SceneNode[] = [];
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  // Cell colors in fill order: category 0 first, gray remainder.
  const fills: string[] = [];
  legendEntries.forEach((e, c) => {
    for (let k = 0; k < cells[c] && fills.length < 100; k++) fills.push(e.color);
  });
  while (fills.length < 100) fills.push(noDataFill(style.background));

  for (let k = 0; k < 100; k++) {
    const row = Math.floor(k / 10); // 0 = bottom row
    const col = k % 10;
    nodes.push({
      kind: "rect",
      x: gx + col * step,
      y: gy + (9 - row) * step,
      w: cellSize,
      h: cellSize,
      fill: fills[k],
      name: `waffle-cell-${k}`,
    });
  }

  /**
   * The font the legend block is drawn at.
   *
   * `legendBandY` slides the block to keep it on the canvas, and sliding cannot
   * help once the block is TALLER than the canvas: it is priced in the chart
   * font — `fs * 4.8` for the big-number treatment, `entries * fs * 1.7` for the
   * chip rows — with nothing bounding it against the height it has to live in.
   * At a 32pt font on a 60pt-tall waffle the single-category block wants 154
   * points of 60, and its name was drawn 87.8pt below the foot of the chart.
   *
   * Shrunk to the room, then dropped past the floor — the order every fit in
   * this engine uses. Where the block already fits, which is every chart at an
   * ordinary font, this is `fs` exactly and nothing moves.
   */
  const legendRoom = Math.max(0, cfg.height - footH - titleH);
  const legendBlockEm = single ? 4.8 : legendEntries.length * 1.7;
  const legendFs = Math.min(fs, legendRoom / (legendBlockEm || 1));
  // Legend: single category gets the big-number treatment, several get chips.
  if (legendW > 0 && legendFs >= MIN_LABEL_FS) {
    const fs = legendFs;
    const lx = gx + gridSize + fs;
    if (single) {
      // The big-number block spans fs*2.4 either side of the grid's middle: the
      // percentage's box starts there and the name's box ends fs*2.4 past it.
      const bigY = legendBandY(fs * 4.8);
      nodes.push({
        kind: "text",
        x: lx,
        y: bigY,
        w: cfg.width - lx - 2,
        h: fs * 3.2,
        text: legendEntries[0].pct,
        fontSize: fs * 2.6,
        bold: true,
        color: legendEntries[0].color,
        align: "left",
        valign: "middle",
        name: "waffle-big-pct",
      });
      nodes.push({
        kind: "text",
        x: lx,
        y: bigY + fs * 3.3,
        w: cfg.width - lx - 2,
        h: fs * 1.5,
        text: legendEntries[0].name,
        fontSize: fs,
        color: style.mutedText,
        align: "left",
        valign: "middle",
        name: "legend-label-0",
      });
    } else {
      const rowH = fs * 1.7;
      const ly = legendBandY(legendEntries.length * rowH);
      legendEntries.forEach((e, c) => {
        const y = ly + c * rowH;
        nodes.push({
          kind: "rect",
          x: lx,
          y: y + rowH / 2 - fs * 0.45,
          w: fs * 0.9,
          h: fs * 0.9,
          fill: e.color,
          name: `legend-chip-${c}`,
        });
        nodes.push({
          kind: "text",
          x: lx + fs * 1.3,
          y,
          w: cfg.width - lx - fs * 1.3 - 2,
          h: rowH,
          // Trimmed: a category whose value is unknown has no percentage, and the
          // separator alone would read as a truncated number.
          text: `${e.name}  ${e.pct}`.trimEnd(),
          fontSize: fs,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `legend-label-${c}`,
        });
      });
    }
  }

  return {
    nodes,
    anchors: {
      categoryX: data.categories.map(() => gx + gridSize / 2),
      categoryWidth: data.categories.map(() => gridSize),
      columnTop: data.categories.map(() => gy),
      columnValue: values,
      baselineY: gy + gridSize,
      plot: { x: gx, y: gy, w: gridSize, h: gridSize },
    },
  };
}
