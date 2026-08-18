import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { clipToWidth } from "../elements";
import { formatNumber, resolveFormat } from "../format";
import { maxOf, minOf } from "../agg";
import { lerpColor, noDataFill, sequentialScale } from "../color";
import { seriesColor } from "../style";
import { detectLayout, TILE_LAYOUTS } from "./tilemap-layouts";
import { bandFontSize, fitPlot, footnoteH, MIN_LABEL_FS, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

/**
 * What the legend below the grid actually occupies, in font sizes — not a round
 * number close to it. It starts `fs * 0.5` under the grid, the swatch strip is
 * `fs * 0.9`, and the min/max labels sit at `fs * 0.95` below the strip's top in
 * a box `fs * 1.2` tall, so the last ink is `fs * 2.65` under the grid.
 *
 * Both the reservation above the legend and the clamp that keeps it inside the
 * frame are stated in terms of this, so they cannot drift apart: the reservation
 * carries an extra gutter of slack, and a clamp written against the reservation
 * would spend that slack and move every ordinary tilemap by half a point.
 */
const LEGEND_INK = 2.65;

/**
 * Tile-grid cartogram ("map chart"): every region is a uniform square, so
 * geography stays recognizable without area distortion — and it renders as
 * native shapes everywhere (no freeform paths needed). Categories are region
 * codes (US postal / ISO-2 / world macro-regions); the single series' values
 * color the tiles on a sequential scale. Regions in the layout without data
 * render as gray "no data" tiles.
 */
export function layoutTilemap(cfg: ChartConfig, style: ChartStyle, decor: Decorations): LayoutResult {
  const { data } = cfg;
  const fs = style.fontSize;
  const layoutKey = cfg.map ?? detectLayout(data.categories);
  const layout = layoutKey ? TILE_LAYOUTS[layoutKey] : undefined;

  const nodes: SceneNode[] = [];
  const titleH = titleHeight(cfg, style);
  const titleN = titleNode(cfg, style);
  if (titleN) nodes.push(titleN);

  const empty: LayoutResult = {
    nodes,
    anchors: {
      categoryX: data.categories.map(() => cfg.width / 2),
      categoryWidth: data.categories.map(() => 10),
      columnTop: data.categories.map(() => titleH),
      columnValue: data.categories.map((_, c) => data.series[0]?.values[c] ?? 0),
      baselineY: cfg.height,
      // Fitted because the decorations read this rect: a frame shorter than its
      // own title would otherwise hand them a negative height to map bands
      // through.
      plot: fitPlot(cfg, { x: 0, y: titleH, w: cfg.width, h: cfg.height - titleH }),
    },
  };
  if (!layout) {
    nodes.push({
      kind: "text",
      x: 0,
      y: titleH + fs,
      w: cfg.width,
      h: fs * 1.5,
      text: 'No recognized region codes — set map: "us" | "eu" | "europe" | "world"',
      fontSize: fs,
      color: style.mutedText,
      align: "center",
      valign: "middle",
      name: "tilemap-error",
    });
    return empty;
  }

  // Value per region code (uppercased), from the first series.
  const values = new Map<string, number>();
  data.categories.forEach((code, c) => {
    const v = data.series[0]?.values[c];
    if (v != null && code.trim().toUpperCase() in layout) values.set(code.trim().toUpperCase(), v);
  });
  const vals = [...values.values()];
  const min = minOf(vals);
  const max = maxOf(vals);
  const base = data.series[0]?.color ?? style.palette[0];
  // Constant data: a flat mid-tone beats a zero-width ramp.
  const fill =
    min === max ? () => lerpColor(style.background, base, 0.5) : sequentialScale(min, max, base, style.background);
  const fmt = resolveFormat(vals, cfg.numberFormat);

  const topts = cfg.tilemap ?? {};
  const hex = topts.shape === "hex";
  const glyph = topts.glyph === "bars" && data.series.length > 1;
  // Mini-glyph mode: each region carries a series of values (one bar each).
  const seriesVals = new Map<string, (number | null)[]>();
  if (glyph) {
    data.categories.forEach((code, c) => {
      const key = code.trim().toUpperCase();
      if (key in layout)
        seriesVals.set(
          key,
          data.series.map((s) => s.values[c] ?? null),
        );
    });
  }
  const glyphVals = glyph ? [...seriesVals.values()].flat().filter((v): v is number => v != null) : [];
  const glyphMax = maxOf(glyphVals, 1);
  // The mini-bar scale spans zero: a negative value draws BELOW the tile's zero
  // line instead of clamping to zero height, which was pixel-identical to "no
  // change". All-positive data keeps glyphMin = 0, i.e. the original scale.
  const glyphMin = Math.min(0, minOf(glyphVals, 0));
  const glyphSpan = glyphMax - glyphMin;

  // Fit uniform square tiles into the plot area.
  const cols = Math.max(...Object.values(layout).map(([c]) => c)) + 1;
  const rows = Math.max(...Object.values(layout).map(([, r]) => r)) + 1;
  const gutter = 2.5;
  // The reservation is the legend's own ink (see LEGEND_INK) plus one gutter of
  // slack, because `rowsBottom` adds a gutter per row where the height budget
  // pays for the gaps BETWEEN rows — one gutter more than it is given. Those two
  // together used to predict `fs * 0.0785 + 0.5` of overflow, and did: the min
  // and max labels had their descenders cut by the frame on EVERY tilemap at
  // EVERY size, 1.3pt at a 10pt font and 1.9 at 18, the measurement to the tenth.
  const legendH = vals.length ? LEGEND_INK * fs + gutter : fs * 0.5;
  const availW = cfg.width - 4;
  const availH = cfg.height - titleH - legendH - footnoteH(cfg, style, decor) - 4;
  // Hex tiles nest: rows step ~0.87·tile and odd rows shift half a column, so
  // the footprint needs an extra half column of width and less height.
  // The 6pt floor keeps a tile visible on an ordinary chart, and it is the last
  // place in this engine that refused to yield to its frame: on a frame too
  // small to pay for it, honouring the floor made the grid taller than the space
  // budgeted for it and pushed the legend under it off the bottom — 27pt on a
  // 120x90 thumbnail. A small tile is still legible as a SHAPE, which is what a
  // cartogram is read for; a tile drawn outside the frame is not there at all.
  //
  // So the floor applies only when the frame can pay for it. It is unreachable
  // on any chart big enough to want it, so nothing of an ordinary size moves.
  const want = hex
    ? Math.min((availW - (cols - 1) * gutter) / (cols + 0.5), availH / ((rows - 1) * 0.87 + 1))
    : Math.min((availW - (cols - 1) * gutter) / cols, (availH - (rows - 1) * gutter) / rows);
  const tile = Math.max(1, want);
  const gridW = (hex ? cols + 0.5 : cols) * tile + (cols - 1) * gutter;
  const x0 = (cfg.width - gridW) / 2;
  // The grid's own height, and a top edge that keeps it on the canvas. The tile
  // floor above is deliberately allowed to overrun the height budget, and on a
  // frame whose title and legend leave nothing it overran the CANVAS too: the
  // bottom row was drawn up to 24.7pt below the foot of a 300x60 chart. Pushed
  // up rather than shrunk further — the same call the comment above makes, since
  // a tile over the title is still a tile and one off the chart is not there.
  const gridH = hex ? (rows - 1) * tile * 0.87 + tile : rows * (tile + gutter);
  const y0 = Math.max(0, Math.min(titleH + 2, cfg.height - gridH));
  const rowsBottom = y0 + gridH;
  /**
   * Where the legend hangs from. Normally the bottom of the grid — but the tile
   * floor above is explicitly allowed to overrun `availH`, and when it does the
   * grid drags the legend off the frame with it: 17pt below a 300x60 chart. So
   * the legend takes the grid's bottom or the last position its own reserved
   * band still fits at, whichever is higher. On any chart whose grid fits its
   * budget the first is always the smaller, so nothing of an ordinary size moves.
   */
  const legendTop = Math.max(0, Math.min(rowsBottom, cfg.height - LEGEND_INK * fs - footnoteH(cfg, style, decor)));

  const hexPts = (cx: number, cy: number, R: number) =>
    [90, 150, 210, 270, 330, 30].map((a) => ({
      x: cx + R * Math.cos((a * Math.PI) / 180),
      y: cy - R * Math.sin((a * Math.PI) / 180),
    }));
  for (const [code, [col, row]] of Object.entries(layout)) {
    const v = values.get(code);
    // In glyph mode the tile is a faint backdrop for the bars; otherwise it
    // carries the value color.
    const tileFill = glyph
      ? lerpColor(style.background, base, 0.1)
      : v == null || !vals.length
        ? noDataFill(style.background)
        : fill(v);
    const x = x0 + col * (tile + gutter) + (hex && row % 2 === 1 ? (tile + gutter) / 2 : 0);
    const y = hex ? y0 + row * tile * 0.87 : y0 + row * (tile + gutter);
    if (hex) {
      nodes.push({
        kind: "polygon",
        points: hexPts(x + tile / 2, y + tile / 2, tile / 2),
        fill: tileFill,
        // Outline in the tile's OWN colour, not the background: Office.js has no
        // freeform fill and degrades a polygon to its stroke (scene.ts's parity
        // contract), so a background-coloured edge left the whole cartogram
        // white-on-white in the add-in. The tiles never touch — the grid steps
        // tile+gutter across and 0.87·tile down, both wider than the hex — so
        // the separator the background stroke used to draw is not needed.
        stroke: tileFill,
        strokeWidth: 1,
        name: `tile-${code}`,
      });
    } else {
      nodes.push({ kind: "rect", x, y, w: tile, h: tile, fill: tileFill, name: `tile-${code}` });
    }
    const ink = contrastInk(tileFill);
    // Mini bar glyph: one bar per series, from the region's row of values.
    if (glyph) {
      const svals = seriesVals.get(code);
      if (svals && tile >= fs * 2) {
        const nb = svals.length;
        const bw = (tile * 0.78) / nb;
        const bx0 = x + tile * 0.11;
        const bBase = y + tile * 0.86;
        const bMax = tile * 0.5;
        const zeroY = bBase + (glyphMin / glyphSpan) * bMax;
        svals.forEach((sv, si) => {
          if (sv == null) return;
          const vy = bBase - ((sv - glyphMin) / glyphSpan) * bMax;
          nodes.push({
            kind: "rect",
            x: bx0 + si * bw,
            y: Math.min(vy, zeroY),
            w: Math.max(1, bw - 0.5),
            h: Math.abs(vy - zeroY),
            fill: seriesColor(style, si),
            name: `glyph-${code}-${si}`,
          });
        });
      }
      // A tile's code is bounded by the TILE, and a grid squeezed onto a sliver
      // of a frame gives tiles of a point or two — so this answered fonts below
      // anything readable, and below anything OOXML accepts. Dropped instead:
      // the tile still carries the value through its fill.
      const glyphCodeFs = bandFontSize(fs * 0.85, tile, 1 / 0.3);
      if (glyphCodeFs > 0)
        nodes.push({
          kind: "text",
          x,
          y: y + tile * 0.06,
          w: tile,
          h: fs * 1.2,
          text: code,
          fontSize: glyphCodeFs,
          bold: true,
          color: ink,
          align: "center",
          valign: "middle",
          name: `tile-code-${code}`,
        });
      continue;
    }
    const showValue =
      v != null && decor.segmentLabels && tile >= fs * 2.6 && textWidth(formatNumber(v, fmt), fs * 0.8) <= tile - 2;
    const codeFs = bandFontSize(fs, tile, 1 / 0.34);
    if (codeFs > 0)
      nodes.push({
        kind: "text",
        x,
        y: showValue ? y + tile / 2 - fs * 1.25 : y,
        w: tile,
        h: showValue ? fs * 1.3 : tile,
        text: code,
        fontSize: codeFs,
        bold: true,
        color: ink,
        align: "center",
        valign: "middle",
        name: `tile-code-${code}`,
      });
    if (showValue) {
      nodes.push({
        kind: "text",
        x,
        y: y + tile / 2,
        w: tile,
        h: fs * 1.2,
        text: formatNumber(v, fmt),
        fontSize: fs * 0.8,
        color: ink,
        align: "center",
        valign: "top",
        name: `tile-value-${code}`,
      });
    }
  }

  // Glyph mode: a series legend instead of the value gradient.
  if (glyph) {
    let lx = x0;
    data.series.forEach((s, si) => {
      const chip = fs * 0.7;
      nodes.push(
        {
          kind: "rect",
          x: lx,
          y: legendTop + fs * 0.6,
          w: chip,
          h: chip,
          fill: seriesColor(style, si, s.color),
          name: `legend-chip-${si}`,
        },
        {
          kind: "text",
          x: lx + chip + 3,
          y: legendTop + fs * 0.3,
          w: textWidth(s.name, fs) + 6,
          h: fs * 1.4,
          text: s.name,
          fontSize: fs * 0.85,
          color: style.text,
          align: "left",
          valign: "middle",
          name: `legend-${si}`,
        },
      );
      lx += chip + 3 + textWidth(s.name, fs) + 12;
    });
  }
  // Gradient legend + "no data" swatch.
  if (!glyph && vals.length && min !== max) {
    const ly = legendTop + fs * 0.5;
    const lw = Math.min(gridW * 0.5, fs * 12);
    const steps = 24;
    for (let i = 0; i < steps; i++) {
      const v = min + ((max - min) * i) / (steps - 1);
      nodes.push({
        kind: "rect",
        x: x0 + (lw / steps) * i,
        y: ly,
        w: lw / steps + 0.5,
        h: fs * 0.9,
        fill: fill(v),
        name: `legend-step-${i}`,
      });
    }
    // The two ends of the scale each own HALF the bar and are anchored to its
    // outer edges, so their ink meets in the middle the moment a number is
    // wider than `lw / 2`. The bar is `min(gridW * 0.5, fs * 12)`, so on a
    // small frame it is a few points wide and the two numbers are drawn on top
    // of each other — inside the frame, so no overflow gate could see it.
    //
    // Same answer the rest of this engine gives: shrink both together to the
    // room they actually have, and drop the PAIR when that would be illegible.
    // Both, never one: a gradient bar labelled at one end says the wrong thing,
    // where an unlabelled bar just says less.
    const endText = [formatNumber(min, fmt), formatNumber(max, fmt)];
    const endFs = (() => {
      const base = fs * 0.85;
      const widest = Math.max(...endText.map((t) => textWidth(t, base)));
      if (!(widest > 0)) return base;
      const f = Math.min(base, base * (lw / 2 / widest));
      return f >= MIN_LABEL_FS ? f : 0;
    })();
    if (endFs > 0)
      nodes.push(
        {
          kind: "text",
          x: x0,
          y: ly + fs * 0.95,
          w: lw / 2,
          h: fs * 1.2,
          text: endText[0],
          fontSize: endFs,
          color: style.mutedText,
          align: "left",
          valign: "top",
          name: "legend-min",
        },
        {
          kind: "text",
          x: x0 + lw / 2,
          y: ly + fs * 0.95,
          w: lw / 2,
          h: fs * 1.2,
          text: endText[1],
          fontSize: endFs,
          color: style.mutedText,
          align: "right",
          valign: "top",
          name: "legend-max",
        },
      );
    if (values.size < Object.keys(layout).length) {
      // The no-data key is a swatch and a caption, and BOTH have to fit to the
      // right of the gradient bar. Bounding the caption's box was not enough:
      // its box was clamped to a point wide and clipped to the empty string,
      // and an empty text box still has an ORIGIN — 27pt past the right edge of
      // an 80x60 chart at 24pt, with the swatch itself hanging 5pt over. A node
      // that draws nothing is still a shape on the slide, and the overflow gate
      // reads it as one.
      //
      // So this is the reservation answer the radar, sunburst, pie and the
      // scale ends above already give: shrink the caption to the room it has,
      // and where that room cannot carry a readable one, DROP THE PAIR. A
      // swatch with no caption says nothing; the tiles it explains are still
      // gray, and the reader loses a legend entry rather than gaining ink
      // outside the chart.
      const swatchX = x0 + lw + fs;
      const textX = x0 + lw + fs * 2.1;
      const room = cfg.width - textX - 2;
      const capFs = (() => {
        const base = fs * 0.85;
        const w = textWidth("no data", base);
        if (!(w > 0)) return base;
        const f = Math.min(base, base * (room / w));
        return f >= MIN_LABEL_FS ? f : 0;
      })();
      if (capFs > 0 && swatchX + fs * 0.9 <= cfg.width)
        nodes.push(
          {
            kind: "rect",
            x: swatchX,
            y: ly,
            w: fs * 0.9,
            h: fs * 0.9,
            fill: noDataFill(style.background),
            name: "legend-nodata",
          },
          {
            kind: "text",
            x: textX,
            y: ly - fs * 0.2,
            // Bounded by the room actually left to the right of the swatch. The
            // flat `fs * 6` is 192 points wide at a 32pt font and starts 67
            // points in, so on a narrow chart the caption ran 37pt off the right
            // edge. Last resort in both directions: where the flat box fits it
            // is still the flat box, so no chart of an ordinary size moves.
            w: Math.min(fs * 6, room),
            h: fs * 1.3,
            text: clipToWidth("no data", capFs, Math.min(fs * 6, room)),
            fontSize: capFs,
            color: style.mutedText,
            align: "left",
            valign: "middle",
            name: "legend-nodata-label",
          },
        );
    }
  }

  return {
    ...empty,
    nodes,
  };
}
