import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { maxOf, minOf } from "../agg";
import { lerpColor, noDataFill, sequentialScale } from "../color";
import { seriesColor } from "../style";
import { detectLayout, TILE_LAYOUTS } from "./tilemap-layouts";
import { footnoteH, titleHeight, titleNode } from "./frame";
import type { LayoutResult } from "./column";

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
      plot: { x: 0, y: titleH, w: cfg.width, h: cfg.height - titleH },
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
  const legendH = vals.length ? fs * 2.4 : fs * 0.5;
  const availW = cfg.width - 4;
  const availH = cfg.height - titleH - legendH - footnoteH(cfg, style, decor) - 4;
  const gutter = 2.5;
  // Hex tiles nest: rows step ~0.87·tile and odd rows shift half a column, so
  // the footprint needs an extra half column of width and less height.
  const tile = hex
    ? Math.max(6, Math.min((availW - (cols - 1) * gutter) / (cols + 0.5), availH / ((rows - 1) * 0.87 + 1)))
    : Math.max(6, Math.min((availW - (cols - 1) * gutter) / cols, (availH - (rows - 1) * gutter) / rows));
  const gridW = (hex ? cols + 0.5 : cols) * tile + (cols - 1) * gutter;
  const x0 = (cfg.width - gridW) / 2;
  const y0 = titleH + 2;
  const rowsBottom = hex ? y0 + (rows - 1) * tile * 0.87 + tile : y0 + rows * (tile + gutter);

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
      nodes.push({
        kind: "text",
        x,
        y: y + tile * 0.06,
        w: tile,
        h: fs * 1.2,
        text: code,
        fontSize: Math.min(fs * 0.85, tile * 0.3),
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
    nodes.push({
      kind: "text",
      x,
      y: showValue ? y + tile / 2 - fs * 1.25 : y,
      w: tile,
      h: showValue ? fs * 1.3 : tile,
      text: code,
      fontSize: Math.min(fs, tile * 0.34),
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
          y: rowsBottom + fs * 0.6,
          w: chip,
          h: chip,
          fill: seriesColor(style, si, s.color),
          name: `legend-chip-${si}`,
        },
        {
          kind: "text",
          x: lx + chip + 3,
          y: rowsBottom + fs * 0.3,
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
    const ly = rowsBottom + fs * 0.5;
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
    nodes.push(
      {
        kind: "text",
        x: x0,
        y: ly + fs * 0.95,
        w: lw / 2,
        h: fs * 1.2,
        text: formatNumber(min, fmt),
        fontSize: fs * 0.85,
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
        text: formatNumber(max, fmt),
        fontSize: fs * 0.85,
        color: style.mutedText,
        align: "right",
        valign: "top",
        name: "legend-max",
      },
    );
    if (values.size < Object.keys(layout).length) {
      nodes.push(
        {
          kind: "rect",
          x: x0 + lw + fs,
          y: ly,
          w: fs * 0.9,
          h: fs * 0.9,
          fill: noDataFill(style.background),
          name: "legend-nodata",
        },
        {
          kind: "text",
          x: x0 + lw + fs * 2.1,
          y: ly - fs * 0.2,
          w: fs * 6,
          h: fs * 1.3,
          text: "no data",
          fontSize: fs * 0.85,
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
