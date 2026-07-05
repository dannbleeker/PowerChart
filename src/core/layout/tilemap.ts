import type { ChartConfig, ChartStyle, Decorations } from "../types";
import { contrastInk, textWidth, type SceneNode } from "../scene";
import { formatNumber, resolveFormat } from "../format";
import { lerpColor, NO_DATA, sequentialScale } from "../color";
import { detectLayout, TILE_LAYOUTS } from "./tilemap-layouts";
import { footnoteH } from "./frame";
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
  const titleH = cfg.title ? fs * 1.6 + 6 : 0;
  if (cfg.title) {
    nodes.push({
      kind: "text", x: 0, y: 0, w: cfg.width, h: fs * 1.6, text: cfg.title,
      fontSize: fs * 1.2, bold: true, color: style.text, align: "left", valign: "top", name: "title",
    });
  }

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
      kind: "text", x: 0, y: titleH + fs, w: cfg.width, h: fs * 1.5,
      text: 'No recognized region codes — set map: "us" | "eu" | "europe" | "world"',
      fontSize: fs, color: style.mutedText, align: "center", valign: "middle", name: "tilemap-error",
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
  const min = Math.min(...vals);
  const max = Math.max(...vals);
  const base = data.series[0]?.color ?? style.palette[0];
  // Constant data: a flat mid-tone beats a zero-width ramp.
  const fill = min === max ? () => lerpColor("#ffffff", base, 0.5) : sequentialScale(min, max, base);
  const fmt = resolveFormat(vals, cfg.numberFormat);

  // Fit uniform square tiles into the plot area.
  const cols = Math.max(...Object.values(layout).map(([c]) => c)) + 1;
  const rows = Math.max(...Object.values(layout).map(([, r]) => r)) + 1;
  const legendH = vals.length ? fs * 2.4 : fs * 0.5;
  const availW = cfg.width - 4;
  const availH = cfg.height - titleH - legendH - footnoteH(cfg, style, decor) - 4;
  const gutter = 2.5;
  const tile = Math.max(6, Math.min((availW - (cols - 1) * gutter) / cols, (availH - (rows - 1) * gutter) / rows));
  const gridW = cols * tile + (cols - 1) * gutter;
  const x0 = (cfg.width - gridW) / 2;
  const y0 = titleH + 2;

  for (const [code, [col, row]] of Object.entries(layout)) {
    const v = values.get(code);
    const tileFill = v == null || !vals.length ? NO_DATA : fill(v);
    const x = x0 + col * (tile + gutter);
    const y = y0 + row * (tile + gutter);
    nodes.push({ kind: "rect", x, y, w: tile, h: tile, fill: tileFill, name: `tile-${code}` });
    const ink = contrastInk(tileFill);
    const showValue = v != null && decor.segmentLabels && tile >= fs * 2.6 && textWidth(formatNumber(v, fmt), fs * 0.8) <= tile - 2;
    nodes.push({
      kind: "text", x, y: showValue ? y + tile / 2 - fs * 1.25 : y, w: tile, h: showValue ? fs * 1.3 : tile,
      text: code, fontSize: Math.min(fs, tile * 0.34), bold: true, color: ink,
      align: "center", valign: "middle", name: `tile-code-${code}`,
    });
    if (showValue) {
      nodes.push({
        kind: "text", x, y: y + tile / 2, w: tile, h: fs * 1.2, text: formatNumber(v, fmt),
        fontSize: fs * 0.8, color: ink, align: "center", valign: "top", name: `tile-value-${code}`,
      });
    }
  }

  // Gradient legend + "no data" swatch.
  if (vals.length && min !== max) {
    const ly = y0 + rows * (tile + gutter) + fs * 0.5;
    const lw = Math.min(gridW * 0.5, fs * 12);
    const steps = 24;
    for (let i = 0; i < steps; i++) {
      const v = min + ((max - min) * i) / (steps - 1);
      nodes.push({ kind: "rect", x: x0 + (lw / steps) * i, y: ly, w: lw / steps + 0.5, h: fs * 0.9, fill: fill(v), name: `legend-step-${i}` });
    }
    nodes.push(
      { kind: "text", x: x0, y: ly + fs * 0.95, w: lw / 2, h: fs * 1.2, text: formatNumber(min, fmt), fontSize: fs * 0.85, color: style.mutedText, align: "left", valign: "top", name: "legend-min" },
      { kind: "text", x: x0 + lw / 2, y: ly + fs * 0.95, w: lw / 2, h: fs * 1.2, text: formatNumber(max, fmt), fontSize: fs * 0.85, color: style.mutedText, align: "right", valign: "top", name: "legend-max" },
    );
    if (values.size < Object.keys(layout).length) {
      nodes.push(
        { kind: "rect", x: x0 + lw + fs, y: ly, w: fs * 0.9, h: fs * 0.9, fill: NO_DATA, name: "legend-nodata" },
        {
          kind: "text", x: x0 + lw + fs * 2.1, y: ly - fs * 0.2, w: fs * 6, h: fs * 1.3,
          text: "no data", fontSize: fs * 0.85, color: style.mutedText, align: "left", valign: "middle", name: "legend-nodata-label",
        },
      );
    }
  }

  return {
    ...empty,
    nodes,
  };
}
