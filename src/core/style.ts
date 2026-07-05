import type { ChartStyle, Decorations } from "./types";

/**
 * Default series palette: the validated brand-neutral categorical set from the
 * dataviz reference (worst adjacent CVD ΔE 24.2 on a white surface; the three
 * sub-3:1-contrast hues are relieved by the engine's always-on direct labels).
 * Swap for the presentation's theme colors via ChartConfig.style.palette.
 */
export const PALETTE = [
  "#2a78d6", // blue
  "#1baf7a", // aqua
  "#eda100", // yellow
  "#008300", // green
  "#4a3aa7", // violet
  "#e34948", // red
  "#e87ba4", // magenta
  "#eb6834", // orange
];

/**
 * Palette presets. "Default" is the validated categorical set; the others are
 * corporate looks — swap in the presentation's accent colors via a custom
 * palette when needed (Office.js exposes no theme-color API).
 */
export const PALETTES: Record<string, string[]> = {
  Default: PALETTE,
  Blues: ["#0d366b", "#1c5cab", "#2a78d6", "#5598e7", "#86b6ef", "#b7d3f6", "#52514e", "#898781"],
  Warm: ["#b3541e", "#eb6834", "#eda100", "#e34948", "#8a6d3b", "#d5a021", "#52514e", "#898781"],
  Grayscale: ["#1a1a19", "#3d3d3b", "#5f5e5b", "#82817d", "#a5a49e", "#c3c2b7", "#dcdbd2", "#8a3ffc"],
};

export const DEFAULT_STYLE: ChartStyle = {
  fontFamily: "Segoe UI, Arial, sans-serif",
  fontSize: 10,
  palette: PALETTE,
  text: "#0b0b0b",
  mutedText: "#52514e",
  axis: "#a5a49e",
  gridline: "#e1e0d9",
  negative: "#e34948",
  neutral: "#898781",
  background: "#ffffff",
};

export const DEFAULT_DECOR: Decorations = {
  segmentLabels: true,
  seriesLabels: true,
  totals: false,
  categoryAxis: true,
  valueAxis: false,
  gridlines: false,
};

export function seriesColor(style: ChartStyle, index: number, override?: string): string {
  return override ?? style.palette[index % style.palette.length];
}
