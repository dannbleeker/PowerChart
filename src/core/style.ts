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
