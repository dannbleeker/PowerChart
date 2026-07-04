/** Chart kinds supported by the layout engine (think-cell equivalents in comments). */
export type ChartKind =
  | "stacked" // think-cell "Stacked" column chart
  | "clustered" // "Clustered"
  | "stacked100" // "100%" chart
  | "waterfall" // "Waterfall" (build up/down, "e" totals)
  | "mekko" // "Mekko" / Marimekko with %-axis
  | "line" // "Line"
  | "area" // "Area"
  | "butterfly" // two back-to-back bar charts sharing one scale
  | "scatter" // "Scatter" — X/Y rows, labelled points
  | "bubble" // "Bubble" — scatter + Size row
  | "gantt"; // simplified numeric-timeline Gantt (Start/End rows)

export interface Series {
  name: string;
  /** One value per category; null = blank cell. */
  values: (number | null)[];
  /** Override the palette color for this series. */
  color?: string;
}

export interface ChartData {
  categories: string[];
  series: Series[];
  /**
   * think-cell's "100%=" datasheet row: per-category denominators for 100%
   * charts. When set, columns whose series sum to less than this stay short
   * of full height. Defaults to the column sum.
   */
  hundredPercent?: (number | null)[];
  /**
   * think-cell's "X extent" row (Mekko with units): explicit column widths,
   * scaled so the total matches the chart width.
   */
  xExtent?: (number | null)[];
}

export interface WaterfallOptions {
  /**
   * Categories rendered as cumulative totals (think-cell's "e" cell):
   * the bar is drawn from the baseline to the running total at that point,
   * regardless of the cell value.
   */
  totalIndices?: number[];
}

/** think-cell style decorations, all computed from layout anchors. */
export interface Decorations {
  /** Value label inside each segment (hidden automatically when it doesn't fit). */
  segmentLabels: boolean;
  /** Series labels to the right of the last category (think-cell placement). */
  seriesLabels: boolean;
  /** Column totals above each column. */
  totals: boolean;
  /** Category labels below the baseline. */
  categoryAxis: boolean;
  /** Numeric value axis on the left (off by default, as in think-cell). */
  valueAxis: boolean;
  gridlines: boolean;
  /** Compound annual growth rate arrow between two category indices. */
  cagr?: { from: number; to: number };
  /**
   * Difference arrow between two category indices. Without `series` it is a
   * total difference arrow (column totals); with `series` it is a level
   * difference arrow comparing the cumulative level at that series.
   */
  difference?: { from: number; to: number; percent?: boolean; series?: number };
  /** Value lines: at fixed values or at the mean of column totals. */
  valueLines?: ({ mode: "mean" } | { mode: "value"; value: number })[];
  /** @deprecated legacy single value line; normalized into valueLines. */
  valueLine?: { mode: "mean" } | { mode: "value"; value: number };
}

export interface NumberFormat {
  /** Number of decimals; "auto" picks based on magnitude. */
  decimals: number | "auto";
  /** Append a suffix such as "%" or "€m". */
  suffix?: string;
  /** Show an explicit "+" on positive deltas (waterfall, difference arrows). */
  forceSign?: boolean;
}

export interface ChartStyle {
  fontFamily: string;
  /** Base label font size in points. */
  fontSize: number;
  palette: string[];
  /** Ink colors. */
  text: string;
  mutedText: string;
  axis: string;
  gridline: string;
  /** Fill for waterfall decreases / negative deltas. */
  negative: string;
  /** Fill for waterfall totals. */
  neutral: string;
  background: string;
}

export interface ChartConfig {
  kind: ChartKind;
  data: ChartData;
  /**
   * Rotate column charts into bar charts (think-cell's rotation handle).
   * Applies to stacked/clustered/100%; decorations that assume a vertical
   * value axis are skipped in horizontal orientation.
   */
  horizontal?: boolean;
  /**
   * Manual value-axis scale (think-cell's axis-handle dragging). Either end
   * may be pinned; the other stays automatic.
   */
  scale?: { min?: number; max?: number };
  /**
   * think-cell's Segment Order menu: stacking order of series within each
   * column. `ascending`/`descending` sort per column by value.
   */
  segmentOrder?: "sheet" | "reverse" | "ascending" | "descending";
  /** Frame size in points (PowerPoint native unit). */
  width: number;
  height: number;
  title?: string;
  decorations?: Partial<Decorations>;
  waterfall?: WaterfallOptions;
  numberFormat?: Partial<NumberFormat>;
  style?: Partial<ChartStyle>;
}

/** Geometry the decoration pass needs from a layout: where each column lives. */
export interface LayoutAnchors {
  /** Center x of each category slot. */
  categoryX: number[];
  /** Width of the column at each category. */
  categoryWidth: number[];
  /** y of the visual top of each column (min y across its segments). */
  columnTop: number[];
  /** Column total values (signed sum, or cumulative value for waterfall totals). */
  columnValue: number[];
  /** Per category: cumulative stack value after each series (for level arrows). */
  seriesLevels?: number[][];
  /** y coordinate of the zero baseline. */
  baselineY: number;
  /** Plot rectangle. */
  plot: { x: number; y: number; w: number; h: number };
  /** Value → y mapping used by the layout (absent for 100% charts). */
  valueToY?: (v: number) => number;
}
