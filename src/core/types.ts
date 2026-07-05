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
  | "gantt" // Gantt (Start/End rows), numeric or calendar timeline
  | "combo" // stacked columns + line series (Series.type === "line")
  | "pie" // pie chart (first series)
  | "doughnut" // doughnut variant of pie
  | "boxplot" // box-and-whisker: precomputed Min/Q1/Median/Q3/Max rows or raw samples
  | "radar" // spider chart: categories = spokes, series = polygons
  | "heatmap" // matrix: series = rows, categories = columns, value → color
  | "tilemap"; // tile-grid cartogram: categories = region codes, value → color

export interface Series {
  name: string;
  /** One value per category; null = blank cell. */
  values: (number | null)[];
  /** Override the palette color for this series. */
  color?: string;
  /**
   * Per-category fill overrides — highlight a single segment, point, or
   * slice ("color on a data point"). null cells keep the series color.
   */
  colors?: (string | null)[];
  /**
   * Hatch/dot pattern over the series fill — extends a tight color budget
   * and survives grayscale printing. Column-family segments; SVG/preview
   * renders the pattern, PowerPoint output degrades to the solid fill.
   */
  pattern?: "diagonal" | "crosshatch" | "dots" | "horizontal";
  /** Combo charts: render this series as a line over the columns. */
  type?: "column" | "line";
  /**
   * Stack group for clustered-stacked charts (think-cell separates stacks
   * with blank datasheet rows). Series sharing a stack index stack together;
   * different indices sit side by side within each category.
   */
  stack?: number;
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
  /**
   * Set when datasheet cells were parsed as calendar dates (values are days
   * since the Unix epoch). Gantt charts render a month/year timeline then.
   */
  dates?: boolean;
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
  /**
   * Numeric value axis on the left (off by default, as in think-cell).
   * "datamarks" draws Tufte-style tick dashes + labels with no axis line.
   */
  valueAxis: boolean | "datamarks";
  /**
   * Tick placement for datamark axes: "nice" round values (default) or
   * "data" — ticks at the actual data min/max (Tufte's range frame).
   */
  tickMode?: "nice" | "data";
  /** Radar: gridline web shape (default "polygon", business style). */
  gridShape?: "polygon" | "circle";
  /** Radar: series fill opacity (default 0.18; 0.25 for a single series). */
  fillOpacity?: number;
  gridlines: boolean;
  /**
   * Segment label content, think-cell's label dropdown: any combination of
   * value, percent (of column), series name, and category name.
   * Default: ["value"] (["percent"] on 100% charts).
   */
  labelContent?: ("value" | "percent" | "series" | "category")[];
  /** CAGR arrow between two categories; `series` computes it on one series. */
  cagr?: { from: number; to: number; series?: number };
  /**
   * Difference arrow between two category indices. Without `series` it is a
   * total difference arrow (column totals); with `series` it is a level
   * difference arrow comparing the cumulative level at that series.
   */
  difference?: {
    from: number;
    to: number;
    percent?: boolean;
    series?: number;
    /** Anchor the arrow's start at valueLines[fromValueLine] instead of a column. */
    fromValueLine?: number;
  };
  /** Value lines: at fixed values or at the mean of column totals. */
  valueLines?: ({ mode: "mean" } | { mode: "value"; value: number })[];
  /** @deprecated legacy single value line; normalized into valueLines. */
  valueLine?: { mode: "mean" } | { mode: "value"; value: number };
  /**
   * Connector lines joining segment boundaries of adjacent stacked
   * columns/bars — they make the development of each segment much easier
   * to follow. Stacked/100% charts (single stack group).
   */
  connectors?: boolean;
  /**
   * Speech-bubble callouts commenting on a value: anchored to a column top,
   * or to the cumulative level of `series` within it. dx/dy nudge the bubble
   * from its default spot above the anchor.
   */
  callouts?: { text: string; category: number; series?: number; dx?: number; dy?: number }[];
  /**
   * Shaded background bands highlighting a region of an axis, drawn behind
   * the data. axis "y" spans a value range; axis "x" spans category indices
   * (scatter/bubble: both axes are in value units).
   */
  bands?: { axis: "x" | "y"; from: number; to: number; color?: string; label?: string }[];
  /**
   * Render a "100% = N" note in the footnote line — the classic annotation
   * telling readers what the percentages are of. Pie/doughnut (series total)
   * and 100% charts (uniform denominator).
   */
  hundredPercentNote?: boolean;
}

export interface NumberFormat {
  /** Number of decimals; "auto" picks based on magnitude. */
  decimals: number | "auto";
  /** Append a suffix such as "%" or "€m". */
  suffix?: string;
  /** Show an explicit "+" on positive deltas (waterfall, difference arrows). */
  forceSign?: boolean;
  /** BCP-47 locale for separators, e.g. "de-DE" → 1.234,5. Default en-US. */
  locale?: string;
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
  /**
   * Sort categories by column total (think-cell's category sorting).
   * Column/Mekko/pie families only — order is meaningful for waterfall/Gantt.
   */
  categorySort?: "ascending" | "descending";
  /**
   * Combo charts: give line series their own right-hand value axis
   * (e.g. margin % over absolute columns).
   */
  secondaryAxis?: boolean;
  /**
   * think-cell axis break: compress the value range [from, to] into a small
   * band so out-of-scale columns fit. Vertical column/waterfall charts only.
   */
  axisBreak?: { from: number; to: number };
  /** Units label shown at the top of the value axis (e.g. "€m"). */
  valueAxisTitle?: string;
  /**
   * Manual label nudges by node name (e.g. {"label-0-2": {dx: 0, dy: -8}}) —
   * the pane-based stand-in for think-cell's label dragging. Applied to text
   * nodes after layout.
   */
  labelOffsets?: Record<string, { dx: number; dy: number }>;
  /** Logarithmic value axis (decade ticks). Clustered/line charts, positive data. */
  logScale?: boolean;
  /**
   * Source/footnote line rendered bottom-left in small muted text
   * (e.g. "Kilde: Danmarks Statistik, 2024"). Good charts always cite
   * their source and period.
   */
  footnote?: string;
  /** Pie/doughnut options: `explode` offsets the listed slice indices radially. */
  pie?: { explode?: number[] };
  /**
   * Boxplot options. Without precomputed Min/Q1/Median/Q3/Max rows, every
   * series row is a raw observation and quartiles/whiskers are computed.
   */
  boxplot?: {
    /** Whisker convention; auto: min/max for precomputed rows, Tukey for raw samples. */
    whiskers?: "tukey" | "minmax";
    /** Quartile method for raw samples (default "exclusive", Excel-style). */
    quartileMethod?: "exclusive" | "inclusive";
    /** Show the mean marker (default: only when a Mean row exists). */
    showMean?: boolean;
    /** Tukey fence multiplier (default 1.5). */
    iqrMultiplier?: number;
  };
  /** Tile-grid map layout; omitted → auto-detected from the region codes. */
  map?: "us" | "eu" | "europe" | "world";
  /** Heatmap color options; mode "auto" picks diverging when data spans zero. */
  heatmap?: { color?: string; negativeColor?: string; mode?: "sequential" | "diverging" | "auto" };
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
