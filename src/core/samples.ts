import type { ChartConfig, ChartKind } from "./types";
import { DEFAULT_SIZE } from "./chart";

/** Starter data per chart type, mirroring think-cell's insert defaults. */
export function sampleConfig(kind: ChartKind): ChartConfig {
  const base = { kind, ...DEFAULT_SIZE };
  switch (kind) {
    case "waterfall":
      return {
        ...base,
        title: "EBITDA bridge",
        data: {
          categories: ["FY23", "Volume", "Price", "Cost", "FX", "FY24"],
          series: [{ name: "Delta", values: [86, 14, 9, -12, -4, 0] }],
        },
        waterfall: { totalIndices: [5] },
        decorations: { categoryAxis: true },
      };
    case "mekko":
      return {
        ...base,
        title: "Revenue by region and segment",
        data: {
          categories: ["EMEA", "Americas", "APAC"],
          series: [
            { name: "Enterprise", values: [42, 55, 18] },
            { name: "SMB", values: [28, 30, 22] },
            { name: "Consumer", values: [14, 25, 30] },
          ],
        },
      };
    case "clustered":
      return {
        ...base,
        title: "Sales by quarter",
        data: {
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "2024", values: [42, 48, 45, 61] },
            { name: "2025", values: [51, 55, 58, 70] },
          ],
        },
        decorations: { seriesLabels: true },
      };
    case "stacked100":
      return {
        ...base,
        title: "Channel mix",
        data: {
          categories: ["2022", "2023", "2024", "2025"],
          series: [
            { name: "Retail", values: [55, 49, 44, 38] },
            { name: "Online", values: [30, 36, 41, 47] },
            { name: "Wholesale", values: [15, 15, 15, 15] },
          ],
        },
      };
    case "line":
      return {
        ...base,
        title: "Gross margin",
        data: {
          categories: ["2021", "2022", "2023", "2024", "2025"],
          series: [
            { name: "Brand A", values: [48, 52, 55, 61, 66] },
            { name: "Brand B", values: [40, 42, 41, 45, 49] },
          ],
        },
        decorations: { segmentLabels: false, valueAxis: true, gridlines: true },
      };
    case "butterfly":
      return {
        ...base,
        title: "Headcount by age band",
        data: {
          categories: ["<30", "30–39", "40–49", "50–59", "60+"],
          series: [
            { name: "Women", values: [420, 610, 540, 320, 120] },
            { name: "Men", values: [380, 650, 600, 410, 160] },
          ],
        },
      };
    case "scatter":
    case "bubble":
      return {
        ...base,
        title: kind === "bubble" ? "Market map" : "Price vs. margin",
        data: {
          categories: ["Alpha", "Bravo", "Core", "Delta", "Echo", "Foxtrot", "Golf", "Hotel"],
          series: [
            { name: "X", values: [12, 25, 40, 55, 62, 74, 30, 48] },
            { name: "Y", values: [30, 55, 42, 70, 35, 60, 22, 50] },
            ...(kind === "bubble" ? [{ name: "Size", values: [10, 40, 90, 25, 55, 70, 15, 35] as (number | null)[] }] : []),
            { name: "Group", values: [1, 1, 2, 2, 3, 3, 1, 2] },
          ],
        },
      };
    case "gantt":
      return {
        ...base,
        title: "Programme plan (weeks)",
        data: {
          categories: ["Scoping", "Design", "Build", "Test", "Rollout"],
          series: [
            { name: "Start", values: [1, 3, 5, 10, 13] },
            { name: "End", values: [3, 6, 11, 13, 15] },
            { name: "Milestone", values: [null, 6, null, 13, 15] },
          ],
        },
      };
    case "area":
      return {
        ...base,
        title: "Active users",
        data: {
          categories: ["2021", "2022", "2023", "2024", "2025"],
          series: [
            { name: "Mobile", values: [20, 32, 45, 60, 74] },
            { name: "Desktop", values: [35, 37, 36, 34, 30] },
          ],
        },
        decorations: { segmentLabels: false, valueAxis: true, gridlines: true },
      };
    default:
      return {
        ...base,
        kind: "stacked",
        title: "Revenue by segment",
        data: {
          categories: ["2022", "2023", "2024", "2025"],
          series: [
            { name: "Enterprise", values: [34, 40, 47, 56] },
            { name: "SMB", values: [22, 26, 31, 35] },
            { name: "Consumer", values: [18, 17, 19, 22] },
          ],
        },
        decorations: { totals: true, cagr: { from: 0, to: 3 } },
      };
  }
}

export const CHART_KINDS: { kind: ChartKind; label: string }[] = [
  { kind: "stacked", label: "Stacked" },
  { kind: "clustered", label: "Clustered" },
  { kind: "stacked100", label: "100%" },
  { kind: "waterfall", label: "Waterfall" },
  { kind: "mekko", label: "Mekko" },
  { kind: "line", label: "Line" },
  { kind: "area", label: "Area" },
  { kind: "butterfly", label: "Butterfly" },
  { kind: "scatter", label: "Scatter" },
  { kind: "bubble", label: "Bubble" },
  { kind: "gantt", label: "Gantt" },
];
