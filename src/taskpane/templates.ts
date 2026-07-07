import type { ChartConfig } from "../core/types";

/**
 * Curated starter charts — polished, decorated presets a user can drop in and
 * then edit, beyond the bare per-kind samples. All data is invented dummy data
 * (the repo is public); keep it that way.
 */
export const BUILTIN_TEMPLATES: { name: string; config: ChartConfig }[] = [
  {
    name: "Revenue bridge (waterfall)",
    config: {
      kind: "waterfall",
      width: 480,
      height: 300,
      title: "EBITDA bridge — FY24",
      data: {
        categories: ["FY23", "Volume", "Price", "Cost", "FX", "FY24"],
        series: [{ name: "Delta", values: [86, 14, 9, -12, -4, 0] }],
      },
      waterfall: { totalIndices: [5] },
      decorations: { categoryAxis: true, totals: true },
      footnote: "Source: management accounts",
    },
  },
  {
    name: "Growth columns + CAGR",
    config: {
      kind: "stacked",
      width: 480,
      height: 300,
      title: "Revenue by segment",
      data: {
        categories: ["2022", "2023", "2024", "2025"],
        series: [
          { name: "Enterprise", values: [34, 40, 47, 56] },
          { name: "SMB", values: [22, 26, 31, 35] },
          { name: "Consumer", values: [18, 17, 19, 22] },
        ],
      },
      decorations: { totals: true, cagr: { from: 0, to: 3 }, seriesLabels: true },
    },
  },
  {
    name: "Channel mix (100%)",
    config: {
      kind: "stacked100",
      width: 480,
      height: 300,
      title: "Channel mix",
      data: {
        categories: ["2022", "2023", "2024", "2025"],
        series: [
          { name: "Retail", values: [55, 49, 44, 38] },
          { name: "Online", values: [30, 36, 41, 47] },
          { name: "Wholesale", values: [15, 15, 15, 15] },
        ],
      },
      decorations: { segmentLabels: true, hundredPercentNote: true },
    },
  },
  {
    name: "Market share (pie)",
    config: {
      kind: "pie",
      width: 460,
      height: 300,
      title: "Revenue share — FY24",
      data: {
        categories: ["EMEA", "Americas", "APAC", "Other"],
        series: [{ name: "Revenue", values: [46, 28, 18, 8] }],
      },
      decorations: { hundredPercentNote: true },
      footnote: "Source: annual report",
    },
  },
  {
    name: "KPI trend (line)",
    config: {
      kind: "line",
      width: 480,
      height: 300,
      title: "Gross margin",
      data: {
        categories: ["2021", "2022", "2023", "2024", "2025"],
        series: [
          { name: "Brand A", values: [48, 52, 55, 61, 66] },
          { name: "Brand B", values: [40, 42, 41, 45, 49] },
        ],
      },
      decorations: { valueAxis: true, gridlines: true, valueLines: [{ mode: "mean" }] },
    },
  },
  {
    name: "Actual vs target (bullet)",
    config: {
      kind: "clustered",
      width: 480,
      height: 300,
      title: "Actual vs target",
      data: {
        categories: ["North", "South", "East", "West"],
        series: [
          { name: "Actual", values: [42, 55, 48, 61] },
          { name: "Target", values: [50, 50, 55, 58] },
        ],
      },
      decorations: {
        bands: [
          { axis: "y", from: 0, to: 35, color: "#e4e2da" },
          { axis: "y", from: 35, to: 55, color: "#efeee8" },
        ],
        valueAxis: true,
        seriesLabels: false,
      },
    },
  },
  {
    name: "Programme plan (Gantt)",
    config: {
      kind: "gantt",
      width: 520,
      height: 300,
      title: "Programme plan (weeks)",
      data: {
        categories: ["Scoping", "Design", "Build", "Test", "Rollout"],
        series: [
          { name: "Start", values: [1, 3, 5, 10, 13] },
          { name: "End", values: [3, 6, 11, 13, 15] },
          { name: "Milestone", values: [null, 6, null, 13, 15] },
        ],
      },
    },
  },
];
