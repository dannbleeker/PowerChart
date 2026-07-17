import type { ChartConfig } from "./types";
import type { Scene } from "./scene";
import { CHART_KINDS, sampleConfig } from "./samples";
import { buildChart, DEFAULT_SIZE } from "./chart";
import { buildAgendaScene } from "./agenda";
import { buildKpiTile, buildProcessFlow, buildTableScene } from "./elements";
import { PALETTE } from "./style";

/**
 * One entry per demo slide. `configJson` is set for real charts so the inserted
 * shape stays re-editable; static elements (agenda, KPI, …) leave it undefined.
 */
export interface DemoItem {
  title: string;
  scene: Scene;
  configJson?: string;
}

const F = { ...DEFAULT_SIZE };

/** Feature-highlight charts — the paths a plain per-kind sample doesn't show. */
function featureConfigs(): { title: string; config: ChartConfig }[] {
  return [
    {
      title: "Small multiples",
      config: {
        ...F,
        kind: "clustered",
        title: "Small multiples",
        multiples: {},
        data: {
          categories: ["Q1", "Q2", "Q3", "Q4"],
          series: [
            { name: "North", values: [12, 15, 14, 18] },
            { name: "South", values: [8, 9, 11, 10] },
            { name: "West", values: [5, 7, 6, 9] },
          ],
        },
      },
    },
    {
      title: "Error bars",
      config: {
        ...F,
        kind: "clustered",
        title: "Error bars",
        decorations: { segmentLabels: true },
        data: {
          categories: ["A", "B", "C"],
          series: [
            { name: "Value", values: [20, 28, 24] },
            { name: "Error", values: [3, 4, 2] },
          ],
        },
      },
    },
    {
      title: "Target markers",
      config: {
        ...F,
        kind: "clustered",
        title: "Target markers",
        data: {
          categories: ["A", "B", "C"],
          series: [
            { name: "Actual", values: [18, 24, 21] },
            { name: "Target", values: [20, 22, 25] },
          ],
        },
      },
    },
    {
      title: "Forecast split",
      config: {
        ...F,
        kind: "line",
        title: "Forecast split",
        decorations: { forecastFrom: 3, seriesLabels: true },
        data: {
          categories: ["Jan", "Feb", "Mar", "Apr", "May"],
          series: [{ name: "Revenue", values: [10, 12, 13, 15, 17] }],
        },
      },
    },
    {
      title: "Mean line & CAGR",
      config: {
        ...F,
        kind: "line",
        title: "Mean line & CAGR",
        decorations: { valueLines: [{ mode: "mean" }], cagr: { from: 0, to: 4 }, seriesLabels: true },
        data: {
          categories: ["FY20", "FY21", "FY22", "FY23", "FY24"],
          series: [{ name: "Users", values: [100, 130, 160, 190, 240] }],
        },
      },
    },
    {
      title: "Smoothed line",
      config: {
        ...F,
        kind: "line",
        title: "Smoothed line",
        decorations: { smooth: true },
        data: {
          categories: ["1", "2", "3", "4", "5", "6"],
          series: [{ name: "Signal", values: [3, 8, 5, 12, 7, 14] }],
        },
      },
    },
  ];
}

/** Static (non-chart) element slides — agenda and the scorecard primitives. */
function elementScenes(): { title: string; scene: Scene }[] {
  return [
    {
      title: "Agenda",
      scene: buildAgendaScene(["Overview", "Market", "Strategy", "Financials", "Next steps"], {
        highlight: 2,
        width: F.width,
        height: F.height,
      }),
    },
    { title: "KPI tile", scene: buildKpiTile({ label: "Revenue", value: "€4.2M", delta: "+12%" }) },
    { title: "Process flow", scene: buildProcessFlow(["Discover", "Design", "Build", "Ship"], 2) },
    {
      title: "Table",
      scene: buildTableScene([
        ["Region", "Q1", "Q2"],
        ["North", "12", "15"],
        ["South", "8", "9"],
      ]),
    },
  ];
}

/** The opening title slide of the demo deck. */
function buildTitleScene(): Scene {
  return {
    width: 840,
    height: 360,
    nodes: [
      { kind: "text", x: 0, y: 96, w: 840, h: 66, text: "PowerChart chart gallery", fontSize: 40, bold: true, color: PALETTE[0], align: "left", valign: "top", name: "title" },
      { kind: "text", x: 0, y: 168, w: 840, h: 28, text: "Every chart kind, feature highlight, and scorecard element — as native, editable PowerPoint shapes.", fontSize: 15, color: "#52514e", align: "left", valign: "top", name: "subtitle" },
      { kind: "text", x: 0, y: 208, w: 840, h: 22, text: "The next slide indexes each chart with its shape count.", fontSize: 12, color: "#8a8984", align: "left", valign: "top", name: "note" },
    ],
  };
}

/**
 * A contents table listing every chart with its shape count — doubles as the
 * regression manifest (shape count vs the ~90 web budget shows which are skipped).
 * Two name/count pairs side by side so all ~35 rows fit one slide, under budget.
 */
function buildIndexScene(charts: DemoItem[]): Scene {
  const per = Math.ceil(charts.length / 2);
  const rows: string[][] = [["Chart", "Shapes", "Chart", "Shapes"]];
  for (let i = 0; i < per; i++) {
    const left = charts[i];
    const right = charts[i + per];
    rows.push([
      left ? `${i + 1}. ${left.title}` : "",
      left ? String(left.scene.nodes.length) : "",
      right ? `${i + per + 1}. ${right.title}` : "",
      right ? String(right.scene.nodes.length) : "",
    ]);
  }
  return buildTableScene(rows, 840);
}

/**
 * The full demo deck: a title slide, a contents/manifest table, then one editable
 * chart per kind, the feature highlights, and the elements. A single button drops
 * all of this onto fresh slides for live testing in PowerPoint.
 */
export function demoItems(): DemoItem[] {
  const charts: DemoItem[] = [];
  for (const { kind, label } of CHART_KINDS) {
    const config: ChartConfig = { ...sampleConfig(kind), title: label };
    charts.push({ title: label, scene: buildChart(config), configJson: JSON.stringify(config) });
  }
  for (const { title, config } of featureConfigs()) {
    charts.push({ title, scene: buildChart(config), configJson: JSON.stringify(config) });
  }
  for (const { title, scene } of elementScenes()) {
    charts.push({ title, scene });
  }
  return [
    { title: "Title", scene: buildTitleScene() },
    { title: "Contents", scene: buildIndexScene(charts) },
    ...charts,
  ];
}
