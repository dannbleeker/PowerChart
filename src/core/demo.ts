import type { ChartConfig } from "./types";
import { estimateOfficeShapes, type Scene, type SceneNode } from "./scene";
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

/**
 * The opening title slide of the demo deck. Stamps the running build AND the host
 * (from Office.context.diagnostics) so a test PDF is fully self-identifying — you
 * can tell at a glance which build on which host produced it.
 */
function buildTitleScene(buildStamp: string, host: string): Scene {
  return {
    width: 840,
    height: 360,
    nodes: [
      {
        kind: "text",
        x: 0,
        y: 96,
        w: 840,
        h: 66,
        text: "SSF Charts gallery",
        fontSize: 40,
        bold: true,
        color: PALETTE[0],
        align: "left",
        valign: "top",
        name: "title",
      },
      {
        kind: "text",
        x: 0,
        y: 168,
        w: 840,
        h: 28,
        text: "Every chart kind, feature highlight, and scorecard element — as native, editable PowerPoint shapes.",
        fontSize: 15,
        color: "#52514e",
        align: "left",
        valign: "top",
        name: "subtitle",
      },
      {
        kind: "text",
        x: 0,
        y: 208,
        w: 840,
        h: 22,
        text: "The next slide indexes each chart with its shape count.",
        fontSize: 12,
        color: "#8a8984",
        align: "left",
        valign: "top",
        name: "note",
      },
      {
        kind: "text",
        x: 0,
        y: 296,
        w: 840,
        h: 22,
        text: `Build ${buildStamp}`,
        fontSize: 12,
        bold: true,
        color: "#52514e",
        align: "left",
        valign: "top",
        name: "build-stamp",
      },
      {
        kind: "text",
        x: 0,
        y: 320,
        w: 840,
        h: 22,
        text: `Host  ${host}`,
        fontSize: 12,
        color: "#8a8984",
        align: "left",
        valign: "top",
        name: "host",
      },
    ],
  };
}

/**
 * Shape ceiling for the harness's own slides — contents and results.
 *
 * NOT the ~90 web shape budget, which these pages already fitted under and
 * which turned out to be too generous for a slide made almost entirely of
 * text. On 2026-07-31 the full deck's contents page — 79 shapes, listing all
 * 37 charts, and explicitly exempted from the budget check so it could never
 * be skipped — was the second thing a shape-by-shape run drew, and PowerPoint
 * on the web crashed five seconds in. The twelve-item run's 27-shape contents
 * page, same path and session, rendered fine.
 *
 * 45 sits between those two observations, and below the ~40-shape charts that
 * have always drawn cleanly. Pages are MEASURED against it rather than
 * estimated from a row count, because a row's cost depends on what is in it.
 */
const HARNESS_PAGE_SHAPES = 45;

/** Upper bound on rows tried per page before measuring shrinks it. */
const INDEX_ROWS_PER_PAGE = 18;

/**
 * Largest page size, in items, whose built scene stays within the budget.
 *
 * Linear from the top down: the counts involved are small, the builder is
 * pure, and a measured answer cannot drift from the thing it measures the way
 * a hand-tuned row count did.
 */
function fitPageSize(max: number, build: (size: number) => Scene): number {
  for (let size = max; size > 1; size--) {
    if (estimateOfficeShapes(build(size)) <= HARNESS_PAGE_SHAPES) return size;
  }
  return 1;
}

/**
 * A contents table listing every chart with its shape count — doubles as the
 * regression manifest (shape count vs the ~90 web budget shows which are
 * skipped). Paginated: two name/count pairs side by side per page, so a small
 * deck fits one slide (as before) and a larger one spreads across N slides
 * rather than truncating.
 */
function buildIndexScenes(charts: DemoItem[]): Scene[] {
  if (charts.length === 0) return [];
  const page = (start: number, size: number): Scene => {
    const slice = charts.slice(start, start + size);
    const per = Math.ceil(slice.length / 2);
    const rows: string[][] = [["Chart", "Shapes", "Chart", "Shapes"]];
    for (let i = 0; i < per; i++) {
      const left = slice[i];
      const right = slice[i + per];
      rows.push([
        left ? `${start + i + 1}. ${left.title}` : "",
        left ? String(estimateOfficeShapes(left.scene)) : "",
        right ? `${start + i + per + 1}. ${right.title}` : "",
        right ? String(estimateOfficeShapes(right.scene)) : "",
      ]);
    }
    return buildTableScene(rows, 840);
  };
  const size = fitPageSize(INDEX_ROWS_PER_PAGE * 2, (n) => page(0, n));
  const pages: Scene[] = [];
  for (let start = 0; start < charts.length; start += size) pages.push(page(start, size));
  return pages;
}

/** Options for {@link demoItems}. All optional so `demoItems()` still works. */
export interface DemoOptions {
  /** Short build id stamped on the title slide, e.g. "10a6fa0 · 2026-07-17 20:11Z". */
  buildStamp?: string;
  /** Host descriptor (from `Office.context.diagnostics`) stamped under the build. */
  host?: string;
}

/**
 * The full demo deck: a title slide, a contents/manifest table, then one editable
 * chart per kind, the feature highlights, and the elements. A single button drops
 * all of this onto fresh slides for live testing in PowerPoint.
 */
export function demoItems(opts: DemoOptions = {}): DemoItem[] {
  const { buildStamp = "local build", host = "unknown host" } = opts;
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
  const indexPages = buildIndexScenes(charts);
  return [
    { title: "Title", scene: buildTitleScene(buildStamp, host) },
    ...indexPages.map((scene, i) => ({
      title: indexPages.length === 1 ? "Contents" : `Contents (page ${i + 1} of ${indexPages.length})`,
      scene,
    })),
    ...charts,
  ];
}

/** One item's outcome, as shown on the results slide. */
export interface ResultRow {
  title: string;
  status: "rendered" | "skipped" | "failed";
  shapes: number;
  ms: number;
}

/** Run-level totals shown on the results slide. `lost` = slides the host dropped. */
export interface ResultsSummary {
  buildStamp: string;
  items: number;
  rendered: number;
  skipped: number;
  failed: number;
  lost: number;
  /** Items that stalled but recovered on a retry; omitted/0 → not shown. */
  retried?: number;
  totalMs: number;
}

/** Upper bound on failure rows per results page before measuring shrinks it. */
const RESULTS_ROWS_PER_PAGE = 20;

/**
 * Paginate the results table so a failure-heavy run's own summary doesn't get
 * skipped as too dense — the exact bug in Presentation_2.pptx where 32 failures
 * pushed the scene to 135 shapes and the results slide came back as a stamped
 * placeholder. Returns ONE scene when the failures fit, otherwise Page N of M
 * scenes each with the run summary at top and a slice of the failure table.
 * A clean run still produces one page saying so.
 */
export function buildResultsScenes(rows: ResultRow[], summary: ResultsSummary): Scene[] {
  const failures = rows.filter((r) => r.status !== "rendered");
  if (failures.length && estimateOfficeShapes(buildResultsScene(rows, summary)) <= HARNESS_PAGE_SHAPES)
    return [buildResultsScene(rows, summary)];
  if (!failures.length) return [buildResultsScene(rows, summary)];
  const perPage = fitPageSize(Math.min(RESULTS_ROWS_PER_PAGE, failures.length), (n) =>
    buildResultsScene(failures.slice(0, n), summary, { page: 1, totalPages: 9 }),
  );
  const pages: Scene[] = [];
  const totalPages = Math.ceil(failures.length / perPage);
  for (let p = 0; p < totalPages; p++) {
    const slice = failures.slice(p * perPage, (p + 1) * perPage);
    // Reconstruct a ResultRow[] the single-page builder can consume — pass ONLY
    // the failures for this page (buildResultsScene filters by status).
    pages.push(buildResultsScene(slice, { ...summary }, { page: p + 1, totalPages }));
  }
  return pages;
}

/**
 * The closing results slide: a self-contained record of one regression run —
 * a summary line, the build stamp, and a table of ONLY the skipped/failed items
 * (chart · status · shapes · ms). Paginated by `buildResultsScenes` when the
 * failure count would push it past the web shape budget; a clean run still fits
 * one slide. The contents slide already lists every chart, so there's no need
 * to repeat the whole 37-row manifest here.
 */
export function buildResultsScene(
  rows: ResultRow[],
  summary: ResultsSummary,
  page?: { page: number; totalPages: number },
): Scene {
  const ink = "#52514e";
  const grey = "#8a8984";
  const secs = (summary.totalMs / 1000).toFixed(1);
  const recovered = summary.retried ?? 0;
  const summaryText = `${summary.items} items · ${summary.rendered} rendered · ${summary.skipped} skipped · ${summary.failed} failed · ${summary.lost} lost${recovered ? ` · ${recovered} recovered` : ""} · total ${secs}s`;
  const nodes: SceneNode[] = [
    {
      kind: "text",
      x: 0,
      y: 40,
      w: 840,
      h: 40,
      text: page ? `Regression results (page ${page.page} of ${page.totalPages})` : "Regression results",
      fontSize: 28,
      bold: true,
      color: PALETTE[0],
      align: "left",
      valign: "top",
      name: "results-title",
    },
    {
      kind: "text",
      x: 0,
      y: 96,
      w: 840,
      h: 24,
      text: summaryText,
      fontSize: 14,
      color: ink,
      align: "left",
      valign: "top",
      name: "results-summary",
    },
    {
      kind: "text",
      x: 0,
      y: 122,
      w: 840,
      h: 20,
      text: `Build ${summary.buildStamp}`,
      fontSize: 12,
      bold: true,
      color: ink,
      align: "left",
      valign: "top",
      name: "results-build",
    },
  ];
  const failures = rows.filter((r) => r.status !== "rendered");
  const headerY = 168;
  if (failures.length === 0) {
    nodes.push({
      kind: "text",
      x: 0,
      y: headerY,
      w: 840,
      h: 24,
      text: "All slides rendered cleanly.",
      fontSize: 14,
      color: ink,
      align: "left",
      valign: "top",
      name: "results-clean",
    });
  } else {
    nodes.push(
      {
        kind: "text",
        x: 0,
        y: headerY,
        w: 380,
        h: 18,
        text: "Chart",
        fontSize: 12,
        bold: true,
        color: ink,
        align: "left",
        valign: "top",
        name: "results-th-chart",
      },
      {
        kind: "text",
        x: 380,
        y: headerY,
        w: 160,
        h: 18,
        text: "Status",
        fontSize: 12,
        bold: true,
        color: ink,
        align: "left",
        valign: "top",
        name: "results-th-status",
      },
      {
        kind: "text",
        x: 560,
        y: headerY,
        w: 100,
        h: 18,
        text: "Shapes",
        fontSize: 12,
        bold: true,
        color: ink,
        align: "right",
        valign: "top",
        name: "results-th-shapes",
      },
      {
        kind: "text",
        x: 680,
        y: headerY,
        w: 140,
        h: 18,
        text: "ms",
        fontSize: 12,
        bold: true,
        color: ink,
        align: "right",
        valign: "top",
        name: "results-th-ms",
      },
    );
    failures.forEach((r, i) => {
      const y = headerY + 26 + i * 22;
      nodes.push(
        { kind: "text", x: 0, y, w: 380, h: 20, text: r.title, fontSize: 12, color: ink, align: "left", valign: "top" },
        {
          kind: "text",
          x: 380,
          y,
          w: 160,
          h: 20,
          text: r.status,
          fontSize: 12,
          color: r.status === "failed" ? "#c0392b" : grey,
          align: "left",
          valign: "top",
        },
        {
          kind: "text",
          x: 560,
          y,
          w: 100,
          h: 20,
          text: String(r.shapes),
          fontSize: 12,
          color: ink,
          align: "right",
          valign: "top",
        },
        {
          kind: "text",
          x: 680,
          y,
          w: 140,
          h: 20,
          text: String(r.ms),
          fontSize: 12,
          color: ink,
          align: "right",
          valign: "top",
        },
      );
    });
  }
  const height = Math.max(360, headerY + 26 + failures.length * 22 + 20);
  return { width: 840, height, nodes };
}
