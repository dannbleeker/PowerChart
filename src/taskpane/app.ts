import { buildChart, DEFAULT_SIZE, valueExtent } from "../core/chart";
import { PALETTES } from "../core/style";
import type { ChartConfig, ChartKind, Decorations } from "../core/types";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import { sceneToSvg } from "../render/svg";
import {
  getSelectionBounds,
  insertAgendaSlides,
  insertDemoDeck,
  insertSceneIntoSlide,
  isPowerPointHost,
  listChartsInDeck,
  listChartsInSelection,
  loadChartFromSelection,
  loadThemePalette,
  updateChartInSlide,
  updateChartsInSlides,
  onLateSync,
  errorText,
  type EditTarget,
  type InsertPhase,
} from "../render/powerpoint";
import { buildAgendaScene } from "../core/agenda";
import { demoItems } from "../core/demo";
import { buildCheckbox, buildHarveyBall, buildKpiTile, buildProcessFlow, buildTableScene, type CheckState } from "../core/elements";
import { localizePane, localizeTree } from "./i18n";
import { dataToSheet, mountDatasheet, sheetToData, type SheetModel } from "./datasheet";
import { BUILTIN_TEMPLATES } from "./templates";

interface AppState {
  kind: ChartKind;
  sheet: SheetModel;
  decorations: Partial<Decorations>;
  horizontal: boolean;
  title: string;
  segmentOrder: NonNullable<ChartConfig["segmentOrder"]>;
  scaleMin: string;
  scaleMax: string;
  breakFrom: string;
  breakTo: string;
  decimals: string; // "auto" | "0" | "1" | "2"
  suffix: string;
  locale: string;
  labelContent: string; // comma-joined parts, "" = default
  paletteName: string;
  /**
   * Style carried in from a loaded chart — fonts, negative/neutral, and a
   * palette matching no preset. Overrides the corporate style file's defaults;
   * `paletteName` still wins over `style.palette` once the user picks one.
   */
  style?: NonNullable<ChartConfig["style"]>;
  /** Per-series color overrides, keyed by series name. */
  seriesColors: Record<string, string>;
  axisTitle: string;
  logScale: boolean;
  /** Footnote / source line ("Kilde: …"). */
  footnote: string;
  /** Comma-separated slice indices to explode (pie/doughnut), 1-based in the UI. */
  pieExplode: string;
  /** Kind-specific config without pane controls, preserved across edits. */
  extras: Pick<ChartConfig, "boxplot" | "heatmap" | "map" | "combo" | "gapWidth" | "overlap">;
  /** When set, "Update chart" replaces this shape in place. */
  editTarget: EditTarget | null;
}

const state: AppState = { ...stateFromConfig(sampleConfig("stacked")), editTarget: null };

function stateFromConfig(cfg: ChartConfig): Omit<AppState, "editTarget"> {
  const sheet = dataToSheet(cfg.data);
  if (cfg.kind === "waterfall") {
    // Show "e" tokens in the sheet where totals are computed.
    for (const i of cfg.waterfall?.totalIndices ?? []) {
      if (sheet.cells[1]) sheet.cells[1][i + 1] = "e";
    }
  }
  return {
    kind: cfg.kind,
    sheet,
    decorations: { ...cfg.decorations },
    horizontal: !!cfg.horizontal,
    title: cfg.title ?? "",
    segmentOrder: cfg.segmentOrder ?? "sheet",
    scaleMin: cfg.scale?.min != null ? String(cfg.scale.min) : "",
    scaleMax: cfg.scale?.max != null ? String(cfg.scale.max) : "",
    breakFrom: cfg.axisBreak ? String(cfg.axisBreak.from) : "",
    breakTo: cfg.axisBreak ? String(cfg.axisBreak.to) : "",
    decimals: cfg.numberFormat?.decimals != null ? String(cfg.numberFormat.decimals) : "auto",
    suffix: cfg.numberFormat?.suffix ?? "",
    locale: cfg.numberFormat?.locale ?? "en-US",
    labelContent: cfg.decorations?.labelContent?.join(",") ?? "",
    paletteName: paletteNameFor(cfg.style?.palette),
    style: cfg.style ? { ...cfg.style } : undefined,
    seriesColors: Object.fromEntries(
      cfg.data.series.filter((s) => s.color).map((s) => [s.name, s.color!]),
    ),
    axisTitle: cfg.valueAxisTitle ?? "",
    logScale: !!cfg.logScale,
    footnote: cfg.footnote ?? "",
    pieExplode: (cfg.pie?.explode ?? []).map((i) => i + 1).join(","),
    extras: { boxplot: cfg.boxplot, heatmap: cfg.heatmap, map: cfg.map, combo: cfg.combo, gapWidth: cfg.gapWidth, overlap: cfg.overlap },
  };
}

/** Corporate style file: persisted defaults merged into every chart. */
interface StyleFile {
  palette?: string[];
  fontFamily?: string;
  fontSize?: number;
  negative?: string;
  neutral?: string;
}
let styleFile: StyleFile = {};
try {
  styleFile = JSON.parse(localStorage.getItem("powerchart-style") ?? "{}");
} catch {
  /* corrupted style file — start fresh */
}

/** Deck theme accents loaded via "Use deck theme" (session-scoped). */
let themePalette: string[] | null = null;

/** Style-file defaults + the palette preset chosen in the pane. */
function mergedStyle(): ChartConfig["style"] {
  // The loaded chart's own style beats the corporate defaults; an explicit
  // palette pick beats both.
  const style = { ...styleFile, ...state.style } as NonNullable<ChartConfig["style"]>;
  if (state.paletteName === "Theme" && themePalette) style.palette = themePalette;
  else if (state.paletteName !== "Default") style.palette = PALETTES[state.paletteName];
  return Object.keys(style).length ? style : undefined;
}

/**
 * The preset name for a palette, or "Default" when it matches none — including
 * a chart's own custom palette, which `state.style` carries instead.
 */
function paletteNameFor(palette?: string[]): string {
  if (!palette) return "Default";
  return Object.entries(PALETTES).find(([, p]) => p.join() === palette.join())?.[0] ?? "Default";
}

function currentConfig(): ChartConfig {
  const totals = new Set<number>();
  const data = sheetToData(state.sheet, state.kind === "waterfall" ? totals : undefined);
  const w = Number(($("chart-w") as HTMLInputElement | null)?.value);
  const h = Number(($("chart-h") as HTMLInputElement | null)?.value);
  const size = {
    width: Number.isFinite(w) && w >= 80 ? w : DEFAULT_SIZE.width,
    height: Number.isFinite(h) && h >= 60 ? h : DEFAULT_SIZE.height,
  };
  const min = state.scaleMin.trim() === "" ? undefined : Number(state.scaleMin);
  const max = state.scaleMax.trim() === "" ? undefined : Number(state.scaleMax);
  const bFrom = Number(state.breakFrom);
  const bTo = Number(state.breakTo);
  const axisBreak =
    state.breakFrom.trim() && state.breakTo.trim() && Number.isFinite(bFrom) && Number.isFinite(bTo) && bTo > bFrom
      ? { from: bFrom, to: bTo }
      : undefined;
  for (const s of data.series) {
    const c = state.seriesColors[s.name];
    if (c) s.color = c;
  }
  const labelParts = state.labelContent
    ? (state.labelContent.split(",") as NonNullable<Decorations["labelContent"]>)
    : undefined;
  const explode = state.pieExplode
    .split(",")
    .map((v) => Number(v.trim()) - 1)
    .filter((v) => Number.isInteger(v) && v >= 0);
  return {
    kind: state.kind,
    data,
    horizontal: state.horizontal || undefined,
    footnote: state.footnote || undefined,
    pie: explode.length ? { explode } : undefined,
    ...state.extras,
    valueAxisTitle: state.axisTitle || undefined,
    logScale: state.logScale || undefined,
    style: mergedStyle(),
    ...size,
    title: state.title || undefined,
    decorations: { ...state.decorations, labelContent: labelParts },
    waterfall: { totalIndices: [...totals] },
    segmentOrder: state.segmentOrder === "sheet" ? undefined : state.segmentOrder,
    axisBreak,
    scale:
      (min != null && Number.isFinite(min)) || (max != null && Number.isFinite(max))
        ? { min: Number.isFinite(min!) ? min : undefined, max: Number.isFinite(max!) ? max : undefined }
        : undefined,
    numberFormat:
      state.decimals !== "auto" || state.suffix || state.locale !== "en-US"
        ? {
            decimals: state.decimals === "auto" ? "auto" : Number(state.decimals),
            suffix: state.suffix || undefined,
            locale: state.locale !== "en-US" ? state.locale : undefined,
          }
        : undefined,
  };
}

// --- UI wiring ------------------------------------------------------------

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const gallery = $("gallery");
const preview = $("preview");
const optionsHost = $("options");
const hostNote = $("host-note");

let sheetApi: { setSheet(next: SheetModel): void };

/**
 * Pending auto-update push. Declared up here because the boot render calls
 * maybeAutoUpdate() long before the wiring below runs, and that now clears the
 * timer before its guard — a `let` further down would be in its dead zone.
 */
let autoUpdateTimer: ReturnType<typeof setTimeout> | undefined;

/**
 * Write the host note together with its status colour. The colour is a
 * parameter rather than an afterthought because only guard() used to set it,
 * so every other message inherited whatever the previous action left behind —
 * an "Invalid JSON" error rendered in the success green.
 */
function note(text: string, status: "ok" | "err" | "busy" | "none" = "none") {
  hostNote.textContent = text;
  hostNote.className = status === "none" ? "hint" : `hint status-${status}`;
}

function applyConfig(cfg: ChartConfig, editTarget: EditTarget | null) {
  Object.assign(state, stateFromConfig(cfg), { editTarget });
  sheetApi.setSheet(state.sheet);
  const titleField = document.getElementById("chart-title") as HTMLInputElement | null;
  if (titleField) titleField.value = state.title;
  // currentConfig() reads the size straight off these fields, so leaving them
  // stale silently resized every loaded chart to the previous one's dimensions.
  const sizeField = (id: string, value: number) => {
    const el = document.getElementById(id) as HTMLInputElement | null;
    if (el) el.value = String(value);
  };
  sizeField("chart-w", cfg.width ?? DEFAULT_SIZE.width);
  sizeField("chart-h", cfg.height ?? DEFAULT_SIZE.height);
  resetHistory();
  renderGallery();
  renderOptions();
  renderPreview();
  renderActionState();
}

/** Miniature preview of a chart kind for the gallery (think-cell's Elements menu). */
function thumbnailSvg(kind: ChartKind): string {
  const cfg: ChartConfig = {
    ...sampleConfig(kind),
    width: 96,
    height: 58,
    title: undefined,
    decorations: {
      segmentLabels: false,
      seriesLabels: false,
      totals: false,
      categoryAxis: false,
      valueAxis: false,
      gridlines: false,
    },
    style: { fontSize: 4 },
  };
  try {
    return sceneToSvg(buildChart(cfg));
  } catch {
    return "";
  }
}

const thumbnails = new Map<ChartKind, string>();

/** Chart kinds grouped by family, so the picker is scannable (think-cell's
 *  Elements menu). Any CHART_KINDS entry not listed here still renders under a
 *  trailing "Other" group, so a new kind can never silently disappear. */
const CHART_GROUPS: { label: string; kinds: ChartKind[] }[] = [
  { label: "Columns & bars", kinds: ["stacked", "clustered", "stacked100", "waterfall", "mekko", "butterfly", "cascade", "funnel"] },
  { label: "Line & area", kinds: ["line", "area", "combo"] },
  { label: "Parts of a whole", kinds: ["pie", "doughnut", "treemap", "sunburst", "waffle"] },
  { label: "Distribution", kinds: ["boxplot", "violin", "candlestick"] },
  { label: "Correlation", kinds: ["scatter", "bubble"] },
  { label: "Matrix & spatial", kinds: ["heatmap", "tilemap", "radar", "gantt"] },
];

function thumbButton(kind: ChartKind, label: string): HTMLButtonElement {
  const b = document.createElement("button");
  b.className = "thumb" + (kind === state.kind ? " active" : "");
  b.dataset.kind = kind;
  b.dataset.label = label.toLowerCase();
  if (!thumbnails.has(kind)) thumbnails.set(kind, thumbnailSvg(kind));
  const pic = document.createElement("span");
  pic.className = "thumb-pic";
  pic.innerHTML = thumbnails.get(kind)!;
  const cap = document.createElement("span");
  cap.className = "thumb-cap";
  cap.textContent = label;
  b.append(pic, cap);
  b.addEventListener("click", () => {
    applyConfig(sampleConfig(kind), null);
    // Auto-collapse the (tall) type grid once a kind is chosen — the summary
    // then shows the current kind, click to re-expand.
    const acc = document.getElementById("type-acc") as HTMLDetailsElement | null;
    if (acc) acc.open = false;
  });
  return b;
}

function renderGallery() {
  gallery.innerHTML = "";
  const labelOf = new Map(CHART_KINDS.map((k) => [k.kind, k.label] as const));
  const grouped = new Set<ChartKind>();
  const groups = CHART_GROUPS.map((g) => ({ label: g.label, kinds: g.kinds.filter((k) => labelOf.has(k)) }));
  groups.forEach((g) => g.kinds.forEach((k) => grouped.add(k)));
  const leftover = CHART_KINDS.filter((k) => !grouped.has(k.kind)).map((k) => k.kind);
  if (leftover.length) groups.push({ label: "Other", kinds: leftover });

  for (const g of groups) {
    if (!g.kinds.length) continue;
    const sec = document.createElement("div");
    sec.className = "type-group";
    const heading = document.createElement("div");
    heading.className = "group-label";
    heading.textContent = g.label;
    const grid = document.createElement("div");
    grid.className = "gallery";
    for (const kind of g.kinds) grid.appendChild(thumbButton(kind, labelOf.get(kind)!));
    sec.append(heading, grid);
    gallery.appendChild(sec);
  }
  localizeTree(gallery);
  updateTypeSummary();
  applyTypeFilter();
}

/** Filter the grouped picker by the search box; hide families with no match. */
function applyTypeFilter() {
  const input = document.getElementById("type-search-input") as HTMLInputElement | null;
  const q = (input?.value ?? "").trim().toLowerCase();
  let anyVisible = false;
  for (const sec of gallery.querySelectorAll<HTMLElement>(".type-group")) {
    let shown = 0;
    for (const btn of sec.querySelectorAll<HTMLButtonElement>(".thumb")) {
      const match = !q || (btn.dataset.label ?? "").includes(q);
      btn.style.display = match ? "" : "none";
      if (match) shown++;
    }
    sec.style.display = shown ? "" : "none";
    if (shown) anyVisible = true;
  }
  const noRes = document.getElementById("type-noresult");
  if (noRes) noRes.style.display = anyVisible ? "none" : "";
}

/** Reflect the selected chart kind in the collapsed "1 · Chart type" summary. */
function updateTypeSummary() {
  const sub = document.getElementById("type-sub");
  if (sub) sub.textContent = CHART_KINDS.find((k) => k.kind === state.kind)?.label ?? state.kind;
}

/** Top-level mode tabs: Chart / Elements / Agenda / Automation. */
function wireTabs() {
  const tabs = Array.from(document.querySelectorAll<HTMLButtonElement>(".tabs .tab"));
  const panels = Array.from(document.querySelectorAll<HTMLElement>(".tab-panel"));
  const bar = document.querySelector<HTMLElement>(".action-bar");
  /**
   * The action bar belongs to the Chart tab alone.
   *
   * Every one of its actions reads the CHART's state — "Insert into slide"
   * inserts currentConfig(), and the ⋯ menu edits, rescales or downloads
   * charts. On Elements it therefore offered a big primary button that inserts
   * a stacked column chart, sitting directly under the small "Insert" that
   * inserts the Harvey ball you are actually looking at: the prominent button
   * was the wrong one. Elements and Agenda already carry their own insert
   * buttons, so hiding it there removes a trap rather than a feature.
   */
  const showBarFor = (name?: string) => bar?.toggleAttribute("hidden", name !== "chart");
  for (const tab of tabs) {
    tab.addEventListener("click", () => {
      const name = tab.dataset.tab;
      tabs.forEach((t) => t.classList.toggle("active", t === tab));
      panels.forEach((p) => p.classList.toggle("active", p.dataset.panel === name));
      showBarFor(name);
    });
  }
  // No initial call: the bar ships visible and Chart ships active, so the
  // default is already right — and a ?tab= deep link CLICKS its tab, which
  // runs the listener above. Setting it here as well only looked prudent.
}

/** The footer "⋯" overflow menu holding the secondary actions (edit selected,
 *  same scale, download). Opens upward; closes on item click, outside click,
 *  or Escape. */
function wireActionsMenu() {
  const btn = document.getElementById("more-actions");
  const menu = document.getElementById("actions-menu");
  if (!btn || !menu) return;
  const setOpen = (open: boolean) => {
    menu.hidden = !open;
    btn.setAttribute("aria-expanded", String(open));
  };
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    setOpen(menu.hidden);
  });
  menu.addEventListener("click", () => setOpen(false));
  document.addEventListener("click", (e) => {
    if (!menu.hidden && e.target !== btn && !menu.contains(e.target as Node)) setOpen(false);
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !menu.hidden) setOpen(false);
  });
}

/** One collapsible Format group (Labels / Axes / Analysis / Layout / Colours). */
interface OptGroup {
  details: HTMLDetailsElement;
  togs: HTMLDivElement;
  body: HTMLDivElement;
}
const FGROUP_ICON: Record<string, string> = {
  labels: '<path d="M3 4h10M8 4v9" stroke-linecap="round"/>',
  axes: '<path d="M4 3v10h9M4 10l3-3 2 2 4-4" stroke-linecap="round" stroke-linejoin="round"/>',
  analysis: '<path d="M3 12l4-4 2 2 5-6M11 4h3v3" stroke-linecap="round" stroke-linejoin="round"/>',
  layout: '<rect x="3" y="3" width="10" height="10" rx="1.5"/><path d="M3 8h10M8 3v10" stroke-width="1.1"/>',
  colours: '<circle cx="8" cy="8" r="5"/><path d="M8 3a5 5 0 010 10z" fill="currentColor" stroke="none"/>',
};
function optGroup(name: string, iconKey: string): OptGroup {
  const details = document.createElement("details");
  details.className = "fgroup";
  const summary = document.createElement("summary");
  summary.innerHTML =
    `<svg class="fgroup-ico" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true">${FGROUP_ICON[iconKey]}</svg>` +
    `<span class="fgroup-name">${name}</span><span class="fgroup-count"></span><span class="fgroup-chev"></span>`;
  const body = document.createElement("div");
  body.className = "fgroup-body";
  const togs = document.createElement("div");
  togs.className = "togs";
  body.appendChild(togs);
  details.append(summary, body);
  return { details, togs, body };
}
/** Reflect each group's enabled-checkbox count in its "N on" pill. */
function updateGroupCounts() {
  for (const g of optionsHost.querySelectorAll<HTMLDetailsElement>(".fgroup")) {
    const on = g.querySelectorAll<HTMLInputElement>('input[type="checkbox"]:checked').length;
    const pill = g.querySelector(".fgroup-count");
    if (pill) {
      pill.textContent = on ? `${on} on` : "";
      pill.classList.toggle("zero", on === 0);
    }
  }
}

function renderOptions() {
  optionsHost.innerHTML = "";
  const d = state.decorations;
  const nCats = () => Math.max(0, state.sheet.cells[0].length - 1);
  // think-cell surfaces these controls contextually on the chart; in the pane
  // they're grouped so the long list stays scannable.
  const G = {
    labels: optGroup("Labels", "labels"),
    axes: optGroup("Axes & scale", "axes"),
    analysis: optGroup("Analysis", "analysis"),
    layout: optGroup("Layout", "layout"),
    colours: optGroup("Colours & style", "colours"),
  };
  G.labels.details.open = true;

  const toggles: { key: keyof Decorations; label: string; group: OptGroup }[] = [
    { key: "segmentLabels", label: "Segment labels", group: G.labels },
    { key: "seriesLabels", label: "Series labels", group: G.labels },
    { key: "totals", label: "Column totals", group: G.labels },
    { key: "categoryAxis", label: "Category labels", group: G.labels },
    { key: "valueAxis", label: "Value axis", group: G.axes },
    { key: "gridlines", label: "Gridlines", group: G.axes },
    { key: "connectors", label: "Connector lines", group: G.layout },
    { key: "hundredPercentNote", label: "100% = note", group: G.labels },
  ];
  for (const t of toggles) {
    const label = document.createElement("label");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!d[t.key];
    cb.addEventListener("change", () => {
      (d as Record<string, unknown>)[t.key] = cb.checked;
      renderPreview();
    });
    label.append(cb, t.label);
    t.group.togs.appendChild(label);
  }

  // Tufte-style datamark axis: tick dashes + labels, no axis line.
  const dm = document.createElement("label");
  const dmCb = document.createElement("input");
  dmCb.type = "checkbox";
  dmCb.checked = d.valueAxis === "datamarks";
  dmCb.addEventListener("change", () => {
    d.valueAxis = dmCb.checked ? "datamarks" : false;
    renderOptions();
    renderPreview();
  });
  dm.append(dmCb, "Datamark axis (ticks only)");
  G.axes.togs.appendChild(dm);

  // think-cell's rotation handle, as a toggle: column ⇄ bar.
  const rot = document.createElement("label");
  const rotCb = document.createElement("input");
  rotCb.type = "checkbox";
  rotCb.checked = state.horizontal;
  rotCb.addEventListener("change", () => {
    state.horizontal = rotCb.checked;
    renderPreview();
  });
  rot.append(rotCb, "Horizontal (bar)");
  G.layout.togs.appendChild(rot);

  G.analysis.body.appendChild(
    pairControl("CAGR arrow", d.cagr, nCats(), (pair) => {
      d.cagr = pair;
      renderPreview();
    }),
  );

  // Difference arrow: totals by default, or a level arrow at a series.
  const diff = document.createElement("label");
  diff.className = "wide";
  const diffCb = document.createElement("input");
  diffCb.type = "checkbox";
  diffCb.checked = !!d.difference;
  const dFrom = numInput((d.difference?.from ?? 0) + 1);
  const dTo = numInput((d.difference?.to ?? Math.max(0, nCats() - 1)) + 1);
  const dSeries = numInput((d.difference?.series ?? -1) + 1, 0);
  dSeries.title = "0 = column totals, 1+ = level of that series";
  const emitDiff = () => {
    const s = Number(dSeries.value) - 1;
    d.difference = diffCb.checked
      ? { from: Number(dFrom.value) - 1, to: Number(dTo.value) - 1, series: s >= 0 ? s : undefined }
      : undefined;
    renderPreview();
  };
  [diffCb, dFrom, dTo, dSeries].forEach((el) => el.addEventListener(el === diffCb ? "change" : "input", emitDiff));
  diff.append(diffCb, "Difference arrow from ", dFrom, " to ", dTo, " series ", dSeries);
  G.analysis.body.appendChild(diff);

  // Value lines: mean and/or comma-separated fixed values.
  const vl = document.createElement("label");
  vl.className = "wide";
  const vlMean = document.createElement("input");
  vlMean.type = "checkbox";
  const existing = d.valueLines ?? (d.valueLine ? [d.valueLine] : []);
  vlMean.checked = existing.some((v) => v.mode === "mean");
  const vlValues = document.createElement("input");
  vlValues.type = "text";
  vlValues.placeholder = "e.g. 50, 100";
  vlValues.style.width = "80px";
  vlValues.value = existing
    .filter((v): v is { mode: "value"; value: number } => v.mode === "value")
    .map((v) => v.value)
    .join(", ");
  const emitVl = () => {
    const lines: NonNullable<Decorations["valueLines"]> = [];
    if (vlMean.checked) lines.push({ mode: "mean" });
    for (const part of vlValues.value.split(",")) {
      const v = Number(part.trim());
      if (part.trim() && Number.isFinite(v)) lines.push({ mode: "value", value: v });
    }
    d.valueLines = lines.length ? lines : undefined;
    d.valueLine = undefined;
    renderPreview();
  };
  vlMean.addEventListener("change", emitVl);
  vlValues.addEventListener("input", emitVl);
  vl.append(vlMean, "Value line: mean Ø", " + values ", vlValues);
  G.analysis.body.appendChild(vl);

  // Segment order (think-cell's mini-toolbar menu).
  const so = document.createElement("label");
  so.className = "wide";
  const soSel = document.createElement("select");
  for (const [value, label] of [
    ["sheet", "Sheet order"],
    ["reverse", "Reversed"],
    ["ascending", "Ascending"],
    ["descending", "Descending"],
  ]) {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = label;
    soSel.appendChild(opt);
  }
  soSel.value = state.segmentOrder;
  soSel.addEventListener("change", () => {
    state.segmentOrder = soSel.value as AppState["segmentOrder"];
    renderPreview();
  });
  so.append("Segment order ", soSel);
  G.layout.body.appendChild(so);

  // Manual axis scale (think-cell's axis-handle dragging).
  const sc = document.createElement("label");
  sc.className = "wide";
  const scMin = document.createElement("input");
  const scMax = document.createElement("input");
  for (const [el, val] of [[scMin, state.scaleMin], [scMax, state.scaleMax]] as const) {
    el.type = "text";
    el.style.width = "48px";
    el.placeholder = "auto";
    el.value = val;
  }
  const emitScale = () => {
    state.scaleMin = scMin.value;
    state.scaleMax = scMax.value;
    renderPreview();
  };
  scMin.addEventListener("input", emitScale);
  scMax.addEventListener("input", emitScale);
  sc.append("Axis scale min ", scMin, " max ", scMax);
  G.axes.body.appendChild(sc);

  // Axis break (compresses the given value range).
  const ab = document.createElement("label");
  ab.className = "wide";
  const abFrom = document.createElement("input");
  const abTo = document.createElement("input");
  for (const [el, val] of [[abFrom, state.breakFrom], [abTo, state.breakTo]] as const) {
    el.type = "text";
    el.style.width = "48px";
    el.placeholder = "none";
    el.value = val;
  }
  const emitBreak = () => {
    state.breakFrom = abFrom.value;
    state.breakTo = abTo.value;
    renderPreview();
  };
  abFrom.addEventListener("input", emitBreak);
  abTo.addEventListener("input", emitBreak);
  ab.append("Axis break from ", abFrom, " to ", abTo);
  G.axes.body.appendChild(ab);

  // Number format.
  const nf = document.createElement("label");
  nf.className = "wide";
  const nfDec = document.createElement("select");
  for (const v of ["auto", "0", "1", "2"]) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v === "auto" ? "auto" : `${v} dp`;
    nfDec.appendChild(opt);
  }
  nfDec.value = state.decimals;
  const nfSuffix = document.createElement("input");
  nfSuffix.type = "text";
  nfSuffix.style.width = "48px";
  nfSuffix.placeholder = "e.g. €m";
  nfSuffix.value = state.suffix;
  const emitNf = () => {
    state.decimals = nfDec.value;
    state.suffix = nfSuffix.value;
    renderPreview();
  };
  nfDec.addEventListener("change", emitNf);
  nfSuffix.addEventListener("input", emitNf);
  const nfLoc = document.createElement("select");
  for (const v of ["en-US", "de-DE", "fr-FR", "da-DK"]) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    nfLoc.appendChild(opt);
  }
  nfLoc.value = state.locale;
  nfLoc.addEventListener("change", () => {
    state.locale = nfLoc.value;
    renderPreview();
  });
  nf.append("Labels: decimals ", nfDec, " suffix ", nfSuffix, " locale ", nfLoc);
  G.labels.body.appendChild(nf);

  // Footnote / source line — good charts always cite their source.
  const fn = document.createElement("label");
  fn.className = "wide";
  const fnInput = document.createElement("input");
  fnInput.type = "text";
  fnInput.placeholder = "e.g. Source: Statistics Denmark, 2024";
  fnInput.style.width = "180px";
  fnInput.value = state.footnote;
  fnInput.addEventListener("input", () => {
    state.footnote = fnInput.value;
    renderPreview();
  });
  fn.append("Footnote / source ", fnInput);
  G.colours.body.appendChild(fn);

  // Exploding slices (pie/doughnut only).
  if (state.kind === "pie" || state.kind === "doughnut") {
    const ex = document.createElement("label");
    ex.className = "wide";
    const exInput = document.createElement("input");
    exInput.type = "text";
    exInput.placeholder = "e.g. 1";
    exInput.style.width = "48px";
    exInput.value = state.pieExplode;
    exInput.addEventListener("input", () => {
      state.pieExplode = exInput.value;
      renderPreview();
    });
    ex.append("Explode slices ", exInput);
    G.layout.body.appendChild(ex);
  }

  // Label content (think-cell's label dropdown).
  const lc = document.createElement("label");
  lc.className = "wide";
  const lcSel = document.createElement("select");
  for (const [value, label] of [
    ["", "Default"],
    ["value", "Value"],
    ["percent", "%"],
    ["value,percent", "Value + %"],
    ["series,value", "Series + value"],
    ["category,percent", "Category + %"],
  ]) {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = label;
    lcSel.appendChild(opt);
  }
  lcSel.value = state.labelContent;
  lcSel.addEventListener("change", () => {
    state.labelContent = lcSel.value;
    renderPreview();
  });
  lc.append("Label content ", lcSel);
  G.labels.body.appendChild(lc);

  // Axis title + log scale.
  const ax = document.createElement("label");
  ax.className = "wide";
  const axTitle = document.createElement("input");
  axTitle.type = "text";
  axTitle.style.width = "56px";
  axTitle.placeholder = "e.g. €m";
  axTitle.value = state.axisTitle;
  axTitle.addEventListener("input", () => {
    state.axisTitle = axTitle.value;
    renderPreview();
  });
  const axLog = document.createElement("input");
  axLog.type = "checkbox";
  axLog.checked = state.logScale;
  axLog.addEventListener("change", () => {
    state.logScale = axLog.checked;
    renderPreview();
  });
  ax.append("Axis title ", axTitle, " ", axLog, " log scale");
  G.axes.body.appendChild(ax);

  // Palette preset + per-series color overrides.
  const pal = document.createElement("label");
  pal.className = "wide";
  const palSel = document.createElement("select");
  const palNames = [...Object.keys(PALETTES), ...(themePalette ? ["Theme"] : [])];
  for (const name of palNames) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    palSel.appendChild(opt);
  }
  palSel.value = state.paletteName;
  palSel.addEventListener("change", () => {
    state.paletteName = palSel.value;
    // An explicit pick — "Default" included — replaces a custom palette the
    // loaded chart brought with it.
    delete state.style?.palette;
    renderPreview();
  });
  // Read the deck's theme accent colors (PowerPointApi 1.10) as a palette.
  const themeBtn = document.createElement("button");
  themeBtn.type = "button";
  themeBtn.textContent = "Use deck theme";
  themeBtn.disabled = !isPowerPointHost();
  themeBtn.addEventListener("click", async () => {
    const loaded = await loadThemePalette();
    if (!loaded) {
      themeBtn.textContent = "Theme unavailable";
      return;
    }
    themePalette = loaded;
    if (![...palSel.options].some((o) => o.value === "Theme")) {
      const opt = document.createElement("option");
      opt.value = "Theme";
      opt.textContent = "Theme";
      palSel.appendChild(opt);
    }
    state.paletteName = "Theme";
    palSel.value = "Theme";
    renderOptions();
    renderPreview();
  });
  pal.append("Palette ", palSel, " ", themeBtn);
  G.colours.body.appendChild(pal);

  const colors = document.createElement("div");
  colors.className = "wide series-colors";
  // Resolve through mergedStyle so the swatches show the colors the chart
  // actually draws with — including a loaded chart's custom palette.
  const palette = mergedStyle()?.palette ?? PALETTES.Default;
  currentSeriesNames().forEach((name, i) => {
    const wrap = document.createElement("label");
    const input = document.createElement("input");
    input.type = "color";
    input.value = state.seriesColors[name] ?? palette[i % palette.length];
    input.addEventListener("input", () => {
      state.seriesColors[name] = input.value;
      renderPreview();
    });
    wrap.append(input, name);
    colors.appendChild(wrap);
  });
  G.colours.body.appendChild(colors);

  for (const g of [G.labels, G.axes, G.analysis, G.layout, G.colours]) optionsHost.appendChild(g.details);
  localizeTree(optionsHost);
  updateGroupCounts();
}

/** Series names from the sheet, excluding special rows. */
function currentSeriesNames(): string[] {
  return sheetToData(state.sheet).series.map((s) => s.name);
}

function numInput(value: number, min = 1): HTMLInputElement {
  const el = document.createElement("input");
  el.type = "number";
  el.min = String(min);
  el.value = String(value);
  return el;
}

function pairControl(
  label: string,
  current: { from: number; to: number } | undefined,
  nCats: number,
  onChange: (pair: { from: number; to: number } | undefined) => void,
): HTMLElement {
  const wrap = document.createElement("label");
  wrap.className = "wide";
  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.checked = !!current;
  const from = numInput((current?.from ?? 0) + 1);
  const to = numInput((current?.to ?? Math.max(0, nCats - 1)) + 1);
  const emit = () =>
    onChange(cb.checked ? { from: Number(from.value) - 1, to: Number(to.value) - 1 } : undefined);
  cb.addEventListener("change", emit);
  from.addEventListener("input", emit);
  to.addEventListener("input", emit);
  wrap.append(cb, label, " from ", from, " to ", to);
  return wrap;
}

function renderPreview() {
  try {
    const scene = buildChart(currentConfig());
    preview.innerHTML = sceneToSvg(scene, { background: "#ffffff" });
  } catch (err) {
    preview.innerHTML = `<p class="hint">Could not render: ${err instanceof Error ? err.message : String(err)}</p>`;
  }
  maybeAutoUpdate();
}

function renderActionState() {
  const insertBtn = $("insert") as HTMLButtonElement;
  insertBtn.textContent = state.editTarget ? "Update chart" : "Insert into slide";
  ($("insert-new") as HTMLButtonElement).style.display = state.editTarget ? "" : "none";
  // This label is rewritten every time the edit target changes — long after
  // localizePane translated the pane — so it has to be re-translated or it
  // reverts to English. localizeTree only looks at descendants, hence the
  // parent.
  if (insertBtn.parentElement) localizeTree(insertBtn.parentElement);
}

// --- boot ------------------------------------------------------------------

// Datasheet undo/redo (Ctrl+Z / Ctrl+Y while the pane has focus).
const history: string[] = [];
const redoStack: string[] = [];
function snapshot() {
  const snap = JSON.stringify(state.sheet.cells);
  if (history[history.length - 1] !== snap) {
    history.push(snap);
    if (history.length > 100) history.shift();
    redoStack.length = 0;
  }
}
/**
 * Start a fresh undo timeline at the current sheet — the same baseline boot
 * establishes, re-established for each newly loaded chart. Without it Ctrl+Z
 * replayed the previous chart's cells into the new one.
 */
function resetHistory() {
  history.length = 0;
  redoStack.length = 0;
  history.push(JSON.stringify(state.sheet.cells));
}
function restore(cells: string[][]) {
  state.sheet = { cells };
  sheetApi.setSheet(state.sheet);
  renderPreview();
}
document.addEventListener("keydown", (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  if (e.key === "z" && history.length > 1) {
    e.preventDefault();
    redoStack.push(history.pop()!);
    restore(JSON.parse(history[history.length - 1]));
  } else if (e.key === "y" && redoStack.length) {
    e.preventDefault();
    const snap = redoStack.pop()!;
    history.push(snap);
    restore(JSON.parse(snap));
  }
});

sheetApi = mountDatasheet($("datasheet"), state.sheet, (sheet) => {
  state.sheet = sheet;
  snapshot();
  renderPreview();
});
snapshot();

const titleInput = $("chart-title") as HTMLInputElement;
titleInput.value = state.title;
titleInput.addEventListener("input", () => {
  state.title = titleInput.value;
  renderPreview();
});
wireTabs();
wireActionsMenu();
document.getElementById("type-search-input")?.addEventListener("input", applyTypeFilter);
optionsHost.addEventListener("change", updateGroupCounts);
renderGallery();
renderOptions();
renderPreview();
renderActionState();

$("download").addEventListener("click", () => {
  const svg = sceneToSvg(buildChart(currentConfig()), { background: "#ffffff" });
  const blob = new Blob([svg], { type: "image/svg+xml" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "powerchart.svg";
  a.click();
  URL.revokeObjectURL(a.href);
});

/** Cascading default insert position so repeated inserts don't pile up. */
let insertOffset = 0;

/**
 * Say which host phase we are in. A stalled Office.js sync never throws — it
 * simply never settles — so "Working…" alone cannot tell a slow host from a
 * dead one. Naming the phase makes a stall legible: whatever it says last is
 * where it stopped.
 */
function phaseNote(phase: InsertPhase, detail?: string) {
  const said: Record<InsertPhase, string> = {
    context: "opening PowerPoint…",
    queue: "building shapes…",
    commit: "sending to PowerPoint…",
    group: "grouping…",
    done: "done",
  };
  note(`Working… ${said[phase]}${detail ? ` (${detail})` : ""}`, "busy");
}

async function doInsert(asNew: boolean) {
  let cfg = currentConfig();
  if (!asNew && state.editTarget) {
    const scene = buildChart(cfg);
    await updateChartInSlide(scene, state.editTarget, { tagData: JSON.stringify(cfg) });
    return;
  }
  // New chart: use the selected placeholder's bounds when one is selected.
  const bounds = await getSelectionBounds();
  if (bounds && bounds.width > 40 && bounds.height > 40) {
    cfg = { ...cfg, width: bounds.width, height: bounds.height };
    await insertSceneIntoSlide(
      buildChart(cfg),
      { tagData: JSON.stringify(cfg), left: bounds.left, top: bounds.top },
      phaseNote,
    );
  } else {
    await insertSceneIntoSlide(
      buildChart(cfg),
      { tagData: JSON.stringify(cfg), left: 60 + insertOffset, top: 90 + insertOffset },
      phaseNote,
    );
    insertOffset = (insertOffset + 14) % 84;
  }
  state.editTarget = null;
  renderActionState();
}

/**
 * think-cell's Set Same Scale: pin every value-axis chart (in the deck, or
 * just the selected ones) to the union of their extents and re-render them.
 */
async function doSameScale(scope: "deck" | "selection" = "deck") {
  const charts = scope === "deck" ? await listChartsInDeck() : await listChartsInSelection();
  const parsed = charts
    .map((c) => ({ target: c.target, cfg: JSON.parse(c.configJson) as ChartConfig }))
    .map((c) => ({ ...c, extent: valueExtent(c.cfg) }))
    .filter((c): c is typeof c & { extent: { min: number; max: number } } => c.extent != null);
  if (parsed.length < 2) {
    note(
      scope === "deck"
        ? "Same scale needs at least two value-axis charts in the deck."
        : "Select two or more PowerCharts (Ctrl-click), then apply Same scale.",
      "err",
    );
    return;
  }
  const min = Math.min(...parsed.map((c) => c.extent.min));
  const max = Math.max(...parsed.map((c) => c.extent.max));
  // One request context for the whole deck, not one per chart: each chart's
  // update costs four round-trips to PowerPoint, and awaiting them in a loop
  // made Same Scale across 20 charts eighty of them.
  await updateChartsInSlides(
    parsed.map((c) => {
      c.cfg.scale = { min: min < 0 ? min : undefined, max };
      return { scene: buildChart(c.cfg), target: c.target, opts: { tagData: JSON.stringify(c.cfg) } };
    }),
  );
  note(`Same scale applied to ${parsed.length} charts (max ${max}).`, "ok");
}

async function doLoadSelection() {
  const found = await loadChartFromSelection();
  if (!found) {
    note("The selection is not a PowerChart — select an inserted chart group first.", "err");
    return;
  }
  note("Chart loaded — edits will update it in place.", "ok");
  applyConfig(JSON.parse(found.configJson) as ChartConfig, found.target);
  // The banner offers to load the selected chart; once it is loaded the offer
  // is stale. Hiding it here rather than in the banner's own click handler
  // covers the other way in — the "Edit selected chart" button.
  $("selection-banner").style.display = "none";
}

// --- Elements (harvey balls, checkboxes, process flow, table) -----------------

function harveyScene() {
  return buildHarveyBall(Number(($("harvey-pct") as HTMLInputElement).value) / 100, 24);
}
function checkScene() {
  return buildCheckbox(($("check-state") as HTMLSelectElement).value as CheckState, 20);
}
function flowScene() {
  const steps = ($("flow-steps") as HTMLInputElement).value.split(",").map((s) => s.trim()).filter(Boolean);
  const hl = Number(($("flow-highlight") as HTMLInputElement).value) - 1;
  return buildProcessFlow(steps, hl, 480, 40);
}
function kpiScene() {
  return buildKpiTile({
    label: ($("kpi-label") as HTMLInputElement).value,
    value: ($("kpi-value") as HTMLInputElement).value,
    delta: ($("kpi-delta") as HTMLInputElement).value || undefined,
    goodIsUp: !($("kpi-down-good") as HTMLInputElement).checked,
  });
}
function renderElementPreviews() {
  $("harvey-val").textContent = `${($("harvey-pct") as HTMLInputElement).value}%`;
  $("harvey-preview").innerHTML = sceneToSvg(harveyScene());
  $("check-preview").innerHTML = sceneToSvg(checkScene());
  $("flow-preview").innerHTML = sceneToSvg(flowScene());
  $("kpi-preview").innerHTML = sceneToSvg(kpiScene());
}
for (const id of ["harvey-pct", "check-state", "flow-steps", "flow-highlight", "kpi-label", "kpi-value", "kpi-delta", "kpi-down-good"]) {
  $(id).addEventListener("input", renderElementPreviews);
}
renderElementPreviews();

// --- Templates & style file ----------------------------------------------------

const TEMPLATES_KEY = "powerchart-templates";
const STYLE_KEY = "powerchart-style";

function loadTemplates(): Record<string, ChartConfig> {
  try {
    return JSON.parse(localStorage.getItem(TEMPLATES_KEY) ?? "{}");
  } catch {
    return {};
  }
}

function renderTemplateList() {
  const sel = $("template-list") as HTMLSelectElement;
  sel.innerHTML = "<option value=''>— templates —</option>";
  const starters = document.createElement("optgroup");
  starters.label = "Starters";
  for (const t of BUILTIN_TEMPLATES) {
    const opt = document.createElement("option");
    opt.value = `builtin:${t.name}`;
    opt.textContent = t.name;
    starters.appendChild(opt);
  }
  sel.appendChild(starters);
  const names = Object.keys(loadTemplates()).sort();
  if (names.length) {
    const mine = document.createElement("optgroup");
    mine.label = "My templates";
    for (const name of names) {
      const opt = document.createElement("option");
      opt.value = `user:${name}`;
      opt.textContent = name;
      mine.appendChild(opt);
    }
    sel.appendChild(mine);
  }
}

$("template-save").addEventListener("click", () => {
  const name = prompt("Template name?", state.title || state.kind);
  if (!name) return;
  const all = loadTemplates();
  all[name] = currentConfig();
  localStorage.setItem(TEMPLATES_KEY, JSON.stringify(all));
  renderTemplateList();
});
$("template-list").addEventListener("change", () => {
  const value = ($("template-list") as HTMLSelectElement).value;
  if (!value) return;
  const sep = value.indexOf(":");
  const [source, name] = [value.slice(0, sep), value.slice(sep + 1)];
  const cfg =
    source === "builtin"
      ? BUILTIN_TEMPLATES.find((t) => t.name === name)?.config
      : loadTemplates()[name];
  if (cfg) applyConfig({ ...DEFAULT_SIZE, ...cfg }, null);
});
$("template-delete").addEventListener("click", () => {
  const value = ($("template-list") as HTMLSelectElement).value;
  if (!value.startsWith("user:")) return; // starters (and the placeholder) can't be deleted
  const name = value.slice("user:".length);
  const all = loadTemplates();
  delete all[name];
  localStorage.setItem(TEMPLATES_KEY, JSON.stringify(all));
  renderTemplateList();
});
renderTemplateList();

$("style-export").addEventListener("click", () => {
  const current: StyleFile = { ...styleFile };
  if (state.paletteName === "Theme" && themePalette) current.palette = themePalette;
  else if (state.paletteName !== "Default") current.palette = PALETTES[state.paletteName];
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(current, null, 2);
  note("Style exported — share the JSON as your corporate style file.", "ok");
});
$("style-import").addEventListener("click", () => {
  try {
    const parsed = JSON.parse(($("json-io") as HTMLTextAreaElement).value);
    if (parsed.kind) throw new Error("that is a chart config — use Import instead");
    styleFile = parsed;
    localStorage.setItem(STYLE_KEY, JSON.stringify(styleFile));
    renderPreview();
    note("Style imported — applied to every chart from this pane.", "ok");
  } catch (err) {
    note(`Style import failed: ${err instanceof Error ? err.message : String(err)}`, "err");
  }
});

// --- Automation (JSON in / out, the open .ppttc idea) -------------------------

$("json-export").addEventListener("click", () => {
  ($("json-io") as HTMLTextAreaElement).value = JSON.stringify(currentConfig(), null, 2);
});
$("json-import").addEventListener("click", () => {
  try {
    const parsed = JSON.parse(($("json-io") as HTMLTextAreaElement).value);
    applyConfig({ ...DEFAULT_SIZE, ...(Array.isArray(parsed) ? parsed[0] : parsed) }, null);
    note(
      Array.isArray(parsed)
        ? `Loaded chart 1 of ${parsed.length} — use "Insert batch" for all.`
        : "Chart config loaded.",
      "ok",
    );
  } catch (err) {
    note(`Invalid JSON: ${err instanceof Error ? err.message : String(err)}`, "err");
  }
});

// --- Agenda ------------------------------------------------------------------

function agendaChapters(): string[] {
  return ($("agenda-chapters") as HTMLTextAreaElement).value
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);
}

function renderAgendaPreview() {
  const chapters = agendaChapters();
  const host = $("agenda-preview");
  host.innerHTML = chapters.length
    ? sceneToSvg(buildAgendaScene(chapters, { highlight: 0 }), { background: "#ffffff" })
    : "";
}

$("agenda-chapters").addEventListener("input", renderAgendaPreview);
renderAgendaPreview();

// Auto-update: push edits to the slide shortly after each change.
function maybeAutoUpdate() {
  // Cancel any pending push BEFORE the guard. The guard returns early once the
  // edit target is gone — which is exactly what loading another chart does — and
  // a timer still armed against the previous one would go on to fire. doInsert
  // then finds no edit target, takes its new-chart branch, and drops a chart
  // nobody asked for onto the slide.
  clearTimeout(autoUpdateTimer);
  const on = ($("auto-update") as HTMLInputElement | null)?.checked;
  if (!on || !state.editTarget || !isPowerPointHost()) return;
  autoUpdateTimer = setTimeout(() => void doInsert(false).catch(() => {}), 900);
}
// Unticking has to cancel a push that is already in flight; ticking on its own
// shouldn't push anything until the next edit.
$("auto-update").addEventListener("change", () => {
  if (!($("auto-update") as HTMLInputElement).checked) clearTimeout(autoUpdateTimer);
});

/** think-cell's "click the chart" feel: watch the slide selection. */
function watchSelection() {
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async () => {
      try {
        const found = await loadChartFromSelection();
        const banner = $("selection-banner");
        banner.style.display =
          found && found.target.shapeId !== state.editTarget?.shapeId ? "" : "none";
      } catch {
        /* selection API hiccup — ignore */
      }
    });
  } catch {
    /* event unavailable on this host */
  }
}

function wireInsert() {
  const insertBtn = $("insert") as HTMLButtonElement;
  const insertNewBtn = $("insert-new") as HTMLButtonElement;
  const loadBtn = $("load-selection") as HTMLButtonElement;
  const agendaBtn = $("agenda-insert") as HTMLButtonElement;
  if (isPowerPointHost()) {
    note("");
    insertBtn.disabled = false;
    loadBtn.disabled = false;
    /**
     * Run a host action with a busy note, and lock out re-entry while it runs.
     *
     * Both buttons matter. The one that was CLICKED has to go dead or a slow
     * action invites a second click that queues the whole job again — "Insert
     * demo deck" is 35 slides and ~1,700 shapes, and it stayed live throughout.
     * The primary Insert button has to go dead too, since it acts on the same
     * deck. Disabling only the primary was the worst of both: the clicked
     * button re-entered freely, while Insert went dead WITHOUT looking it (the
     * CSS greys `.el-insert:disabled`, not the primary), so a stuck action read
     * as "Insert does nothing" rather than "Insert is busy".
     *
     * The clicked button comes from the event, so no call site can forget it.
     */
    const guard = (fn: () => Promise<void>) =>
      async function (this: unknown, ev?: Event) {
        const clicked = ev?.currentTarget as HTMLButtonElement | undefined;
        const lock = [insertBtn, clicked].filter((b): b is HTMLButtonElement => !!b && !b.disabled);
        for (const b of lock) b.disabled = true;
        note("Working…", "busy");
        try {
          await fn();
          if (hostNote.textContent?.startsWith("Working…")) {
            note("Done.", "ok");
          }
        } catch (err) {
          // errorText, not err.message: a RichApi.Error's message is generic
          // ("An internal error has occurred") and the useful part is in code
          // and debugInfo, which String(err) throws away.
          note(`Failed: ${errorText(err)}`, "err");
        } finally {
          // Only re-enable what this call disabled — never resurrect a button
          // some other state (no host, no selection) means to keep dead.
          for (const b of lock) b.disabled = false;
        }
      };
    insertBtn.addEventListener("click", guard(() => doInsert(false)));
    insertNewBtn.addEventListener("click", guard(() => doInsert(true)));
    loadBtn.addEventListener("click", guard(doLoadSelection));
    $("selection-banner-load").addEventListener("click", guard(doLoadSelection));
    watchSelection();
    const sameScaleBtn = $("same-scale") as HTMLButtonElement;
    sameScaleBtn.disabled = false;
    sameScaleBtn.addEventListener("click", guard(() => doSameScale("deck")));
    const sameScaleSelBtn = $("same-scale-sel") as HTMLButtonElement;
    sameScaleSelBtn.disabled = false;
    sameScaleSelBtn.addEventListener("click", guard(() => doSameScale("selection")));
    const batchBtn = $("json-insert-batch") as HTMLButtonElement;
    batchBtn.disabled = false;
    batchBtn.addEventListener(
      "click",
      guard(async () => {
        const parsed = JSON.parse(($("json-io") as HTMLTextAreaElement).value);
        const configs: ChartConfig[] = Array.isArray(parsed) ? parsed : [parsed];
        for (const c of configs) {
          const cfg = { ...DEFAULT_SIZE, ...c };
          await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg) });
        }
        note(`Inserted ${configs.length} chart(s) on the current slide.`, "ok");
      }),
    );
    // Elements insert at a small default offset (they're compact shapes).
    for (const [id, scene] of [
      ["harvey-insert", harveyScene],
      ["check-insert", checkScene],
      ["flow-insert", flowScene],
      ["kpi-insert", kpiScene],
      ["table-insert", () => buildTableScene(state.sheet.cells, 480, { totalRow: ($("table-total") as HTMLInputElement).checked })],
    ] as const) {
      const btn = $(id) as HTMLButtonElement;
      btn.disabled = false;
      btn.addEventListener("click", guard(() => insertSceneIntoSlide(scene(), { left: 120, top: 160 }, phaseNote)));
    }
    agendaBtn.disabled = false;
    agendaBtn.addEventListener(
      "click",
      guard(async () => {
        const chapters = agendaChapters();
        if (!chapters.length) return;
        await insertAgendaSlides(chapters.map((_, i) => buildAgendaScene(chapters, { highlight: i })));
      }),
    );
    // Testing aid: one demo slide per chart kind + feature/element highlights.
    // A call we gave up on may still answer. Whatever it says is the only
    // evidence we get about a host that went quiet, so surface it even though
    // the action has already failed — the note is stale by then, but the
    // information is what a bug report needs.
    onLateSync((msg) => note(`Host answered late — ${msg}`, "err"));
    const demoBtn = $("demo-insert") as HTMLButtonElement;
    demoBtn.disabled = false;
    demoBtn.addEventListener(
      "click",
      guard(async () => {
        const items = demoItems();
        // The slowest thing the pane can do — say where it has got to, or a
        // multi-minute run is indistinguishable from a hang.
        await insertDemoDeck(
          items.map((i) => ({ scene: i.scene, tagData: i.configJson })),
          (done, total) => note(`Inserting demo slides… ${done} of ${total}`, "busy"),
        );
        note(`Inserted ${items.length} demo slides at the end of the deck.`, "ok");
      }),
    );
  } else {
    insertBtn.disabled = true;
    loadBtn.disabled = true;
    ($("agenda-insert") as HTMLButtonElement).disabled = true;
    note(
      "Not running inside PowerPoint — use Download SVG, or sideload the manifest to insert native shapes.",
    );
  }
}

// Ribbon deep-link: taskpane.html?kind=waterfall preselects a chart type;
// ?tab=elements opens a tab and ?el=harvey focuses that element's card
// (the ribbon's "Insert element" menu uses these).
const deepLink = new URLSearchParams(location.search);
const requestedKind = deepLink.get("kind");
if (requestedKind && CHART_KINDS.some((k) => k.kind === requestedKind)) {
  applyConfig(sampleConfig(requestedKind as ChartKind), null);
}
const requestedTab = deepLink.get("tab");
if (requestedTab) {
  document.querySelector<HTMLButtonElement>(`.tabs .tab[data-tab="${requestedTab}"]`)?.click();
}
const requestedEl = deepLink.get("el");
if (requestedEl) {
  const card = document.getElementById(`${requestedEl}-insert`)?.closest(".el-card");
  if (card) {
    card.scrollIntoView({ block: "center" });
    card.classList.add("el-flash");
    setTimeout(() => card.classList.remove("el-flash"), 1600);
  }
}

// Injected by vite (see vite.config.ts). Shown in the header so the running
// build is always identifiable — PowerPoint caches the pane, and a stale one is
// otherwise indistinguishable from a fixed one.
declare const __BUILD_STAMP__: string;
const stampEl = document.getElementById("build-stamp");
if (stampEl) stampEl.textContent = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";

const sizeInputs = [$("chart-w"), $("chart-h")] as HTMLInputElement[];
for (const el of sizeInputs) el?.addEventListener("input", renderPreview);

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    wireInsert();
    try {
      localizePane(Office.context.displayLanguage);
    } catch {
      /* no display language available */
    }
  });
} else {
  wireInsert();
  localizePane(new URLSearchParams(location.search).get("lang") ?? undefined);
}
