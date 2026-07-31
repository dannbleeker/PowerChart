import { buildChart, DEFAULT_SIZE, valueExtent } from "../core/chart";
import { PALETTES } from "../core/style";
import type { ChartConfig, ChartKind, Decorations, Series } from "../core/types";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import { sceneToSvg } from "../render/svg";
import {
  canInsertPicture,
  getSelectionBounds,
  insertAgendaSlides,
  insertDemoDeck,
  insertSceneIntoSlide,
  isPowerPointHost,
  newRunId,
  applyReconcilePlan,
  canInsertSlidesFromBase64,
  insertSlidesFromPptx,
  OFFSCREEN_BATCH,
  replaceSlideWithDeck,
  slideHoldsOnlyChart,
  withSlideDeselected,
  slideCount,
  snapshotAddedSlides,
  traceEnvironment,
  wantsAutoPicture,
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
  type ReconcileOutcome,
} from "../render/powerpoint";
import { buildAgendaScene } from "../core/agenda";
import { demoItems, buildResultsScenes, type ResultRow, type ResultsSummary } from "../core/demo";
import type { Scene } from "../core/scene";
import { estimateOfficeShapes } from "../core/scene";
import { describeReconcile, planReconcile } from "../core/reconcile";
import { setTracing, trace, traceLog, tracing } from "../core/trace";
import { buildDeckBase64 } from "../render/pptx-deck";
import type { ExpectedItem, SlideSnapshot } from "../core/reconcile";
import { buildTableScene } from "../core/elements";
import { localizePane, localizeTree, t } from "./i18n";
import { dataToSheet, mountDatasheet, sheetToData, type SheetModel } from "./datasheet";
import { BUILTIN_TEMPLATES } from "./templates";
import { harveyScene, checkScene, flowScene, kpiScene, wireElementPreviews } from "./elements-ui";
import { agendaChapters, wireAgendaPreview } from "./agenda-ui";

interface AppState {
  kind: ChartKind;
  sheet: SheetModel;
  decorations: Partial<Decorations>;
  horizontal: boolean;
  title: string;
  /**
   * Frame size in points. The single source of truth — the size inputs write
   * here and currentConfig() reads here. It used to be read straight off the DOM
   * inputs while every other field lived in state, and that split silently
   * resized loaded charts to the previous chart's dimensions.
   */
  width: number;
  height: number;
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
  /** Per-series color overrides, by ROW INDEX (see seriesMeta). */
  seriesColors: (string | undefined)[];
  /**
   * Series fields the datasheet can't carry (combo `type`, hatch `pattern`,
   * per-point `colors`). The pane rebuilds ChartConfig from the sheet on every
   * preview/insert — without this these fields would be silently dropped from an
   * imported chart and destroyed on re-save.
   *
   * Indexed by ROW, not by series NAME: renaming a series in the datasheet is the
   * sheet's core edit, and a name key meant the rename found no entry, so the
   * overlay line collapsed back into a column and lost its colour/pattern/
   * scenario. (Two series sharing a name collapsed onto one entry for the same
   * reason.) The row index is what the datasheet actually preserves.
   */
  seriesMeta: (Pick<Series, "type" | "pattern" | "colors" | "scenario"> | undefined)[];
  axisTitle: string;
  logScale: boolean;
  /** `render: "image"` — insert one raster picture instead of native shapes.
   *  Owns the key outright now that #render-image exists, so it leaves `extras`. */
  renderImage: boolean;
  /** Footnote / source line ("Kilde: …"). */
  footnote: string;
  /** Comma-separated slice indices to explode (pie/doughnut), 1-based in the UI. */
  pieExplode: string;
  /**
   * Kind-specific / advanced config that has no pane control, preserved verbatim
   * across a load→edit→export (and the shape-tag re-save). Without this, importing
   * an authored chart and re-saving it silently strips these shipped features.
   * pie / waterfall / numberFormat are carried here too but MERGED with their
   * pane-driven parts (explode, total "e" tokens, decimals/suffix/locale) in
   * currentConfig, so the control still wins for the fields it owns.
   */
  extras: Pick<
    ChartConfig,
    | "boxplot"
    | "heatmap"
    | "map"
    | "combo"
    | "gapWidth"
    | "overlap"
    | "multiples"
    | "scatter"
    | "radar"
    | "gantt"
    | "butterfly"
    | "tilemap"
    | "otherBucket"
    | "pareto"
    | "categorySort"
    | "secondaryAxis"
    | "labelOffsets"
    | "pie"
    | "waterfall"
    | "numberFormat"
    | "labels"
  >;
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
    width: cfg.width ?? DEFAULT_SIZE.width,
    height: cfg.height ?? DEFAULT_SIZE.height,
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
    seriesColors: cfg.data.series.map((s) => s.color),
    seriesMeta: cfg.data.series.map((s) =>
      s.type || s.pattern || s.colors || s.scenario
        ? { type: s.type, pattern: s.pattern, colors: s.colors, scenario: s.scenario }
        : undefined,
    ),
    axisTitle: cfg.valueAxisTitle ?? "",
    logScale: !!cfg.logScale,
    renderImage: cfg.render === "image",
    footnote: cfg.footnote ?? "",
    pieExplode: (cfg.pie?.explode ?? []).map((i) => i + 1).join(","),
    extras: {
      boxplot: cfg.boxplot,
      heatmap: cfg.heatmap,
      map: cfg.map,
      combo: cfg.combo,
      gapWidth: cfg.gapWidth,
      overlap: cfg.overlap,
      multiples: cfg.multiples,
      scatter: cfg.scatter,
      radar: cfg.radar,
      gantt: cfg.gantt,
      butterfly: cfg.butterfly,
      tilemap: cfg.tilemap,
      otherBucket: cfg.otherBucket,
      pareto: cfg.pareto,
      categorySort: cfg.categorySort,
      secondaryAxis: cfg.secondaryAxis,
      labelOffsets: cfg.labelOffsets,
      pie: cfg.pie,
      waterfall: cfg.waterfall,
      numberFormat: cfg.numberFormat,
      labels: cfg.labels,
    },
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
  // Size is state now, not a DOM read — the inputs write state on edit.
  const size = { width: state.width, height: state.height };
  const min = state.scaleMin.trim() === "" ? undefined : Number(state.scaleMin);
  const max = state.scaleMax.trim() === "" ? undefined : Number(state.scaleMax);
  const bFrom = Number(state.breakFrom);
  const bTo = Number(state.breakTo);
  const axisBreak =
    state.breakFrom.trim() && state.breakTo.trim() && Number.isFinite(bFrom) && Number.isFinite(bTo) && bTo > bFrom
      ? { from: bFrom, to: bTo }
      : undefined;
  data.series.forEach((s, i) => {
    const c = state.seriesColors[i];
    if (c) s.color = c;
    // Re-attach the fields the datasheet can't carry (combo type, pattern,
    // per-point colors) so an imported chart survives the sheet round-trip.
    const m = state.seriesMeta[i];
    if (m) {
      if (m.type) s.type = m.type;
      if (m.pattern) s.pattern = m.pattern;
      if (m.colors) s.colors = m.colors;
      if (m.scenario) s.scenario = m.scenario;
    }
  });
  const labelParts = state.labelContent
    ? (state.labelContent.split(",") as NonNullable<Decorations["labelContent"]>)
    : undefined;
  const explode = state.pieExplode
    .split(",")
    .map((v) => Number(v.trim()) - 1)
    .filter((v) => Number.isInteger(v) && v >= 0);
  // pie / waterfall / numberFormat are half pane-driven, half carried in extras:
  // the control owns explode / total "e" tokens / decimals+suffix+locale, but an
  // imported chart's semi, breakout, variableRadius, detailGroups, spacerIndices
  // and forceSign live only in extras and would be lost if the control overwrote
  // the whole object. Merge so the pane wins its own fields and the rest survives.
  const pieBase = { ...(state.extras.pie ?? {}) };
  delete pieBase.explode;
  const pie =
    explode.length || Object.keys(pieBase).length ? { ...pieBase, ...(explode.length ? { explode } : {}) } : undefined;
  const numberFormat: ChartConfig["numberFormat"] =
    state.decimals !== "auto" || state.suffix || state.locale !== "en-US" || state.extras.numberFormat
      ? {
          ...(state.extras.numberFormat ?? {}),
          decimals: state.decimals === "auto" ? "auto" : Number(state.decimals),
          suffix: state.suffix || undefined,
          locale: state.locale !== "en-US" ? state.locale : undefined,
        }
      : undefined;
  // The engine's one injected word ("Other"); pass a localized override only when
  // a non-English pane is active, so English configs stay byte-identical.
  const otherLabel = t("Other");
  const labels =
    otherLabel !== "Other" || state.extras.labels
      ? { ...state.extras.labels, ...(otherLabel !== "Other" ? { other: otherLabel } : {}) }
      : undefined;
  return {
    kind: state.kind,
    data,
    ...state.extras,
    labels,
    horizontal: state.horizontal || undefined,
    footnote: state.footnote || undefined,
    pie,
    valueAxisTitle: state.axisTitle || undefined,
    logScale: state.logScale || undefined,
    // Omitted rather than written as "shapes": the key is optional and the
    // engine treats undefined as shapes, so spelling out the default would add
    // noise to every exported config.
    render: state.renderImage ? "image" : undefined,
    style: mergedStyle(),
    ...size,
    title: state.title || undefined,
    decorations: { ...state.decorations, labelContent: labelParts },
    waterfall: { ...(state.extras.waterfall ?? {}), totalIndices: [...totals] },
    segmentOrder: state.segmentOrder === "sheet" ? undefined : state.segmentOrder,
    axisBreak,
    scale:
      (min != null && Number.isFinite(min)) || (max != null && Number.isFinite(max))
        ? { min: Number.isFinite(min!) ? min : undefined, max: Number.isFinite(max!) ? max : undefined }
        : undefined,
    numberFormat,
  };
}

// --- UI wiring ------------------------------------------------------------

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const gallery = $("gallery");
const preview = $("preview");
const optionsHost = $("options");
const hostNote = $("host-note");

// eslint-disable-next-line prefer-const -- forward-declared; assigned once after wiring
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
const statusStrip = document.getElementById("status-strip");
const statusBar = document.getElementById("status-bar");
const statusElapsed = document.getElementById("status-elapsed");

function note(text: string, status: "ok" | "err" | "busy" | "none" = "none", params?: Record<string, string | number>) {
  // Route status text through the runtime translator so a localized pane
  // announces it in the user's language. `text` is the English source string
  // (an EN catalogue key), optionally carrying {placeholders} that `params`
  // fills after translation. The aria-live host-note reads the change to a
  // screen reader.
  hostNote.textContent = t(text, params);
  hostNote.className = status === "none" ? "hint" : `hint status-${status}`;
  // The strip carries the note now, so it has to follow it: shown whenever
  // there is something to say, collapsed when there is not.
  statusStrip?.toggleAttribute("hidden", !text);
  statusBar?.toggleAttribute("hidden", status !== "busy");
  if (status !== "busy") setProgress(null);
}

/**
 * How far along, when we honestly know: a fraction for work we complete in
 * steps, "busy" for work we hand to PowerPoint in one go.
 *
 * A single insert is ONE context.sync() and Office.js reports nothing until it
 * lands, so any percentage there would be invented — and a bar that sticks at
 * 99% is a worse lie than no bar. Chunked work (the demo deck) really does know,
 * so it gets a real one.
 */
function setProgress(p: number | "busy" | null) {
  if (!statusBar) return;
  const fill = statusBar.querySelector("i");
  if (p === "busy") {
    statusBar.classList.add("indeterminate");
    if (fill) fill.style.width = "";
  } else {
    statusBar.classList.remove("indeterminate");
    if (fill) fill.style.width = p == null ? "0" : `${Math.round(Math.max(0, Math.min(1, p)) * 100)}%`;
  }
}

/**
 * Count the seconds while the host works. It is the only number we can report
 * mid-sync, and on a host that takes 20s to draw a chart, a number that moves
 * is the difference between "working" and "dead".
 */
let elapsedTimer: ReturnType<typeof setInterval> | undefined;
function startElapsed() {
  const t0 = Date.now();
  stopElapsed();
  const tick = () => {
    if (statusElapsed) statusElapsed.textContent = `${Math.round((Date.now() - t0) / 1000)}s`;
  };
  tick();
  elapsedTimer = setInterval(tick, 1000);
}
function stopElapsed() {
  clearInterval(elapsedTimer);
  elapsedTimer = undefined;
  if (statusElapsed) statusElapsed.textContent = "";
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
  const renderBox = document.getElementById("render-image") as HTMLInputElement | null;
  if (renderBox) renderBox.checked = state.renderImage;
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
  {
    label: "Columns & bars",
    kinds: ["stacked", "clustered", "stacked100", "waterfall", "mekko", "butterfly", "cascade", "funnel"],
  },
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
  // Select a tab and mirror it into the ARIA tab pattern: one tab is
  // aria-selected and in the tab order (tabindex 0), the rest are out of it
  // (tabindex -1) and reachable only via the arrow keys. Without this a screen
  // reader can't tell which mode is active and keyboard users tab through every
  // one. `focus` is false for the initial paint (don't steal focus on boot).
  const select = (tab: HTMLButtonElement, focus: boolean) => {
    const name = tab.dataset.tab;
    tabs.forEach((t) => {
      const on = t === tab;
      t.classList.toggle("active", on);
      t.setAttribute("aria-selected", String(on));
      t.tabIndex = on ? 0 : -1;
    });
    panels.forEach((p) => p.classList.toggle("active", p.dataset.panel === name));
    showBarFor(name);
    if (focus) tab.focus();
  };
  tabs.forEach((tab, i) => {
    tab.addEventListener("click", () => select(tab, false));
    tab.addEventListener("keydown", (e) => {
      // Left/Right move between tabs (wrapping); Home/End jump to the ends —
      // the WAI-ARIA tabs keyboard model.
      const delta = e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0;
      let next = -1;
      if (delta) next = (i + delta + tabs.length) % tabs.length;
      else if (e.key === "Home") next = 0;
      else if (e.key === "End") next = tabs.length - 1;
      if (next >= 0) {
        e.preventDefault();
        select(tabs[next], true);
      }
    });
  });
  // Seed the ARIA state to match the markup's default active tab, without
  // moving focus. A ?tab= deep link CLICKS its tab afterwards, re-running select.
  const initial = tabs.find((t) => t.classList.contains("active")) ?? tabs[0];
  if (initial) select(initial, false);
}

/** The footer "⋯" overflow menu holding the secondary actions (edit selected,
 *  same scale, download). Opens upward; closes on item click, outside click,
 *  or Escape. */
function wireActionsMenu() {
  const btn = document.getElementById("more-actions");
  const menu = document.getElementById("actions-menu");
  if (!btn || !menu) return;
  // Give the popup the ARIA menu role so a screen reader announces it as a menu
  // and its entries as menu items, reachable with the arrow keys.
  btn.setAttribute("aria-haspopup", "menu");
  menu.setAttribute("role", "menu");
  const items = Array.from(menu.querySelectorAll<HTMLButtonElement>("button"));
  items.forEach((it) => {
    it.setAttribute("role", "menuitem");
    it.tabIndex = -1;
  });
  const focusItem = (i: number) => items[(i + items.length) % items.length]?.focus();
  const setOpen = (open: boolean, restoreFocus = false) => {
    menu.hidden = !open;
    btn.setAttribute("aria-expanded", String(open));
    if (open)
      focusItem(0); // move into the menu so the keyboard lands there
    else if (restoreFocus) btn.focus(); // Escape/close returns focus to the trigger
  };
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    setOpen(menu.hidden);
  });
  menu.addEventListener("click", () => setOpen(false, true));
  menu.addEventListener("keydown", (e) => {
    const i = items.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === "ArrowDown") {
      e.preventDefault();
      focusItem(i + 1);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      focusItem(i - 1);
    } else if (e.key === "Home") {
      e.preventDefault();
      focusItem(0);
    } else if (e.key === "End") {
      e.preventDefault();
      focusItem(items.length - 1);
    }
  });
  document.addEventListener("click", (e) => {
    if (!menu.hidden && e.target !== btn && !menu.contains(e.target as Node)) setOpen(false);
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !menu.hidden) setOpen(false, true);
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
    { key: "grandTotal", label: "Grand total", group: G.labels },
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

  // pairControl only knows from/to, but decorations.cagr also carries the series
  // the arrow is anchored to — and core/decor.ts reads it. Emitting the bare pair
  // dropped that anchor the moment the spinner was touched, so the arrow silently
  // switched to the column totals and printed a different growth rate.
  const cagrSeries = d.cagr?.series;
  G.analysis.body.appendChild(
    pairControl("CAGR arrow", d.cagr, nCats(), (pair) => {
      d.cagr = pair ? { ...pair, series: cagrSeries } : undefined;
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
  for (const [el, val] of [
    [scMin, state.scaleMin],
    [scMax, state.scaleMax],
  ] as const) {
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
  for (const [el, val] of [
    [abFrom, state.breakFrom],
    [abTo, state.breakTo],
  ] as const) {
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
  const decOptions = ["auto", "0", "1", "2"];
  // An imported chart may carry decimals outside the four presets (3, say).
  // Without an option of its own the select falls to value "", and emitNf — which
  // writes BOTH controls — then rewrote decimals to 0 as soon as the suffix box
  // was touched. Carry the loaded value as its own option instead.
  if (!decOptions.includes(state.decimals)) decOptions.push(state.decimals);
  for (const v of decOptions) {
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
    input.value = state.seriesColors[i] ?? palette[i % palette.length];
    input.addEventListener("input", () => {
      state.seriesColors[i] = input.value;
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
  const emit = () => onChange(cb.checked ? { from: Number(from.value) - 1, to: Number(to.value) - 1 } : undefined);
  cb.addEventListener("change", emit);
  from.addEventListener("input", emit);
  to.addEventListener("input", emit);
  wrap.append(cb, label, " from ", from, " to ", to);
  return wrap;
}

/**
 * The canvas colour an SVG is painted on. Forcing white here put a dark-theme
 * config's light text on a white rect (1.13:1 contrast) in the preview AND in
 * both downloads, while insertSceneIntoSlide drops the same shapes onto the real
 * (dark) slide with no background rect at all — canvas != export. Default charts
 * are byte-identical: style.background already defaults to #ffffff.
 */
const canvasColor = (cfg: ChartConfig) => cfg.style?.background ?? "#ffffff";

function renderPreviewNow() {
  try {
    const cfg = currentConfig();
    const scene = buildChart(cfg);
    preview.innerHTML = sceneToSvg(scene, { background: canvasColor(cfg) });
  } catch (err) {
    // textContent, not innerHTML: the error message can carry config-derived
    // strings, and an innerHTML sink here would be a second injection path.
    preview.replaceChildren();
    const p = document.createElement("p");
    p.className = "hint";
    p.textContent = `Could not render: ${err instanceof Error ? err.message : String(err)}`;
    preview.appendChild(p);
  }
  maybeAutoUpdate();
}

/**
 * Coalesce preview renders onto a trailing timer. Every control and every
 * datasheet keystroke calls renderPreview(); each render is a full
 * buildChart()+sceneToSvg()+innerHTML reparse, so firing one per keystroke was
 * measurable jank on a large sheet. A short trailing debounce collapses a burst
 * of edits into a single render of the final state. (snapshot() stays
 * synchronous — undo granularity depends on it.)
 */
let renderTimer: ReturnType<typeof setTimeout> | undefined;
function renderPreview() {
  clearTimeout(renderTimer);
  renderTimer = setTimeout(renderPreviewNow, 80);
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

sheetApi = mountDatasheet(
  $("datasheet"),
  state.sheet,
  (sheet) => {
    state.sheet = sheet;
    snapshot();
    renderPreview();
  },
  // seriesColors / seriesMeta are POSITIONAL, so a row or column the grid
  // splices in or out has to move them too — otherwise every series below the
  // edit inherits its neighbour's colour and combo type, and a routine
  // datasheet edit silently changes what the chart draws.
  (change) => {
    const colorsOf = (m: AppState["seriesMeta"][number]) => m?.colors;
    switch (change.kind) {
      case "series-insert":
        state.seriesColors.splice(change.index, 0, undefined);
        state.seriesMeta.splice(change.index, 0, undefined);
        break;
      case "series-remove":
        state.seriesColors.splice(change.index, 1);
        state.seriesMeta.splice(change.index, 1);
        break;
      case "category-insert":
      case "category-remove":
        // Per-point colours are indexed by CATEGORY, so they follow the column.
        for (const m of state.seriesMeta) {
          const cs = colorsOf(m);
          if (!cs) continue;
          if (change.kind === "category-insert") cs.splice(change.index, 0, null);
          else cs.splice(change.index, 1);
        }
        break;
      case "reset":
        // A transpose turns series into categories: no positional mapping survives.
        state.seriesColors = [];
        state.seriesMeta = [];
        break;
    }
  },
);
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
  const cfg = currentConfig();
  const svg = sceneToSvg(buildChart(cfg), { background: canvasColor(cfg) });
  const blob = new Blob([svg], { type: "image/svg+xml" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "powerchart.svg";
  a.click();
  URL.revokeObjectURL(a.href);
});

/** Oversample for rasterized output: 2× the scene's point size, i.e. 144 dpi.
 *  A point-for-pixel PNG looks fuzzy the moment a slide is zoomed. */
const RASTER_SCALE = 2;

/**
 * Rasterize a scene to a bare base64 PNG for `InsertOptions.pictureBase64`.
 *
 * Deliberately **transparent** — no background is painted. Shapes mode drops the
 * chart's ink straight onto the real slide with no background rect at all (see
 * `canvasColor`), so painting one here would make image mode land an opaque
 * white box over a themed or dark slide. Transparency is what keeps the two
 * modes visually interchangeable, which is the whole promise of the flag.
 *
 * The `download-png` handler below does the same SVG→Image→canvas dance but ends
 * in `toBlob` for a file download; this one ends in `toDataURL` because Office.js
 * wants base64. The shared middle is three lines, and factoring it would couple
 * an export button to the insert path for no real gain.
 *
 * Resolves with the payload after the comma; rejects on a decode failure, a
 * missing 2D context, or a `toDataURL` that hands back something that isn't a
 * PNG data URI (jsdom returns null). The caller degrades to native shapes.
 */
function rasterizeScene(scene: Scene): Promise<string> {
  const svg = sceneToSvg(scene);
  return new Promise<string>((resolve, reject) => {
    const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
    const img = new Image();
    img.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(scene.width * RASTER_SCALE);
        canvas.height = Math.round(scene.height * RASTER_SCALE);
        const ctx = canvas.getContext("2d");
        if (!ctx) throw new Error("no 2d canvas context");
        ctx.scale(RASTER_SCALE, RASTER_SCALE);
        ctx.drawImage(img, 0, 0, scene.width, scene.height);
        const uri = canvas.toDataURL("image/png");
        if (!uri || !uri.startsWith("data:image/png")) throw new Error("canvas could not encode a PNG");
        resolve(uri.slice(uri.indexOf(",") + 1));
      } catch (err) {
        reject(err instanceof Error ? err : new Error(String(err)));
      } finally {
        URL.revokeObjectURL(url);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("the browser could not decode the chart SVG"));
    };
    img.src = url;
  });
}

/**
 * The picture payload for one insert, plus the reason there isn't one.
 *
 * `warn` is returned rather than noted here on purpose: every insert path posts
 * `phaseNote` as its first act, which would immediately overwrite a note posted
 * before the insert. The caller notes it AFTER the insert resolves, where it
 * survives — `guard` only prints "Done." when the note is unchanged from the
 * busy text, and any phase note has already changed it.
 */
async function chartPicture(cfg: ChartConfig, scene: Scene): Promise<{ png?: string; warn?: string }> {
  // A chart nobody asked to rasterize, on the one host that cannot survive
  // drawing it. The densest kinds are far past what PowerPoint on the web will
  // take as shapes — violin 253, area 176, tile map 122, waffle 103 — and the
  // budget below is the same number the demo deck has always used to decide a
  // chart "too dense for this host". It used to skip those and stamp a
  // placeholder; drawing a picture is strictly better than not drawing.
  const shapes = estimateOfficeShapes(scene);
  if (
    wantsAutoPicture(shapes, {
      web: isWebHost(),
      canPicture: canInsertPicture(),
      alreadyPicture: cfg.render === "image",
    })
  ) {
    try {
      return {
        png: await rasterizeScene(scene),
        warn: `That chart is ${shapes} shapes — too many for PowerPoint on the web to draw. Inserted as a picture; "Explode to native shapes" turns it back.`,
      };
    } catch {
      // Could not rasterize — fall through and draw the shapes. Slow and
      // risky beats refusing to insert the user's chart.
    }
  }
  if (cfg.render !== "image") return {};
  if (!canInsertPicture()) {
    return {
      warn: "This PowerPoint can't insert pictures (needs PowerPointApi 1.8) — inserted native shapes instead.",
    };
  }
  try {
    return { png: await rasterizeScene(scene) };
  } catch (err) {
    console.warn(`PC-IMG-RASTER could not rasterize the chart in the browser. ${errorText(err)}`);
    return { warn: "Couldn't rasterize the chart — inserted native shapes instead." };
  }
}

/**
 * PNG export: the native-shapes output is the real deliverable, but SVG doesn't
 * render in email/chat, so a rasterized fallback earns its place. Rasterize the
 * SAME preview SVG through a canvas at 2× for a crisp bitmap. The SVG carries no
 * foreignObject or external refs, so the canvas never tags as tainted and
 * toBlob() is allowed. Best-effort: a decode failure surfaces as an error note
 * rather than a silent no-op.
 */
$("download-png").addEventListener("click", () => {
  const cfg = currentConfig();
  const scene = buildChart(cfg);
  const svg = sceneToSvg(scene, { background: canvasColor(cfg) });
  const scale = 2;
  const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
  const img = new Image();
  img.onload = () => {
    try {
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(scene.width * scale);
      canvas.height = Math.round(scene.height * scale);
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("no 2d canvas context");
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, scene.width, scene.height);
      canvas.toBlob((png) => {
        if (!png) {
          note("Couldn't encode the PNG on this browser.", "err");
          return;
        }
        const a = document.createElement("a");
        a.href = URL.createObjectURL(png);
        a.download = "powerchart.png";
        a.click();
        URL.revokeObjectURL(a.href);
      }, "image/png");
    } catch (err) {
      note("Couldn't render PNG: {error}", "err", { error: err instanceof Error ? err.message : String(err) });
    } finally {
      URL.revokeObjectURL(url);
    }
  };
  img.onerror = () => {
    URL.revokeObjectURL(url);
    note("Couldn't render the preview to PNG.", "err");
  };
  img.src = url;
});

/**
 * Encode the current chart into a shareable deep link. The config rides in the
 * URL hash (base64 of the JSON) so it never hits a server log, and the hash is
 * decoded on boot — reopening the exact chart on the hosted gallery. Round-trips
 * are UTF-8-safe via encodeURIComponent before btoa.
 */
const CONFIG_HASH = "#c=";
$("copy-link").addEventListener("click", async () => {
  const json = JSON.stringify(currentConfig());
  const link = location.origin + location.pathname + location.search + CONFIG_HASH + btoa(encodeURIComponent(json));
  try {
    await navigator.clipboard.writeText(link);
    note("Shareable chart link copied to the clipboard.", "ok");
  } catch {
    // Clipboard blocked (no permission / not focused): drop the link into the
    // JSON box so it can still be copied by hand.
    ($("json-io") as HTMLTextAreaElement).value = link;
    note("Clipboard blocked — the link is in the JSON box, copy it from there.", "err");
  }
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
  // The phase word is itself a catalogue key; the detail (a dynamic id) trails
  // untranslated in parentheses.
  note("Working… {phase}", "busy", { phase: t(said[phase]) + (detail ? ` (${detail})` : "") });
}

async function doInsert(asNew: boolean) {
  let cfg = currentConfig();
  if (!asNew && state.editTarget) {
    const scene = buildChart(cfg);
    const { png, warn } = await chartPicture(cfg, scene);
    // Adopt the target the update hands back. Every shape was replaced, so the
    // one we just used is dead: keeping it made the SECOND update resolve a
    // shape id that no longer existed, get filtered out as "the user deleted
    // this chart", and do nothing — silently. With auto-update on, that meant
    // only the first debounced push ever landed.
    const { next, swapped, duplicated, picture } = await updateChartResilient(scene, state.editTarget, {
      tagData: JSON.stringify(cfg),
      pictureBase64: png,
    });
    if (duplicated) {
      // Both slides hold the chart: the rebuilt one and the original PowerPoint
      // would not remove. Nothing here can safely pick one — deleting the wrong
      // slide loses the user's notes or transitions — so name it and let them.
      state.editTarget = null;
      renderActionState();
      note(
        "Rebuilt that slide, but PowerPoint would not remove the original — the chart is now on two slides. Delete whichever you don't want.",
        "err",
      );
      return;
    }
    if (swapped) {
      // The chart is on a NEW slide, so the old target is dead. Say so rather
      // than leaving a stale target that makes every later push no-op.
      state.editTarget = null;
      renderActionState();
      note("Rebuilt that slide — PowerPoint would not redraw it in place. Select the chart to keep editing.", "ok");
      return;
    }
    if (next) {
      state.editTarget = next;
      if (picture && !png) note("Drawn as a picture — PowerPoint would not redraw the shapes.", "err");
      else if (warn) note(warn, "err");
      return;
    }
    // null means the target slide or shape is gone — nothing was written. The
    // caller's guard saw an unchanged note and printed "Done." in green, and the
    // stale target kept the button reading "Update chart", so every later push
    // (including every auto-update) no-opped just as silently. Say so, and fall
    // back to inserting a new chart.
    state.editTarget = null;
    renderActionState();
    note("That chart is no longer on the slide — insert it again.", "err");
    return;
  }
  // New chart: use the selected placeholder's bounds when one is selected.
  const bounds = await getSelectionBounds();
  // Resize BEFORE building the scene, so the raster matches the frame that sizes
  // the picture's rect — a raster of a differently-sized scene would be stretched.
  if (bounds && bounds.width > 40 && bounds.height > 40) {
    cfg = { ...cfg, width: bounds.width, height: bounds.height };
  }
  const scene = buildChart(cfg);
  const { png, warn } = await chartPicture(cfg, scene);
  if (bounds && bounds.width > 40 && bounds.height > 40) {
    await insertSceneIntoSlide(
      scene,
      { tagData: JSON.stringify(cfg), left: bounds.left, top: bounds.top, pictureBase64: png },
      phaseNote,
    );
  } else {
    await insertSceneIntoSlide(
      scene,
      {
        tagData: JSON.stringify(cfg),
        left: 60 + insertOffset,
        top: 90 + insertOffset,
        pictureBase64: png,
      },
      phaseNote,
    );
    insertOffset = (insertOffset + 14) % 84;
  }
  state.editTarget = null;
  renderActionState();
  // After the insert, never before: phaseNote would have overwritten it.
  if (warn) note(warn, "err");
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
  // Apply the shared scale FIRST, then build. Rasterising inside the same map
  // would capture the pre-scale scene — an image chart would be re-inserted
  // showing the old axis while the pane reported success.
  const rescaled = parsed.map((c) => {
    c.cfg.scale = { min: min < 0 ? min : undefined, max };
    return { ...c, scene: buildChart(c.cfg) };
  });
  // Rasterise every image-mode chart before opening the request context — one
  // context for the whole deck is the property doSameScale exists to protect,
  // and awaiting a canvas inside it would be both slower and unsafe.
  const pictures = await Promise.all(rescaled.map((c) => chartPicture(c.cfg, c.scene)));
  // A chart whose redraw stalls has already had its old shapes deleted, so it
  // is left blank and the user has to be told which. Silence here meant a
  // deck-wide operation could quietly empty charts the user never looked at.
  const stalled: string[] = [];
  await updateChartsInSlides(
    rescaled.map((c, i) => ({
      scene: c.scene,
      target: c.target,
      opts: { tagData: JSON.stringify(c.cfg), pictureBase64: pictures[i].png },
    })),
    (item) => stalled.push(item.scene.title || "an untitled chart"),
  );
  if (stalled.length) {
    note("Same scale: PowerPoint would not redraw {n} chart(s) — {which}. They are now empty; undo (Ctrl+Z) restores them.", "err", {
      n: stalled.length,
      which: stalled.join(", "),
    });
    return;
  }
  note("Same scale applied to {n} charts (max {max}).", "ok", { n: parsed.length, max });
  const degraded = pictures.filter((p) => p.warn).length;
  if (degraded) {
    note("Same scale applied, but {n} image chart(s) fell back to native shapes.", "err", { n: degraded });
  }
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

/**
 * Turn the selected picture chart back into native, editable shapes.
 *
 * The whole cost of image mode is that the chart stops being editable, so this
 * is the door back. It adds NO renderer code: read the config off the picture's
 * tag, force `render: "shapes"`, and hand it to the ordinary in-place update
 * with **no** `pictureBase64`. That omission is the explode — `wantsPicture`
 * goes false and the node loop runs instead.
 *
 * Works on a shapes-mode chart too, where it is an idempotent redraw rather than
 * a special case worth refusing. Acts on the SELECTION, which need not be the
 * chart the pane is currently editing.
 */
async function doExplode() {
  const found = await loadChartFromSelection();
  if (!found) {
    note("Select an inserted PowerChart first — Explode re-draws it as native shapes.", "err");
    return;
  }
  const cfg = { ...(JSON.parse(found.configJson) as ChartConfig), render: "shapes" as const };
  // Same live-canvas wall as an ordinary update: exploding a picture draws
  // every shape onto the slide in view. Look away while it happens.
  const next = await withSlideDeselected([found.target.slideId], (deselected) =>
    updateChartInSlide(buildChart(cfg), found.target, {
      tagData: JSON.stringify(cfg),
      ...(deselected ? { shapesPerSync: OFFSCREEN_BATCH } : {}),
    }),
  );
  if (!next) {
    // The picture is gone from the slide. Only clear the pane's edit target if
    // it was pointing at this same shape — Explode works on the selection.
    if (state.editTarget?.shapeId === found.target.shapeId) {
      state.editTarget = null;
      renderActionState();
    }
    note("That chart is no longer on the slide.", "err");
    return;
  }
  // Adopt the returned target: every shape was replaced, so the old one is dead
  // (the trap doInsert documents — a stale target makes the NEXT update no-op).
  applyConfig(cfg, next);
  $("selection-banner").style.display = "none";
  note("Exploded to native shapes — edits now update them in place.", "ok");
}

// --- Elements (harvey balls, checkboxes, process flow, table) -----------------

wireElementPreviews();

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
  const cfg = source === "builtin" ? BUILTIN_TEMPLATES.find((t) => t.name === name)?.config : loadTemplates()[name];
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
    note("Style import failed: {error}", "err", { error: err instanceof Error ? err.message : String(err) });
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
    if (Array.isArray(parsed))
      note('Loaded chart 1 of {total} — use "Insert batch" for all.', "ok", { total: parsed.length });
    else note("Chart config loaded.", "ok");
  } catch (err) {
    note("Invalid JSON: {error}", "err", { error: err instanceof Error ? err.message : String(err) });
  }
});

// --- Agenda ------------------------------------------------------------------

wireAgendaPreview();

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

// The picture/shapes choice changes nothing about the SCENE, so the preview is
// unaffected — it always shows the vector chart. It only changes what the next
// insert hands PowerPoint, which is why this touches state and nothing else.
$("render-image").addEventListener("change", () => {
  state.renderImage = ($("render-image") as HTMLInputElement).checked;
});

/** think-cell's "click the chart" feel: watch the slide selection. */
function watchSelection() {
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async () => {
      try {
        const found = await loadChartFromSelection();
        const banner = $("selection-banner");
        banner.style.display = found && found.target.shapeId !== state.editTarget?.shapeId ? "" : "none";
      } catch {
        /* selection API hiccup — ignore */
      }
    });
  } catch {
    /* event unavailable on this host */
  }
}

/**
 * A one-line host descriptor for the demo title/results slides, e.g.
 * "PowerPoint · OfficeOnline · 16.0.1". Web vs desktop vs Mac behave differently
 * under the demo's shape budget, so a run's PDF should say which one it was.
 * Guarded — `Office.context.diagnostics` is absent outside a host.
 */
function describeHost(): string {
  try {
    const d = Office.context?.diagnostics;
    if (d?.host) return `${d.host} · ${d.platform} · ${d.version}`;
  } catch {
    /* Office.context unavailable — fall through */
  }
  return "unknown host";
}

/**
 * The last demo run, kept at module scope so it can be saved AFTER the fact.
 *
 * A run's own record used to live and die inside the click handler's closure:
 * when a run ended badly the summary slide was the first casualty, and the
 * only surviving evidence was whatever the user thought to copy out of a
 * one-line note. A run that reveals a host bug is worth more than that.
 */
let lastRunLog: RunLog | undefined;

interface RunLog {
  build: string;
  host: string;
  smoke: boolean;
  totalMs: number;
  items: {
    title: string;
    status: string;
    shapes: number;
    ms: number;
    grouped: boolean;
    lateOutcome: string;
  }[];
  deck: {
    slidesAdded: number;
    addsIssued: number;
    lost: number;
    blank: { position: number; title: string | null }[];
  };
  /** The settled truth: what the deck held once the host stopped moving. */
  reconcile?: ReconcileOutcome;
  /** Step-by-step record, when Verbose trace was on for the run. */
  trace?: {
    entries: { ms: number; scope: string; message: string; data?: Record<string, unknown> }[];
    dropped: number;
  };
}

/**
 * Verify and repair the demo slides wherever they are in the deck.
 *
 * By slot tag, never by position. The deck's own layout is the host's
 * business: `insertSlidesFromBase64` puts slides at the FRONT unless told
 * otherwise, and a run that assumed "the slides we just added are at the end"
 * read one slide short at the front and swept in the user's own title slide at
 * the back. The span between the first and last tagged slide is ours; nothing
 * outside it is read, let alone touched.
 *
 * "Tagged" means tagged BY THIS RUN. The span used to be drawn around every
 * PowerChart slot tag in the presentation, which is only the same thing the
 * first time the deck is inserted. Insert it twice — or duplicate one demo
 * slide, which copies its tag — and the span covered both copies, every item
 * matched two slides, and the pass deleted one whole healthy run as the other
 * one's duplicate. Anything inside the span that is not this run's is reported
 * and left alone.
 */
async function repairDeckSpan(
  expected: ExpectedItem[],
  tagFor: (slot: number) => string | undefined,
  run: string,
): Promise<VerifyResult> {
  let snapshots: SlideSnapshot[];
  try {
    const count = await slideCount();
    snapshots = await snapshotAddedSlides(0, count);
  } catch (err) {
    return { kind: "error", why: errorText(err) };
  }
  const tagged = snapshots.filter((s) => s.run === run);
  // No slot tag anywhere means the read found slides but could not identify
  // one of them. Saying "verified" would be a lie and saying nothing is how a
  // whole verification pass went missing from a run report without anyone
  // noticing — so name it.
  if (!tagged.length) {
    return {
      kind: "unidentified",
      why: snapshots.length
        ? `read ${snapshots.length} slide(s), none carrying this run's slot tag`
        : "read no slides at all",
    };
  }
  const first = Math.min(...tagged.map((s) => s.index));
  const last = Math.max(...tagged.map((s) => s.index));
  const inSpan = snapshots.filter((s) => s.index >= first && s.index <= last);
  const plan = planReconcile(inSpan, expected, { dropOrphanBlanks: true, run });
  try {
    if (!plan.actions.length)
      return {
        kind: "ok",
        outcome: { snapshots: inSpan, plan, applied: { unstamped: 0, regrouped: 0, deleted: 0 }, refused: 0 },
      };
    return { kind: "ok", outcome: await applyReconcilePlan(plan, tagFor, { left: 60, top: 90 }, inSpan) };
  } catch (err) {
    return { kind: "error", why: errorText(err) };
  }
}

/**
 * Why a verification pass produced no verdict, when it produced none.
 *
 * A `null` here used to mean three different things — nothing tagged, a host
 * that would not answer, and a repair that threw — and the caller could tell
 * them apart from none of them. The first fast-path run reported a raw slide
 * count instead of a settled verdict and gave no hint why, which is precisely
 * the failure mode this whole line of work exists to eliminate.
 */
type VerifyResult =
  { kind: "ok"; outcome: ReconcileOutcome } | { kind: "unidentified"; why: string } | { kind: "error"; why: string };

/**
 * Update a chart in place without asking the live canvas to do the impossible.
 *
 * Redrawing a chart is the add-in's worst case on PowerPoint web: every shape
 * replaced, onto the one slide guaranteed to be on screen. A real run died on
 * the FIRST batch — "did not respond while drawing shapes 1-10 of 39". Three
 * attempts, least destructive first:
 *
 *  1. Look away and redraw. Selecting another slide puts the target
 *     off-screen, where the host swallows 40 shapes a sync instead of choking
 *     at 10, and the selection is restored afterwards. Nothing is lost.
 *  2. Swap the slide. Generate a one-slide .pptx and replace the slide with
 *     it — one host call, no drawing at all. Only offered when the slide holds
 *     NOTHING but the chart, because the replacement is a new slide and does
 *     not carry the old one's notes or transitions. The user is told when this
 *     happens, since their edit target moves with it.
 *  3. Draw it as a picture. Never stalls, keeps the config tag so the chart is
 *     still re-editable and can be exploded back to shapes — but it is a
 *     raster, and that is a real cost, so it is the floor and not the default.
 */
async function updateChartResilient(
  scene: Scene,
  target: EditTarget,
  opts: { tagData?: string; pictureBase64?: string },
): Promise<{ next: EditTarget | null; swapped?: boolean; duplicated?: boolean; picture?: boolean }> {
  let stall: unknown;
  try {
    const next = await withSlideDeselected([target.slideId], (deselected) =>
      updateChartInSlide(scene, target, deselected ? { ...opts, shapesPerSync: OFFSCREEN_BATCH } : opts),
    );
    return { next };
  } catch (err) {
    console.warn("PowerChart: in-place redraw stalled — trying the slide swap", err);
    stall = err;
  }

  if (canInsertSlidesFromBase64() && opts.tagData && (await slideHoldsOnlyChart(target.slideId))) {
    try {
      note("Rebuilding that slide…", "busy");
      const built = await buildDeckBase64([{ scene, title: "Chart", configJson: opts.tagData }]);
      const swap = await replaceSlideWithDeck(target.slideId, built.base64);
      if (swap === "swapped") return { next: null, swapped: true };
      // The new slide landed; only the old one's removal failed. Falling
      // through to the picture layer here would rasterize the chart onto that
      // surviving original — leaving the user with the chart twice, once as
      // shapes and once as a picture, and a message saying it went fine. Stop
      // and say what happened instead.
      if (swap === "duplicated") return { next: null, duplicated: true };
    } catch (err) {
      console.warn("PowerChart: slide swap failed — falling back to a picture", err);
    }
  }

  try {
    // Bounded: a canvas that never fires `onload` would otherwise hang the
    // update forever, which is a worse outcome than the stall we are already
    // recovering from.
    const png = await Promise.race([
      rasterizeScene(scene),
      new Promise<never>((_, reject) => setTimeout(() => reject(new Error("rasterizing timed out")), 10_000)),
    ]);
    const next = await updateChartInSlide(scene, target, { ...opts, pictureBase64: png });
    return { next, picture: true };
  } catch (err) {
    // Nothing left to try. Surface the host's own words rather than a
    // rasterizer error the user can do nothing about.
    console.warn("PowerChart: picture fallback failed too", err);
    throw stall;
  }
}

/**
 * Insert the demo deck as ONE generated .pptx instead of drawing it shape by
 * shape — see `src/render/pptx-deck.ts` for why that is worth doing.
 *
 * Returns null when the fast path did not run and it is SAFE to fall back to
 * the shape-by-shape path: the host has no `insertSlidesFromBase64`, the file
 * could not be built, or the call failed with nothing landed. It returns a
 * message — never null — once any slide has landed, because falling back after
 * a partial insert would draw the whole deck a second time on top of it.
 */
async function insertDemoDeckAsFile(
  items: { scene: Scene; title: string; configJson?: string }[],
  smoke: boolean,
): Promise<{ text: string; status: "ok" | "err" } | null> {
  const t0 = Date.now();
  const before = await slideCount();
  // Identity for this insert, carried on every slide's slot tag. Without it the
  // repair pass below cannot tell these slides from the ones an earlier insert
  // left in the same deck, and "cannot tell" ends in a delete.
  const run = newRunId();
  let built: { base64: string; shapesPerSlide: number[] };
  try {
    note("Building the deck…", "busy");
    built = await buildDeckBase64(
      items.map((it, i) => ({ scene: it.scene, title: it.title, configJson: it.configJson, slot: i, run })),
    );
  } catch (err) {
    console.warn("PowerChart: could not build the deck file — falling back to shapes", err);
    return null;
  }
  let added: number;
  try {
    note("Handing the deck to PowerPoint…", "busy");
    setProgress("busy");
    added = await insertSlidesFromPptx(built.base64, items.length);
  } catch (err) {
    // A throw is not proof that nothing landed — the same lesson the
    // shape path learned the hard way. Measure before deciding.
    console.warn("PowerChart: one-shot deck insert failed", err);
    const after = await slideCount().catch(() => before);
    added = after - before;
    if (added <= 0) return null;
  }
  if (added <= 0) return null;

  const secs = ((Date.now() - t0) / 1000).toFixed(1);
  const expected: ExpectedItem[] = items.map((it, i) => ({
    slot: i,
    title: it.title,
    // What the FILE renderer drew, not what the Office.js one would have.
    // They disagree by design: a pie is one custGeom wedge here and a
    // sixteen-triangle fan there, so `estimateOfficeShapes` measured five
    // perfect charts as wreckage on the first real run.
    shapes: built.shapesPerSlide[i] ?? estimateOfficeShapes(it.scene),
    chart: !!it.configJson,
  }));
  // Verify against the deck rather than trusting the count: the file carried
  // its own grouping and tags, so anything missing here is the host's doing.
  // Located by slot tag, not by position — where the host puts the slides is
  // its business, and on the first real run it put them at the FRONT.
  const verified = await repairDeckSpan(expected, (slot) => items[slot]?.configJson, run);
  const outcome = verified.kind === "ok" ? verified.outcome : undefined;

  let text = outcome
    ? `Inserted as one file in ${secs}s${smoke ? " (smoke subset)" : ""} — ${describeReconcile(outcome.plan)}.`
    : `Inserted ${added} of ${items.length} slides as one file in ${secs}s${smoke ? " (smoke subset)" : ""}.`;
  if (verified.kind !== "ok") text += ` (Not verified: ${verified.why}.)`;
  if (added < items.length) text += ` ⚠ the host took ${added} of ${items.length} slides.`;
  const clean = added >= items.length && !!outcome && outcome.plan.summary.lost === 0;
  return { text, status: clean ? "ok" : "err" };
}

/**
 * Shapes a web host may be asked to draw in one run before the add-in stops
 * asking. Every 12-item run in this project's history (~400 shapes) survived;
 * every 37-item one (~1850) took the whole client down. 600 sits between them,
 * nearer the side that lived.
 *
 * Only the web needs a number at all. Microsoft's documented runtime limits —
 * CPU, memory, four-crashes-per-session, five-seconds-unresponsive — are
 * scoped to Windows and Mac and explicitly NOT to a browser, so on the web
 * nothing throttles a runaway add-in: the tab dies and takes the session with
 * it. Desktop has that safety net and does not need this one.
 */
const WEB_SHAPE_BUDGET = 600;

/**
 * Rasterize, or give up. A canvas that never fires `onload` would otherwise
 * hang a run at the exact moment it is trying to rescue one.
 */
function boundedRaster(scene: Scene): Promise<string | undefined> {
  return Promise.race([
    rasterizeScene(scene),
    new Promise<undefined>((resolve) => setTimeout(() => resolve(undefined), 8_000)),
  ]).catch(() => undefined);
}

/** True when the add-in is running inside PowerPoint on the web. */
function isWebHost(): boolean {
  try {
    return Office.context?.diagnostics?.platform === Office.PlatformType.OfficeOnline;
  } catch {
    return false;
  }
}

/** Save an object as a JSON file, via the same Blob dance as Download SVG. */
function downloadJson(name: string, payload: unknown) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
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
        // Every host action in the pane comes through here, so this is the one
        // place that can log them all — inserts, updates, agenda, elements,
        // same-scale, explode, the demo deck. Normal use, not just the harness.
        const action = clicked?.id ?? "action";
        const startedAt = Date.now();
        trace("pane", "action started", { action });
        const lock = [insertBtn, clicked].filter((b): b is HTMLButtonElement => !!b && !b.disabled);
        for (const b of lock) b.disabled = true;
        note("Working…", "busy");
        // Compare against the TRANSLATED busy text: note() routes through t(), so
        // under a localized pane hostNote reads e.g. "Arbeite…" and a check against
        // the English literal never matched — the success note was never shown.
        const busyText = hostNote.textContent;
        setProgress("busy");
        startElapsed();
        try {
          await fn();
          trace("pane", "action finished", { action, ms: Date.now() - startedAt });
          if (hostNote.textContent === busyText) {
            note("Done.", "ok");
          }
        } catch (err) {
          trace("pane", "action failed", { action, ms: Date.now() - startedAt, error: errorText(err) });
          // errorText, not err.message: a RichApi.Error's message is generic
          // ("An internal error has occurred") and the useful part is in code
          // and debugInfo, which String(err) throws away.
          note("Failed: {error}", "err", { error: errorText(err) });
        } finally {
          stopElapsed();
          // Only re-enable what this call disabled — never resurrect a button
          // some other state (no host, no selection) means to keep dead.
          for (const b of lock) b.disabled = false;
        }
      };
    insertBtn.addEventListener(
      "click",
      guard(() => doInsert(false)),
    );
    insertNewBtn.addEventListener(
      "click",
      guard(() => doInsert(true)),
    );
    loadBtn.addEventListener("click", guard(doLoadSelection));
    $("selection-banner-load").addEventListener("click", guard(doLoadSelection));
    // Explode needs the same host surface as "Edit selected chart" (it reads the
    // selection's tag), so it lives or dies with the same gate.
    const explodeBtn = $("explode") as HTMLButtonElement;
    explodeBtn.disabled = false;
    explodeBtn.addEventListener("click", guard(doExplode));
    watchSelection();
    const sameScaleBtn = $("same-scale") as HTMLButtonElement;
    sameScaleBtn.disabled = false;
    sameScaleBtn.addEventListener(
      "click",
      guard(() => doSameScale("deck")),
    );
    const sameScaleSelBtn = $("same-scale-sel") as HTMLButtonElement;
    sameScaleSelBtn.disabled = false;
    sameScaleSelBtn.addEventListener(
      "click",
      guard(() => doSameScale("selection")),
    );
    const batchBtn = $("json-insert-batch") as HTMLButtonElement;
    batchBtn.disabled = false;
    batchBtn.addEventListener(
      "click",
      guard(async () => {
        const parsed = JSON.parse(($("json-io") as HTMLTextAreaElement).value);
        const configs: ChartConfig[] = Array.isArray(parsed) ? parsed : [parsed];
        let degraded = 0;
        for (const c of configs) {
          const cfg = { ...DEFAULT_SIZE, ...c };
          const scene = buildChart(cfg);
          // Batch insert honours render:"image" per config — without this a
          // pasted image-mode array silently drew native shapes.
          const { png, warn } = await chartPicture(cfg, scene);
          if (warn) degraded++;
          await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg), pictureBase64: png });
        }
        note("Inserted {n} chart(s) on the current slide.", "ok", { n: configs.length });
        if (degraded) {
          note("Inserted {n} chart(s), but {d} image chart(s) fell back to native shapes.", "err", {
            n: configs.length,
            d: degraded,
          });
        }
      }),
    );
    // Elements insert at a small default offset (they're compact shapes).
    for (const [id, scene] of [
      ["harvey-insert", harveyScene],
      ["check-insert", checkScene],
      ["flow-insert", flowScene],
      ["kpi-insert", kpiScene],
      [
        "table-insert",
        () => buildTableScene(state.sheet.cells, 480, { totalRow: ($("table-total") as HTMLInputElement).checked }),
      ],
    ] as const) {
      const btn = $(id) as HTMLButtonElement;
      btn.disabled = false;
      btn.addEventListener(
        "click",
        guard(() => insertSceneIntoSlide(scene(), { left: 120, top: 160 }, phaseNote)),
      );
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
    onLateSync((msg) => note("Host answered late — {message}", "err", { message: msg }));
    // Turning the fast path OFF on the web is a decision worth flagging. At
    // volume the shape-by-shape path there does not merely stall: the full
    // 37-item deck took the whole web client down five seconds in on
    // 2026-07-31 — "Sorry, we ran into a problem. Please try again." The
    // twelve-item subset survived, so this is a warning and not a block.
    const fileToggle = $("demo-file") as HTMLInputElement | null;
    fileToggle?.addEventListener("change", () => {
      if (fileToggle.checked || !canInsertSlidesFromBase64()) return;
      if (isWebHost()) {
        note(
          "Heads up: the full deck drawn shape by shape has crashed PowerPoint on the web. The smoke subset survives it; the fast path handles both.",
          "err",
        );
      }
    });
    // Enabled by a run, not by the host: with nothing to save it would only
    // ever produce an empty file.
    // ON by default for now. Nothing in this project has been diagnosed from
    // anything but an after-the-fact artifact, and the runs that matter happen
    // on a host nobody can attach a debugger to. When the add-in stops being
    // validated against real hosts, uncheck it in taskpane.html and drop the
    // `checked` — the module, its call sites and this toggle all keep working,
    // so a future investigation is one click away rather than a re-implementation.
    const traceToggle = $("demo-trace") as HTMLInputElement | null;
    if (traceToggle?.checked) {
      setTracing(true);
      traceEnvironment(typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev");
    }
    traceToggle?.addEventListener("change", () => {
      setTracing(traceToggle.checked);
      if (traceToggle.checked) {
        traceEnvironment(typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev");
        note("Verbose trace on — it rides along in the run log.", "ok");
      } else note("Verbose trace off.", "ok");
    });
    $("demo-log").addEventListener("click", () => {
      if (!lastRunLog) {
        note("No run to save yet — insert the demo deck first.", "err");
        return;
      }
      downloadJson("powerchart-run-log.json", lastRunLog);
      note("Run log saved.", "ok");
    });
    const demoBtn = $("demo-insert") as HTMLButtonElement;
    demoBtn.disabled = false;
    demoBtn.addEventListener(
      "click",
      guard(async () => {
        const buildStamp = typeof __BUILD_STAMP__ === "string" ? __BUILD_STAMP__ : "dev";
        const host = describeHost();
        const smoke = ($("demo-smoke") as HTMLInputElement | null)?.checked ?? false;
        const items = demoItems({ buildStamp, host, smoke });
        // The slowest thing the pane can do — say where it has got to, or a
        // multi-minute run is indistinguishable from a hang.
        // Fast path first: one generated .pptx, one host call. Falls through
        // to the shape-by-shape renderer when the host cannot take it, or when
        // the attempt landed nothing — never after a partial insert, which
        // would draw the whole deck again on top of what is already there.
        const useFile = ($("demo-file") as HTMLInputElement | null)?.checked ?? true;
        if (useFile && canInsertSlidesFromBase64()) {
          const outcome = await insertDemoDeckAsFile(items, smoke);
          if (outcome) {
            note(outcome.text, outcome.status);
            return;
          }
          note("The host would not take a generated deck — drawing it shape by shape instead.", "busy");
        }
        // NO budget exemption for the harness's own slides any more. It existed
        // because a large deck's contents page ran past the limit, and it is
        // what let a 79-shape text slide through as the SECOND thing a run
        // drew — five seconds before PowerPoint on the web crashed on
        // 2026-07-31. `buildIndexScenes` and `buildResultsScenes` now measure
        // their pages and split them instead, so there is nothing left to
        // exempt: a harness page that somehow still runs over is a page worth
        // skipping, exactly like any other.
        const {
          results,
          slidesAdded,
          addsIssued,
          blankSlides,
          blankItems,
          blanksRead,
          reconcile,
          degradedAt,
          degradeReason,
          totalMs,
        } = await insertDemoDeck(
          items.map((i) => ({
            scene: i.scene,
            tagData: i.configJson,
            title: i.title,
          })),
          (done, total) => {
            note("Inserting demo slides… {done} of {total}", "busy", { done, total });
            setProgress(done / total); // one slide per context, so a real bar
          },
          {
            // Close the run by reading the deck back and repairing it — the
            // per-item bookkeeping below is written while the host is still
            // committing, and has been observed calling a chart failed that
            // in fact landed twice.
            reconcile: true,
            // And give it somewhere to go when the host starts failing.
            // Rasterizing needs the pane's canvas, so the renderer asks and
            // this supplies; it decides when.
            // Web only. Degrading exists because a browser has no safety
            // net — Office's CPU/memory/crash-tolerance limits are scoped to
            // Windows and Mac — so on desktop, where the host throttles the
            // add-in rather than dying, drawing shapes remains the right
            // answer however long it takes.
            pictureFor: isWebHost() ? (i) => boundedRaster(items[i].scene) : undefined,
            shapeBudget: isWebHost() ? WEB_SHAPE_BUDGET : undefined,
          },
        );
        // Self-check: the deck is a regression harness, so report what the HOST
        // actually did, not what we asked for. The full table goes to the console.
        const named = (s: "skipped" | "failed") =>
          results.map((r, i) => (r.status === s ? items[i].title : "")).filter(Boolean);
        const skipped = named("skipped");
        const failedNames = named("failed");
        const rendered = results.filter((r) => r.status === "rendered").length;
        // Loss vs adds ISSUED, not vs items.length: a retry/fail stray inflates
        // slidesAdded, so measuring against items.length reads 0 during real
        // corruption when a stray cancels a lost slide. addsIssued − slidesAdded
        // counts adds that never landed (strays that landed cancel out).
        const lost = Math.max(0, addsIssued - slidesAdded);
        const secs = (totalMs / 1000).toFixed(1);
        console.log("PowerChart demo self-check:");
        console.table(
          results.map((r, i) => ({
            chart: items[i].title,
            shapes: r.created,
            status: r.status,
            grouped: !!r.grouped,
            ms: r.ms,
            lateOutcome: r.lateOutcome ?? "",
          })),
        );
        console.log(
          `deck grew by ${slidesAdded}, issued ${addsIssued} adds${lost > 0 ? ` — ${lost} LOST` : ""}; blank slots ${blankSlides.length ? blankSlides.join(", ") : "none"}${blanksRead ? "" : " (blank check incomplete)"} · total ${secs}s`,
        );
        // The headline comes from the settled read when there is one. Every
        // number above was computed while the host was still committing, and
        // the run that produced Presentation_4.pptx got three of them wrong:
        // it blamed Gantt for failing (it landed twice), called a full Agenda
        // slide blank, and counted duplicate slides as successes.
        const verdicts = reconcile?.plan.verdicts ?? [];
        const missing = verdicts.filter((v) => v.status === "lost" || v.status === "empty").map((v) => v.title);
        let msg = reconcile
          ? `Deck settled: ${describeReconcile(reconcile.plan)} — in ${secs}s${smoke ? " (smoke subset)" : ""}.`
          : `Inserted ${rendered} of ${items.length} in ${secs}s${smoke ? " (smoke subset)" : ""}.`;
        if (skipped.length) msg += ` Skipped as too dense (stamped): ${skipped.join(", ")}.`;
        if (reconcile) {
          if (missing.length) msg += ` Never landed: ${missing.join(", ")}.`;
          const { deleted, unstamped, regrouped } = reconcile.applied;
          // Deletions are NOT all duplicates — most are usually empty slides a
          // lost add left behind. Calling nine removals "9 duplicate slide(s)"
          // in the same breath as the plan's own "1 duplicate slide removed ·
          // 8 orphan slides" made the run contradict itself in one sentence.
          const dupPlanned = reconcile.plan.actions.filter((a) => a.kind === "delete" && a.slot !== null).length;
          const emptyPlanned = reconcile.plan.actions.filter((a) => a.kind === "delete" && a.slot === null).length;
          if (deleted || unstamped || regrouped)
            msg +=
              ` Repaired: removed ${deleted} slide(s)` +
              (deleted ? ` (${dupPlanned} duplicate, ${emptyPlanned} empty)` : "") +
              `, cleared ${unstamped} false banner(s), re-grouped ${regrouped} chart(s).`;
          if (reconcile.refused) msg += ` ⚠ ${reconcile.refused} repair step(s) the host refused.`;
        } else if (failedNames.length) msg += ` Host failed on: ${failedNames.join(", ")}.`;
        if (degradedAt !== undefined)
          msg += ` Drew the last ${items.length - degradedAt} slide(s) as pictures — ${degradeReason}. Use "Explode to native shapes" on any of them to get real shapes back.`;
        // A rendered but ungrouped chart is not re-editable — flag them so
        // Phase 2 doesn't quietly count them as full successes.
        const ungrouped = reconcile
          ? verdicts.filter(
              (v) => !v.tagged && v.status !== "lost" && v.status !== "skipped" && items[v.slot]?.configJson,
            ).length
          : results.filter((r, i) => r.status === "rendered" && !r.grouped && items[i].scene.nodes.length > 1).length;
        if (ungrouped) msg += ` ⚠ ${ungrouped} chart${ungrouped === 1 ? "" : "s"} landed ungrouped (not re-editable).`;
        if (lost > 0)
          msg += ` ⚠ ${lost} add${lost === 1 ? "" : "s"} did not land — the host lost slides (issued ${addsIssued}, deck grew by ${slidesAdded}).`;
        // Blank slides carry the slot tag (item title) where the host has 1.3
        // slide tags; without them the entry has title null and only its deck
        // position is shown, same as before slot tags landed.
        if (blankSlides.length) {
          const named = blankItems.map((b) => (b.title ? `${b.title} (slide ${b.position})` : `slide ${b.position}`));
          msg += ` ⚠ ${blankSlides.length} slide${blankSlides.length === 1 ? "" : "s"} came back BLANK: ${named.join(", ")}.`;
        } else if (!blanksRead) msg += ` (Blank check did not finish.)`;
        // Keep the whole run, not just the sentence. A run that ends badly is
        // the one worth reporting, and it is also the one whose results slide
        // is most likely to be the thing the host drops.
        lastRunLog = {
          build: buildStamp,
          host,
          smoke,
          totalMs,
          items: results.map((r, i) => ({
            title: items[i].title,
            status: r.status,
            shapes: r.created,
            ms: r.ms,
            grouped: !!r.grouped,
            lateOutcome: r.lateOutcome ?? "",
          })),
          deck: { slidesAdded, addsIssued, lost, blank: blankItems },
          reconcile,
          trace: tracing() ? traceLog() : undefined,
        };
        ($("demo-log") as HTMLButtonElement).disabled = false;
        // Close the deck with a self-contained results slide so the exported PDF is
        // a complete run record. A second insertDemoDeck reuses the same add/render/
        // self-check machinery; wrap it so a host stall here can't swallow the run's
        // own summary (its failure is itself just another data point).
        const rows: ResultRow[] = results.map((r, i) => ({
          title: items[i].title,
          status: r.status,
          shapes: r.created,
          ms: r.ms,
        }));
        const summary: ResultsSummary = {
          buildStamp,
          items: items.length,
          rendered,
          skipped: skipped.length,
          failed: failedNames.length,
          lost,
          totalMs,
        };
        // Each results page inserts in its OWN insertDemoDeck call. A single
        // page failing must NOT drop the rest — insertDemoDeck throws when
        // every item in its batch failed, so batching all pages together
        // meant page 2's failure lost page 1. Presentation_3.pptx surfaced
        // this: "(results slide not added)" with zero pages landed.
        let resultsPages: Scene[] = [];
        try {
          resultsPages = buildResultsScenes(rows, summary);
        } catch (e) {
          console.warn("PowerChart: results scene build failed", e);
        }
        let resultsLanded = 0;
        for (const [i, scene] of resultsPages.entries()) {
          try {
            await insertDemoDeck([
              {
                scene,
                title: resultsPages.length === 1 ? "Results" : `Results (page ${i + 1} of ${resultsPages.length})`,
              },
            ]);
            resultsLanded += 1;
          } catch (e) {
            console.warn(`PowerChart: results page ${i + 1} failed to insert`, e);
          }
        }
        if (!resultsPages.length) msg += " (results slide not added)";
        else if (resultsLanded === 0) msg += " (results slide not added)";
        else if (resultsLanded < resultsPages.length)
          msg += ` (${resultsLanded} of ${resultsPages.length} results pages added)`;
        note(msg, lost > 0 || failedNames.length || blankSlides.length ? "err" : "ok");
      }),
    );
  } else {
    insertBtn.disabled = true;
    loadBtn.disabled = true;
    ($("agenda-insert") as HTMLButtonElement).disabled = true;
    note("Not running inside PowerPoint — use Download SVG, or sideload the manifest to insert native shapes.");
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
// A shared chart link (#c=<base64 config>) reopens the exact chart. Applied
// after ?kind so an explicit link wins; malformed links are ignored silently.
if (location.hash.startsWith(CONFIG_HASH)) {
  try {
    const cfg = JSON.parse(decodeURIComponent(atob(location.hash.slice(CONFIG_HASH.length)))) as ChartConfig;
    if (cfg && typeof cfg === "object" && cfg.kind) {
      applyConfig({ ...DEFAULT_SIZE, ...cfg }, null);
      note("Chart loaded from a shared link.", "ok");
    }
  } catch {
    /* malformed share link — fall through to the default chart */
  }
}
const requestedTab = deepLink.get("tab");
if (requestedTab) {
  // Match on the dataset value rather than interpolating the parameter into a
  // selector: ?tab=chart"] made querySelector THROW at module top level, which
  // aborted the rest of boot — no build stamp, no size inputs, and no
  // wireInsert(), leaving the Insert button enabled with no click handler.
  // An unknown tab is now ignored silently, like a malformed share link.
  [...document.querySelectorAll<HTMLButtonElement>(".tabs .tab")]
    .find((el) => el.dataset.tab === requestedTab)
    ?.click();
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

const chartWInput = $("chart-w") as HTMLInputElement | null;
const chartHInput = $("chart-h") as HTMLInputElement | null;
// Seed the inputs from state, then let them WRITE state on edit (state is the
// source of truth currentConfig reads). A sub-usable value is ignored so the
// preview holds its last good size instead of snapping to the default.
if (chartWInput) chartWInput.value = String(state.width);
if (chartHInput) chartHInput.value = String(state.height);
const syncSizeFromInputs = () => {
  const w = Number(chartWInput?.value);
  const h = Number(chartHInput?.value);
  if (Number.isFinite(w) && w >= 80) state.width = w;
  if (Number.isFinite(h) && h >= 60) state.height = h;
  renderPreview();
};
chartWInput?.addEventListener("input", syncSizeFromInputs);
chartHInput?.addEventListener("input", syncSizeFromInputs);

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
