import { buildChart, DEFAULT_SIZE, valueExtent } from "../core/chart";
import { PALETTES } from "../core/style";
import type { ChartConfig, ChartKind, Decorations } from "../core/types";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import { sceneToSvg } from "../render/svg";
import {
  getSelectionBounds,
  insertAgendaSlides,
  insertSceneIntoSlide,
  isPowerPointHost,
  listChartsInDeck,
  listChartsInSelection,
  loadChartFromSelection,
  loadThemePalette,
  updateChartInSlide,
  type EditTarget,
} from "../render/powerpoint";
import { buildAgendaScene } from "../core/agenda";
import { buildCheckbox, buildHarveyBall, buildKpiTile, buildProcessFlow, buildTableScene, type CheckState } from "../core/elements";
import { localizePane } from "./i18n";
import { dataToSheet, mountDatasheet, sheetToData, type SheetModel } from "./datasheet";

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
  const style = { ...styleFile } as NonNullable<ChartConfig["style"]>;
  if (state.paletteName === "Theme" && themePalette) style.palette = themePalette;
  else if (state.paletteName !== "Default") style.palette = PALETTES[state.paletteName];
  return Object.keys(style).length ? style : undefined;
}

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

function applyConfig(cfg: ChartConfig, editTarget: EditTarget | null) {
  Object.assign(state, stateFromConfig(cfg), { editTarget });
  sheetApi.setSheet(state.sheet);
  const titleField = document.getElementById("chart-title") as HTMLInputElement | null;
  if (titleField) titleField.value = state.title;
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

function renderGallery() {
  gallery.innerHTML = "";
  for (const { kind, label } of CHART_KINDS) {
    const b = document.createElement("button");
    b.className = "thumb" + (kind === state.kind ? " active" : "");
    if (!thumbnails.has(kind)) thumbnails.set(kind, thumbnailSvg(kind));
    const pic = document.createElement("span");
    pic.className = "thumb-pic";
    pic.innerHTML = thumbnails.get(kind)!;
    const cap = document.createElement("span");
    cap.className = "thumb-cap";
    cap.textContent = label;
    b.append(pic, cap);
    b.addEventListener("click", () => applyConfig(sampleConfig(kind), null));
    gallery.appendChild(b);
  }
}

function renderOptions() {
  optionsHost.innerHTML = "";
  const d = state.decorations;
  const toggles: { key: keyof Decorations; label: string }[] = [
    { key: "segmentLabels", label: "Segment labels" },
    { key: "seriesLabels", label: "Series labels" },
    { key: "totals", label: "Column totals" },
    { key: "categoryAxis", label: "Category labels" },
    { key: "valueAxis", label: "Value axis" },
    { key: "gridlines", label: "Gridlines" },
    { key: "connectors", label: "Connector lines" },
    { key: "hundredPercentNote", label: "100% = note" },
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
    optionsHost.appendChild(label);
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
  optionsHost.appendChild(dm);

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
  optionsHost.appendChild(rot);

  const nCats = () => Math.max(0, state.sheet.cells[0].length - 1);
  optionsHost.appendChild(
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
  optionsHost.appendChild(diff);

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
  optionsHost.appendChild(vl);

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
  optionsHost.appendChild(so);

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
  optionsHost.appendChild(sc);

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
  optionsHost.appendChild(ab);

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
  optionsHost.appendChild(nf);

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
  optionsHost.appendChild(fn);

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
    optionsHost.appendChild(ex);
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
  optionsHost.appendChild(lc);

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
  optionsHost.appendChild(ax);

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
  optionsHost.appendChild(pal);

  const colors = document.createElement("div");
  colors.className = "wide series-colors";
  const palette = (state.paletteName === "Theme" && themePalette) || PALETTES[state.paletteName] || PALETTES.Default;
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
  optionsHost.appendChild(colors);
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
    await insertSceneIntoSlide(buildChart(cfg), {
      tagData: JSON.stringify(cfg),
      left: bounds.left,
      top: bounds.top,
    });
  } else {
    await insertSceneIntoSlide(buildChart(cfg), {
      tagData: JSON.stringify(cfg),
      left: 60 + insertOffset,
      top: 90 + insertOffset,
    });
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
    hostNote.textContent =
      scope === "deck"
        ? "Same scale needs at least two value-axis charts in the deck."
        : "Select two or more PowerCharts (Ctrl-click), then apply Same scale.";
    return;
  }
  const min = Math.min(...parsed.map((c) => c.extent.min));
  const max = Math.max(...parsed.map((c) => c.extent.max));
  for (const c of parsed) {
    c.cfg.scale = { min: min < 0 ? min : undefined, max };
    await updateChartInSlide(buildChart(c.cfg), c.target, { tagData: JSON.stringify(c.cfg) });
  }
  hostNote.textContent = `Same scale applied to ${parsed.length} charts (max ${max}).`;
}

async function doLoadSelection() {
  const found = await loadChartFromSelection();
  if (!found) {
    hostNote.textContent = "The selection is not a PowerChart — select an inserted chart group first.";
    return;
  }
  hostNote.textContent = "Chart loaded — edits will update it in place.";
  applyConfig(JSON.parse(found.configJson) as ChartConfig, found.target);
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
  for (const name of Object.keys(loadTemplates()).sort()) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
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
  const name = ($("template-list") as HTMLSelectElement).value;
  const cfg = loadTemplates()[name];
  if (cfg) applyConfig({ ...DEFAULT_SIZE, ...cfg }, null);
});
$("template-delete").addEventListener("click", () => {
  const name = ($("template-list") as HTMLSelectElement).value;
  if (!name) return;
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
  hostNote.textContent = "Style exported — share the JSON as your corporate style file.";
});
$("style-import").addEventListener("click", () => {
  try {
    const parsed = JSON.parse(($("json-io") as HTMLTextAreaElement).value);
    if (parsed.kind) throw new Error("that is a chart config — use Import instead");
    styleFile = parsed;
    localStorage.setItem(STYLE_KEY, JSON.stringify(styleFile));
    renderPreview();
    hostNote.textContent = "Style imported — applied to every chart from this pane.";
  } catch (err) {
    hostNote.textContent = `Style import failed: ${err instanceof Error ? err.message : String(err)}`;
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
    hostNote.textContent = Array.isArray(parsed)
      ? `Loaded chart 1 of ${parsed.length} — use "Insert batch" for all.`
      : "Chart config loaded.";
  } catch (err) {
    hostNote.textContent = `Invalid JSON: ${err instanceof Error ? err.message : String(err)}`;
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
let autoUpdateTimer: ReturnType<typeof setTimeout> | undefined;
function maybeAutoUpdate() {
  const on = ($("auto-update") as HTMLInputElement | null)?.checked;
  if (!on || !state.editTarget || !isPowerPointHost()) return;
  clearTimeout(autoUpdateTimer);
  autoUpdateTimer = setTimeout(() => void doInsert(false).catch(() => {}), 900);
}

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
    hostNote.textContent = "";
    insertBtn.disabled = false;
    loadBtn.disabled = false;
    const guard = (fn: () => Promise<void>) => async () => {
      insertBtn.disabled = true;
      hostNote.textContent = "Working…";
      hostNote.className = "hint status-busy";
      try {
        await fn();
        if (hostNote.textContent === "Working…") {
          hostNote.textContent = "Done.";
          hostNote.className = "hint status-ok";
        }
      } catch (err) {
        hostNote.textContent = `Failed: ${err instanceof Error ? err.message : String(err)}`;
        hostNote.className = "hint status-err";
      } finally {
        insertBtn.disabled = false;
      }
    };
    insertBtn.addEventListener("click", guard(() => doInsert(false)));
    insertNewBtn.addEventListener("click", guard(() => doInsert(true)));
    loadBtn.addEventListener("click", guard(doLoadSelection));
    $("selection-banner-load").addEventListener(
      "click",
      guard(async () => {
        await doLoadSelection();
        $("selection-banner").style.display = "none";
      }),
    );
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
        hostNote.textContent = `Inserted ${configs.length} chart(s) on the current slide.`;
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
      btn.addEventListener("click", guard(() => insertSceneIntoSlide(scene(), { left: 120, top: 160 })));
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
  } else {
    insertBtn.disabled = true;
    loadBtn.disabled = true;
    ($("agenda-insert") as HTMLButtonElement).disabled = true;
    hostNote.textContent =
      "Not running inside PowerPoint — use Download SVG, or sideload the manifest to insert native shapes.";
  }
}

// Ribbon deep-link: taskpane.html?kind=waterfall preselects a chart type.
const requestedKind = new URLSearchParams(location.search).get("kind");
if (requestedKind && CHART_KINDS.some((k) => k.kind === requestedKind)) {
  applyConfig(sampleConfig(requestedKind as ChartKind), null);
}

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
