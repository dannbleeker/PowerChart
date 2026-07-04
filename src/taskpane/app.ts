import { buildChart, DEFAULT_SIZE, valueExtent } from "../core/chart";
import type { ChartConfig, ChartKind, Decorations } from "../core/types";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import { sceneToSvg } from "../render/svg";
import {
  insertAgendaSlides,
  insertSceneIntoSlide,
  isPowerPointHost,
  listChartsInDeck,
  loadChartFromSelection,
  updateChartInSlide,
  type EditTarget,
} from "../render/powerpoint";
import { buildAgendaScene } from "../core/agenda";
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
  };
}

function currentConfig(): ChartConfig {
  const totals = new Set<number>();
  const data = sheetToData(state.sheet, state.kind === "waterfall" ? totals : undefined);
  const min = state.scaleMin.trim() === "" ? undefined : Number(state.scaleMin);
  const max = state.scaleMax.trim() === "" ? undefined : Number(state.scaleMax);
  const bFrom = Number(state.breakFrom);
  const bTo = Number(state.breakTo);
  const axisBreak =
    state.breakFrom.trim() && state.breakTo.trim() && Number.isFinite(bFrom) && Number.isFinite(bTo) && bTo > bFrom
      ? { from: bFrom, to: bTo }
      : undefined;
  return {
    kind: state.kind,
    data,
    horizontal: state.horizontal || undefined,
    ...DEFAULT_SIZE,
    title: state.title || undefined,
    decorations: state.decorations,
    waterfall: { totalIndices: [...totals] },
    segmentOrder: state.segmentOrder === "sheet" ? undefined : state.segmentOrder,
    axisBreak,
    scale:
      (min != null && Number.isFinite(min)) || (max != null && Number.isFinite(max))
        ? { min: Number.isFinite(min!) ? min : undefined, max: Number.isFinite(max!) ? max : undefined }
        : undefined,
    numberFormat:
      state.decimals !== "auto" || state.suffix
        ? {
            decimals: state.decimals === "auto" ? "auto" : Number(state.decimals),
            suffix: state.suffix || undefined,
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
  nf.append("Labels: decimals ", nfDec, " suffix ", nfSuffix);
  optionsHost.appendChild(nf);
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
}

function renderActionState() {
  const insertBtn = $("insert") as HTMLButtonElement;
  insertBtn.textContent = state.editTarget ? "Update chart" : "Insert into slide";
  ($("insert-new") as HTMLButtonElement).style.display = state.editTarget ? "" : "none";
}

// --- boot ------------------------------------------------------------------

sheetApi = mountDatasheet($("datasheet"), state.sheet, (sheet) => {
  state.sheet = sheet;
  renderPreview();
});

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

async function doInsert(asNew: boolean) {
  const cfg = currentConfig();
  const scene = buildChart(cfg);
  const tagData = JSON.stringify(cfg);
  if (!asNew && state.editTarget) {
    await updateChartInSlide(scene, state.editTarget, { tagData });
  } else {
    await insertSceneIntoSlide(scene, { tagData });
    state.editTarget = null;
    renderActionState();
  }
}

/**
 * think-cell's Set Same Scale: pin every value-axis chart in the deck to the
 * union of their extents and re-render them in place.
 */
async function doSameScale() {
  const charts = await listChartsInDeck();
  const parsed = charts
    .map((c) => ({ target: c.target, cfg: JSON.parse(c.configJson) as ChartConfig }))
    .map((c) => ({ ...c, extent: valueExtent(c.cfg) }))
    .filter((c): c is typeof c & { extent: { min: number; max: number } } => c.extent != null);
  if (parsed.length < 2) {
    hostNote.textContent = "Same scale needs at least two value-axis charts in the deck.";
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
      try {
        await fn();
      } catch (err) {
        hostNote.textContent = `Failed: ${err instanceof Error ? err.message : String(err)}`;
      } finally {
        insertBtn.disabled = false;
      }
    };
    insertBtn.addEventListener("click", guard(() => doInsert(false)));
    insertNewBtn.addEventListener("click", guard(() => doInsert(true)));
    loadBtn.addEventListener("click", guard(doLoadSelection));
    const sameScaleBtn = $("same-scale") as HTMLButtonElement;
    sameScaleBtn.disabled = false;
    sameScaleBtn.addEventListener("click", guard(doSameScale));
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

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => wireInsert());
} else {
  wireInsert();
}
