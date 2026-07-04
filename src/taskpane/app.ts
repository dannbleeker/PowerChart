import { buildChart, DEFAULT_SIZE } from "../core/chart";
import type { ChartConfig, ChartKind, Decorations } from "../core/types";
import { CHART_KINDS, sampleConfig } from "../core/samples";
import { sceneToSvg } from "../render/svg";
import { insertSceneIntoSlide, isPowerPointHost } from "../render/powerpoint";
import { dataToSheet, mountDatasheet, sheetToData, type SheetModel } from "./datasheet";

interface AppState {
  kind: ChartKind;
  sheet: SheetModel;
  decorations: Partial<Decorations>;
  waterfallTotals: Set<number>;
  title: string;
}

const state: AppState = initFromSample("stacked");

function initFromSample(kind: ChartKind): AppState {
  const cfg = sampleConfig(kind);
  const totals = new Set(cfg.waterfall?.totalIndices ?? []);
  const sheet = dataToSheet(cfg.data);
  if (kind === "waterfall") {
    // Show "e" tokens in the sheet where totals are computed.
    for (const i of totals) sheet.cells[1][i + 1] = "e";
  }
  return { kind, sheet, decorations: { ...cfg.decorations }, waterfallTotals: totals, title: cfg.title ?? "" };
}

function currentConfig(): ChartConfig {
  const totals = new Set<number>();
  const data = sheetToData(state.sheet, state.kind === "waterfall" ? totals : undefined);
  state.waterfallTotals = totals;
  return {
    kind: state.kind,
    data,
    ...DEFAULT_SIZE,
    title: state.title || undefined,
    decorations: state.decorations,
    waterfall: { totalIndices: [...totals] },
  };
}

// --- UI wiring ------------------------------------------------------------

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const gallery = $("gallery");
const preview = $("preview");
const optionsHost = $("options");

let sheetApi: { setSheet(next: SheetModel): void };

function renderGallery() {
  gallery.innerHTML = "";
  for (const { kind, label } of CHART_KINDS) {
    const b = document.createElement("button");
    b.textContent = label;
    b.className = kind === state.kind ? "active" : "";
    b.addEventListener("click", () => {
      const next = initFromSample(kind);
      state.kind = next.kind;
      state.sheet = next.sheet;
      state.decorations = next.decorations;
      state.title = next.title;
      sheetApi.setSheet(state.sheet);
      renderGallery();
      renderOptions();
      renderPreview();
    });
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

  const nCats = () => Math.max(0, state.sheet.cells[0].length - 1);
  optionsHost.appendChild(
    pairControl("CAGR arrow", d.cagr, nCats(), (pair) => {
      d.cagr = pair;
      renderPreview();
    }),
  );
  optionsHost.appendChild(
    pairControl("Difference arrow", d.difference, nCats(), (pair) => {
      d.difference = pair;
      renderPreview();
    }),
  );

  const vl = document.createElement("label");
  vl.className = "wide";
  const vlCb = document.createElement("input");
  vlCb.type = "checkbox";
  vlCb.checked = !!d.valueLine;
  vlCb.addEventListener("change", () => {
    d.valueLine = vlCb.checked ? { mode: "mean" } : undefined;
    renderPreview();
  });
  vl.append(vlCb, "Value line (mean of totals)");
  optionsHost.appendChild(vl);
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
  const from = document.createElement("input");
  from.type = "number";
  from.min = "1";
  from.value = String((current?.from ?? 0) + 1);
  const to = document.createElement("input");
  to.type = "number";
  to.min = "1";
  to.value = String((current?.to ?? Math.max(0, nCats - 1)) + 1);
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

// --- boot ------------------------------------------------------------------

sheetApi = mountDatasheet($("datasheet"), state.sheet, () => renderPreview());
renderGallery();
renderOptions();
renderPreview();

$("download").addEventListener("click", () => {
  const svg = sceneToSvg(buildChart(currentConfig()), { background: "#ffffff" });
  const blob = new Blob([svg], { type: "image/svg+xml" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "powerchart.svg";
  a.click();
  URL.revokeObjectURL(a.href);
});

const insertBtn = $("insert") as HTMLButtonElement;
const hostNote = $("host-note");

function wireInsert() {
  if (isPowerPointHost()) {
    hostNote.textContent = "";
    insertBtn.disabled = false;
    insertBtn.addEventListener("click", async () => {
      insertBtn.disabled = true;
      try {
        const cfg = currentConfig();
        const scene = buildChart(cfg);
        await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg) });
      } catch (err) {
        hostNote.textContent = `Insert failed: ${err instanceof Error ? err.message : String(err)}`;
      } finally {
        insertBtn.disabled = false;
      }
    });
  } else {
    insertBtn.disabled = true;
    hostNote.textContent = "Not running inside PowerPoint — use Download SVG, or sideload the manifest to insert native shapes.";
  }
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => wireInsert());
} else {
  wireInsert();
}
