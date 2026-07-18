/**
 * Excel companion pane: turns the selected range into a PowerChart JSON
 * config (same datasheet convention: row 1 = categories, column A = series).
 * Users paste it into PowerChart's Automation box in PowerPoint and re-run
 * whenever the data changes — the feasible substitute for live data links.
 */
import type { ChartConfig, ChartKind } from "../core/types";
import { sheetToData, transposeSheet } from "../taskpane/datasheet";
import { DEFAULT_SIZE } from "../core/chart";

/* global Excel, Office */

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

function rangeToConfig(values: unknown[][], kind: ChartKind, title: string, transpose: boolean): ChartConfig {
  const cells = values.map((row) => row.map((v) => (v == null ? "" : String(v))));
  // The datasheet convention is row 1 = categories, column A = series. A user
  // whose sheet is laid out the other way (series across the top) would silently
  // get a transposed chart — the transpose toggle swaps axes before parsing.
  const sheet = transpose ? transposeSheet({ cells }) : { cells };
  const totals = new Set<number>();
  const data = sheetToData(sheet, kind === "waterfall" ? totals : undefined);
  return {
    kind,
    data,
    ...DEFAULT_SIZE,
    title: title || undefined,
    waterfall: kind === "waterfall" ? { totalIndices: [...totals] } : undefined,
  };
}

async function generate() {
  const note = $("note");
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values,address");
      await context.sync();
      const cfg = rangeToConfig(
        range.values as unknown[][],
        ($("kind") as HTMLSelectElement).value as ChartKind,
        ($("title") as HTMLInputElement).value,
        ($("transpose") as HTMLInputElement | null)?.checked ?? false,
      );
      ($("output") as HTMLTextAreaElement).value = JSON.stringify(cfg, null, 2);
      note.textContent = `Generated from ${range.address}. Paste into PowerChart → Automation → Import.`;
    });
  } catch (err) {
    note.textContent = `Failed: ${err instanceof Error ? err.message : String(err)}`;
  }
}

function wire() {
  const inExcel = typeof Excel !== "undefined" && !!Office.context?.host;
  ($("generate") as HTMLButtonElement).disabled = !inExcel;
  if (!inExcel) {
    $("note").textContent = "Not running inside Excel — sideload manifest-excel.xml to use the data bridge.";
  }
  $("generate").addEventListener("click", () => void generate());
  $("copy").addEventListener("click", () => {
    const out = ($("output") as HTMLTextAreaElement).value;
    if (out) void navigator.clipboard?.writeText(out).catch(() => {});
  });
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => wire());
} else {
  wire();
}
