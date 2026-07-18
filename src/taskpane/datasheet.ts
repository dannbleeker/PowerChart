import type { ChartData } from "../core/types";
import { isTotalToken } from "../core/layout/waterfall";
import { parseDateToken } from "../core/format";

export interface SheetModel {
  /** Raw cell text; row 0 = category names, column 0 = series names. */
  cells: string[][];
}

// --- Formulas ----------------------------------------------------------------

/** Column letters → 0-based index (A=0, Z=25, AA=26 …). */
function colIndex(letters: string): number {
  let v = 0;
  for (const ch of letters.toUpperCase()) v = v * 26 + (ch.charCodeAt(0) - 64);
  return v - 1;
}

/** Numeric value of a cell, following "=" formulas with cycle protection. */
function cellNumeric(cells: string[][], row: number, col: number, visiting: Set<string>): number {
  const key = `${row},${col}`;
  if (visiting.has(key)) return NaN; // circular reference
  const raw = (cells[row]?.[col] ?? "").trim();
  if (raw === "") return 0;
  if (raw.startsWith("=")) {
    visiting.add(key);
    const v = evaluateFormula(cells, raw.slice(1), visiting);
    visiting.delete(key);
    return v ?? NaN;
  }
  const n = Number(raw.replace(/,/g, ""));
  // A non-numeric NON-blank cell (text, an error token) is not a value — return
  // NaN so it propagates as an error, the same stance parseRow takes when it
  // returns null for the same cell in place. (A BLANK cell stays 0 above, so SUM
  // still treats gaps as zero — Excel's convention.) The old silent 0 meant a
  // stray "n/a" in a referenced cell vanished into a computed total.
  return Number.isFinite(n) ? n : NaN;
}

/**
 * Evaluate a spreadsheet formula against the datasheet: A1-style refs
 * (row 1 = the category header row), + - * / ( ), and SUM/AVG/MIN/MAX
 * over ranges like B2:E2. Returns null on parse errors or cycles.
 */
export function evaluateFormula(cells: string[][], expr: string, visiting: Set<string> = new Set()): number | null {
  const s = expr.replace(/\s+/g, "");
  let i = 0;

  const ref = () => {
    const m = /^([A-Za-z]{1,2})([0-9]{1,3})/.exec(s.slice(i));
    if (!m) return null;
    i += m[0].length;
    return cellNumeric(cells, Number(m[2]) - 1, colIndex(m[1]), visiting);
  };
  const rangeValues = (m: RegExpExecArray): number[] => {
    const c1 = colIndex(m[1]);
    const r1 = Number(m[2]) - 1;
    const c2 = colIndex(m[3]);
    const r2 = Number(m[4]) - 1;
    const out: number[] = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.push(cellNumeric(cells, r, c, visiting));
    return out;
  };

  function expr0(): number {
    let v = term();
    while (s[i] === "+" || s[i] === "-") {
      const op = s[i++];
      const r = term();
      v = op === "+" ? v + r : v - r;
    }
    return v;
  }
  function term(): number {
    let v = factor();
    while (s[i] === "*" || s[i] === "/") {
      const op = s[i++];
      const r = factor();
      v = op === "*" ? v * r : v / r;
    }
    return v;
  }
  function factor(): number {
    if (s[i] === "-") {
      i++;
      return -factor();
    }
    if (s[i] === "(") {
      i++;
      const v = expr0();
      if (s[i] === ")") i++;
      return v;
    }
    const fn = /^(SUM|AVG|MIN|MAX)\(/i.exec(s.slice(i));
    if (fn) {
      i += fn[0].length;
      const args: number[] = [];
      while (i < s.length && s[i] !== ")") {
        const range = /^([A-Za-z]{1,2})([0-9]{1,3}):([A-Za-z]{1,2})([0-9]{1,3})/.exec(s.slice(i));
        if (range) {
          i += range[0].length;
          args.push(...rangeValues(range));
        } else {
          args.push(expr0());
        }
        if (s[i] === ",") i++;
      }
      i++;
      if (!args.length) return NaN;
      const name = fn[1].toUpperCase();
      if (name === "SUM") return args.reduce((a, b) => a + b, 0);
      if (name === "AVG") return args.reduce((a, b) => a + b, 0) / args.length;
      if (name === "MIN") return Math.min(...args);
      return Math.max(...args);
    }
    const num = /^[0-9]*\.?[0-9]+/.exec(s.slice(i));
    if (num) {
      i += num[0].length;
      return Number(num[0]);
    }
    const r = ref();
    if (r != null) return r;
    i++;
    return NaN;
  }

  const result = expr0();
  return i >= s.length && Number.isFinite(result) ? result : null;
}

/** think-cell's special datasheet rows, matched by row name. */
const HUNDRED_ROW = /^100\s*%\s*=?$/i;
const XEXTENT_ROW = /^x(\s*extent)?$/i;

const GANTT_DATE_ROW = /^(start|end|milestone)$/i;

export function dataToSheet(data: ChartData): SheetModel {
  const cells: string[][] = [["", ...data.categories]];
  const numRow = (name: string, values: (number | null)[]) => [
    name,
    ...data.categories.map((_, i) => (values[i] == null ? "" : String(values[i]))),
  ];
  if (data.hundredPercent) cells.push(numRow("100%=", data.hundredPercent));
  if (data.xExtent) cells.push(numRow("X extent", data.xExtent));
  let prevStack: number | undefined;
  for (const s of data.series) {
    // Blank separator row between stack groups (clustered-stacked round trip).
    if (s.stack != null && prevStack != null && s.stack !== prevStack) {
      cells.push(Array.from({ length: data.categories.length + 1 }, () => ""));
    }
    prevStack = s.stack;
    // Calendar Gantt round trip: show epoch-day values as ISO dates again.
    const asDate = data.dates && GANTT_DATE_ROW.test(s.name.trim());
    cells.push([
      s.name,
      ...s.values.map((v) => (v == null ? "" : asDate ? new Date(v * 86400000).toISOString().slice(0, 10) : String(v))),
    ]);
  }
  return { cells };
}

export function sheetToData(sheet: SheetModel, waterfallTotals?: Set<number>): ChartData {
  const [head = [], ...rows] = sheet.cells;
  const nCats = Math.max(0, head.length - 1);
  const categories = Array.from({ length: nCats }, (_, i) => head[i + 1] ?? `Cat ${i + 1}`);
  let sawDate = false;
  const parseRow = (r: string[], catIdxTotals?: Set<number>) =>
    Array.from({ length: nCats }, (_, i) => {
      const raw = (r[i + 1] ?? "").trim();
      if (raw === "") return null;
      if (catIdxTotals && isTotalToken(raw)) {
        catIdxTotals.add(i);
        return 0;
      }
      if (raw.startsWith("=") && raw.length > 1) {
        return evaluateFormula(sheet.cells, raw.slice(1));
      }
      const num = Number(raw.replace(/,/g, ""));
      if (Number.isFinite(num)) return num;
      // Calendar dates (for Gantt timelines) become days since the epoch.
      const day = parseDateToken(raw);
      if (day != null) {
        sawDate = true;
        return day;
      }
      return null;
    });

  let hundredPercent: (number | null)[] | undefined;
  let xExtent: (number | null)[] | undefined;
  const series: ChartData["series"] = [];
  // Blank rows split stacks (think-cell's clustered-stacked convention).
  let stack = 0;
  let usedStacks = false;
  rows.forEach((r) => {
    const blank = r.every((c) => c.trim() === "");
    if (blank) {
      if (series.length) {
        stack++;
        usedStacks = true;
      }
      return;
    }
    const name = (r[0] ?? "").trim();
    if (HUNDRED_ROW.test(name)) {
      hundredPercent = parseRow(r);
    } else if (XEXTENT_ROW.test(name)) {
      xExtent = parseRow(r);
    } else {
      series.push({ name: name || `Series ${series.length + 1}`, values: parseRow(r, waterfallTotals), stack });
    }
  });
  if (!usedStacks) for (const s of series) delete s.stack;
  return { categories, series, hundredPercent, xExtent, dates: sawDate || undefined };
}

/** Swap rows and columns (think-cell's Transpose): series become categories. */
export function transposeSheet(sheet: SheetModel): SheetModel {
  const rows = sheet.cells.length;
  const cols = Math.max(0, ...sheet.cells.map((r) => r.length));
  const cells = Array.from({ length: cols }, (_, c) => Array.from({ length: rows }, (_, r) => sheet.cells[r][c] ?? ""));
  return { cells };
}

/**
 * Editable datasheet grid: an HTML table of inputs with Excel-style TSV paste.
 * Calls onChange with the raw sheet on every edit.
 */
export function mountDatasheet(
  host: HTMLElement,
  sheet: SheetModel,
  onChange: (sheet: SheetModel) => void,
): { setSheet(next: SheetModel): void } {
  let model = sheet;
  /** Last focused cell, so row/column operations act at the cursor. */
  let cursor = { row: 1, col: 1 };

  function focusCell(row: number, col: number) {
    const el = host.querySelector<HTMLInputElement>(`input[data-row="${row}"][data-col="${col}"]`);
    if (el) {
      el.focus();
      el.select();
    }
  }

  /** Excel-style navigation: Enter ↓, arrows move at the text boundaries. */
  function handleNav(e: KeyboardEvent, ri: number, ci: number, input: HTMLInputElement) {
    const atStart = input.selectionStart === 0 && input.selectionEnd === 0;
    const atEnd = input.selectionStart === input.value.length && input.selectionEnd === input.value.length;
    const move = (r: number, c: number) => {
      e.preventDefault();
      focusCell(r, c);
    };
    if (e.key === "Enter" || e.key === "ArrowDown") {
      if (ri < model.cells.length - 1) move(ri + 1, ci);
    } else if (e.key === "ArrowUp") {
      if (ri > 0) move(ri - 1, ci);
    } else if (e.key === "ArrowRight" && atEnd) {
      if (ci < model.cells[0].length - 1) move(ri, ci + 1);
    } else if (e.key === "ArrowLeft" && atStart) {
      if (ci > 0) move(ri, ci - 1);
    }
  }

  function render() {
    host.innerHTML = "";
    const table = document.createElement("table");
    table.className = "datasheet";
    model.cells.forEach((row, ri) => {
      const tr = document.createElement("tr");
      row.forEach((cell, ci) => {
        const td = document.createElement("td");
        const input = document.createElement("input");
        input.value = cell;
        input.dataset.row = String(ri);
        input.dataset.col = String(ci);
        if (ri === 0 || ci === 0) input.classList.add("header");
        if (ri === 0 && ci === 0) input.disabled = true;
        input.addEventListener("input", () => {
          model.cells[ri][ci] = input.value;
          onChange(model);
        });
        input.addEventListener("focus", () => {
          cursor = { row: ri, col: ci };
        });
        input.addEventListener("keydown", (e) => handleNav(e, ri, ci, input));
        input.addEventListener("paste", (e) => handlePaste(e, ri, ci));
        td.appendChild(input);
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });
    host.appendChild(table);

    const controls = document.createElement("div");
    controls.className = "sheet-controls";
    // Row/column operations act at the cursor (insert after / delete at).
    controls.append(
      button("+ Row", () => {
        const at = Math.min(cursor.row + 1, model.cells.length);
        model.cells.splice(
          at,
          0,
          model.cells[0].map(() => ""),
        );
        render();
        onChange(model);
      }),
      button("+ Column", () => {
        const at = Math.min(cursor.col + 1, model.cells[0].length);
        model.cells.forEach((r) => r.splice(at, 0, ""));
        render();
        onChange(model);
      }),
      button("− Row", () => {
        if (model.cells.length > 2 && cursor.row > 0) {
          model.cells.splice(Math.min(cursor.row, model.cells.length - 1), 1);
          render();
          onChange(model);
        }
      }),
      button("− Column", () => {
        if (model.cells[0].length > 2 && cursor.col > 0) {
          const at = Math.min(cursor.col, model.cells[0].length - 1);
          model.cells.forEach((r) => r.splice(at, 1));
          render();
          onChange(model);
        }
      }),
      button("⇄ Transpose", () => {
        model = transposeSheet(model);
        render();
        onChange(model);
      }),
    );
    host.appendChild(controls);
  }

  function handlePaste(e: ClipboardEvent, ri: number, ci: number) {
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (!text.includes("\t") && !text.includes("\n")) return; // single cell — default behavior
    e.preventDefault();
    const rows = text
      .replace(/\r/g, "")
      .split("\n")
      .filter((r) => r.length);
    rows.forEach((row, dr) => {
      row.split("\t").forEach((val, dc) => {
        const r = ri + dr;
        const c = ci + dc;
        while (model.cells.length <= r) model.cells.push(model.cells[0].map(() => ""));
        model.cells.forEach((mr) => {
          while (mr.length <= c) mr.push("");
        });
        model.cells[r][c] = val;
      });
    });
    render();
    onChange(model);
  }

  function button(label: string, onClick: () => void): HTMLButtonElement {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.addEventListener("click", onClick);
    return b;
  }

  render();
  return {
    setSheet(next: SheetModel) {
      model = next;
      render();
    },
  };
}
