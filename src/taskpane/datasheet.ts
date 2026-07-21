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

/** 0-based index → column letters (0=A, 25=Z, 26=AA …), colIndex's inverse. */
function colLetters(index: number): string {
  let s = "";
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) s = String.fromCharCode(65 + (i % 26)) + s;
  return s;
}

/**
 * A cell's number, accepting the US thousands grouping Excel copies ("1,234").
 * A comma ANYWHERE ELSE is not a separator we can read: stripping every comma
 * turned a European "1.234,5" into 1.2345 (a silent 1000× error) and "1,5" into
 * 15. Refuse those instead — a visible gap beats a wrong number.
 */
function numericValue(text: string): number {
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(text)) return Number(text.replace(/,/g, ""));
  return text.includes(",") ? NaN : Number(text);
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
  const n = numericValue(raw);
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
  // A blank cell in a range is `null`, not 0: Excel's MIN/MAX/AVG ignore empty
  // cells (only SUM treats them as 0, which it still does below). Distinguish a
  // blank from a real 0 by looking at the raw cell before cellNumeric coerces it.
  const rangeValues = (m: RegExpExecArray): (number | null)[] => {
    const c1 = colIndex(m[1]);
    const r1 = Number(m[2]) - 1;
    const c2 = colIndex(m[3]);
    const r2 = Number(m[4]) - 1;
    const out: (number | null)[] = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++)
        out.push((cells[r]?.[c] ?? "").trim() === "" ? null : cellNumeric(cells, r, c, visiting));
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
      const args: (number | null)[] = [];
      while (i < s.length && s[i] !== ")") {
        const range = /^([A-Za-z]{1,2})([0-9]{1,3}):([A-Za-z]{1,2})([0-9]{1,3})/.exec(s.slice(i));
        if (range) {
          i += range[0].length;
          args.push(...rangeValues(range));
        } else {
          // A BARE cell reference must contribute the cell's EMPTINESS, not
          // expr0()'s 0 — otherwise MIN/MAX/AVG over comma-separated args counted a
          // blank as a real 0 while the range form correctly ignores it, so
          // =MIN(B2,C2,D2) returned 0 where =MIN(B2:D2) returned 10.
          const one = /^([A-Za-z]{1,2})([0-9]{1,3})\s*(?=[,)])/.exec(s.slice(i));
          if (one) {
            i += one[0].length;
            const rr = Number(one[2]) - 1;
            const cc = colIndex(one[1]);
            args.push((cells[rr]?.[cc] ?? "").trim() === "" ? null : cellNumeric(cells, rr, cc, visiting));
          } else {
            args.push(expr0());
          }
        }
        if (s[i] === ",") i++;
      }
      i++;
      if (!args.length) return NaN;
      const name = fn[1].toUpperCase();
      // SUM counts a blank range cell as 0 (Excel convention); MIN/MAX/AVG ignore
      // it — an all-negative range with a gap must not report a max of 0.
      if (name === "SUM") return args.reduce((a: number, b) => a + (b ?? 0), 0);
      const nums = args.filter((v): v is number => v != null);
      if (!nums.length) return NaN;
      if (name === "AVG") return nums.reduce((a, b) => a + b, 0) / nums.length;
      if (name === "MIN") return Math.min(...nums);
      return Math.max(...nums);
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
// "extent" is REQUIRED: scatter/bubble use a series literally named "X"
// (layout/scatter.ts matches /^x$/i), and a lenient pattern here silently ate it
// into the Mekko-only xExtent field on every pane round-trip. dataToSheet only
// ever writes the full "X extent", so nothing legitimate depends on the short form.
const XEXTENT_ROW = /^x\s*extent$/i;

// Every Gantt row whose values are calendar positions — the same set layoutGantt
// feeds through its time scale (core/layout/gantt.ts). With only start/end/
// milestone here, re-opening your own calendar Gantt showed the raw epoch day
// ("20494") where you had typed "2026-02-10", so the row could only be edited in
// epoch days. "% complete" and "After" are NOT dates and stay out.
const GANTT_DATE_ROW = /^(?:start|end|milestone|today|holidays?|baseline\s*(?:start|end))$|^bracket\b/i;

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
      // Strip a trailing percent sign the same way the thousands separator is
      // stripped. Excel copies a percent-formatted cell to the clipboard as its
      // DISPLAYED text ("35%"), so pasting a share table — the canonical source
      // for a 100%/stacked chart — arrives percent-suffixed. Read "35%" as 35,
      // matching how a think-cell datasheet holds a share (and how dataToSheet
      // writes it back). Without this the value is dropped to a blank gap.
      const num = numericValue(raw.replace(/\s*%$/, ""));
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

/**
 * Swap the row and column of every A1 reference in a formula. A cell at (r,c)
 * lands at (c,r) after a transpose, so each of its references has to make the
 * same move — carrying "=SUM(B2:B3)" over verbatim left it summing a different
 * set of cells, silently changing the very numbers the transpose preserves.
 */
function transposeFormula(text: string): string {
  return text.replace(/([A-Za-z]{1,2})([0-9]{1,3})/g, (_m, letters: string, digits: string) => {
    const c = colIndex(letters);
    const r = Number(digits) - 1;
    return `${colLetters(r)}${c + 1}`;
  });
}

/** Swap rows and columns (think-cell's Transpose): series become categories. */
export function transposeSheet(sheet: SheetModel): SheetModel {
  const rows = sheet.cells.length;
  const cols = Math.max(0, ...sheet.cells.map((r) => r.length));
  const cells = Array.from({ length: cols }, (_, c) =>
    Array.from({ length: rows }, (_, r) => {
      const v = sheet.cells[r][c] ?? "";
      return v.startsWith("=") ? transposeFormula(v) : v;
    }),
  );
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
        // Seed the name cell: a FULLY blank row is sheetToData's stack separator,
        // so "+ Row" used to split a stacked chart into two stacks before the
        // user had typed anything. A placeholder name (unique, so it cannot
        // collide with an existing series) reads as a new empty series instead.
        const used = new Set(model.cells.map((r) => (r[0] ?? "").trim()));
        let n = at;
        while (used.has(`Series ${n}`)) n++;
        model.cells.splice(
          at,
          0,
          model.cells[0].map((_, i) => (i === 0 ? `Series ${n}` : "")),
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
