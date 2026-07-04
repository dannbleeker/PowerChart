import type { ChartData } from "../core/types";
import { isTotalToken } from "../core/layout/waterfall";

export interface SheetModel {
  /** Raw cell text; row 0 = category names, column 0 = series names. */
  cells: string[][];
}

export function dataToSheet(data: ChartData): SheetModel {
  const cells: string[][] = [["", ...data.categories]];
  for (const s of data.series) {
    cells.push([s.name, ...s.values.map((v) => (v == null ? "" : String(v)))]);
  }
  return { cells };
}

export function sheetToData(sheet: SheetModel, waterfallTotals?: Set<number>): ChartData {
  const [head = [], ...rows] = sheet.cells;
  const nCats = Math.max(0, head.length - 1);
  const categories = Array.from({ length: nCats }, (_, i) => head[i + 1] ?? `Cat ${i + 1}`);
  const series = rows
    .filter((r) => r.some((c, i) => i > 0 && c.trim() !== "") || (r[0] ?? "").trim() !== "")
    .map((r, ri) => ({
      name: (r[0] ?? "").trim() || `Series ${ri + 1}`,
      values: Array.from({ length: nCats }, (_, i) => {
        const raw = (r[i + 1] ?? "").trim();
        if (raw === "") return null;
        if (waterfallTotals && isTotalToken(raw)) {
          waterfallTotals.add(i);
          return 0;
        }
        const num = Number(raw.replace(/,/g, ""));
        return Number.isFinite(num) ? num : null;
      }),
    }));
  return { categories, series };
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
        input.addEventListener("paste", (e) => handlePaste(e, ri, ci));
        td.appendChild(input);
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });
    host.appendChild(table);

    const controls = document.createElement("div");
    controls.className = "sheet-controls";
    controls.append(
      button("+ Row", () => {
        model.cells.push(model.cells[0].map(() => ""));
        render();
        onChange(model);
      }),
      button("+ Column", () => {
        model.cells.forEach((r) => r.push(""));
        render();
        onChange(model);
      }),
      button("− Row", () => {
        if (model.cells.length > 2) model.cells.pop();
        render();
        onChange(model);
      }),
      button("− Column", () => {
        if (model.cells[0].length > 2) model.cells.forEach((r) => r.pop());
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
    const rows = text.replace(/\r/g, "").split("\n").filter((r) => r.length);
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
