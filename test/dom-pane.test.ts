// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { mountDatasheet, sheetToData, type SheetModel } from "../src/taskpane/datasheet";
import { localizePane, localizeTree } from "../src/taskpane/i18n";

const sheet = (): SheetModel => ({
  cells: [
    ["", "A", "B"],
    ["S1", "1", "2"],
    ["S2", "3", "4"],
  ],
});

const inputs = (host: HTMLElement) => [...host.querySelectorAll<HTMLInputElement>("input")];
const cell = (host: HTMLElement, r: number, c: number) =>
  host.querySelector<HTMLInputElement>(`input[data-row="${r}"][data-col="${c}"]`)!;
const clickButton = (host: HTMLElement, label: string) =>
  [...host.querySelectorAll<HTMLButtonElement>(".sheet-controls button")].find((b) => b.textContent === label)!.click();

describe("mountDatasheet", () => {
  it("renders a grid with headers and a disabled corner", () => {
    const host = document.createElement("div");
    mountDatasheet(host, sheet(), () => {});
    expect(inputs(host)).toHaveLength(9);
    expect(cell(host, 0, 0).disabled).toBe(true);
    expect(cell(host, 0, 1).classList.contains("header")).toBe(true);
    expect(cell(host, 1, 1).classList.contains("header")).toBe(false);
  });

  it("propagates edits through onChange", () => {
    const host = document.createElement("div");
    const onChange = vi.fn();
    mountDatasheet(host, sheet(), onChange);
    const input = cell(host, 1, 1);
    input.value = "42";
    input.dispatchEvent(new Event("input"));
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0][0].cells[1][1]).toBe("42");
  });

  it("navigates with Enter and arrow keys at text boundaries", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    mountDatasheet(host, sheet(), () => {});
    const start = cell(host, 1, 1);
    start.focus();
    start.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    expect(document.activeElement).toBe(cell(host, 2, 1));
    cell(host, 2, 1).dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowUp", bubbles: true }));
    expect(document.activeElement).toBe(cell(host, 1, 1));
    // Caret sits at the end after focusCell's select() → ArrowRight moves cell.
    const cur = cell(host, 1, 1);
    cur.setSelectionRange(cur.value.length, cur.value.length);
    cur.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    expect(document.activeElement).toBe(cell(host, 1, 2));
    const right = cell(host, 1, 2);
    right.setSelectionRange(0, 0);
    right.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowLeft", bubbles: true }));
    expect(document.activeElement).toBe(cell(host, 1, 1));
    host.remove();
  });

  it("inserts and deletes rows/columns at the cursor", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    let model = sheet();
    mountDatasheet(host, model, (m) => (model = m));
    cell(host, 1, 1).focus();
    clickButton(host, "+ Row");
    expect(model.cells).toHaveLength(4);
    // The name cell is seeded (see the stack-separator test below); the values
    // are blank.
    expect(model.cells[2]).toEqual(["Series 2", "", ""]);
    clickButton(host, "− Row");
    expect(model.cells).toHaveLength(3);
    cell(host, 1, 1).focus();
    clickButton(host, "+ Column");
    expect(model.cells[0]).toEqual(["", "A", "", "B"]);
    // Deletion acts at the cursor: focus the inserted blank column first.
    cell(host, 1, 2).focus();
    clickButton(host, "− Column");
    expect(model.cells[0]).toEqual(["", "A", "B"]);
    host.remove();
  });

  it("names the row it adds, so it is not read as a stack separator", () => {
    // A FULLY blank row is sheetToData's clustered-stacked separator, so "+ Row"
    // used to split a stacked chart into two stacks — restructuring the preview
    // — before the user had typed anything. Nothing in the pane documents blank
    // rows as stack breaks, so the button must not create one by itself.
    const host = document.createElement("div");
    document.body.appendChild(host);
    let model = sheet();
    mountDatasheet(host, model, (m) => (model = m));
    cell(host, 1, 1).focus();
    clickButton(host, "+ Row");
    const series = sheetToData(model).series;
    expect(series.map((s) => s.name)).toEqual(["S1", "Series 2", "S2"]);
    expect(series.every((s) => s.stack === undefined)).toBe(true);
    host.remove();
  });

  it("refuses to delete below the 2×2 minimum or the header row/column", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const model: SheetModel = {
      cells: [
        ["", "A"],
        ["S1", "1"],
      ],
    };
    const onChange = vi.fn();
    mountDatasheet(host, model, onChange);
    cell(host, 1, 1).focus();
    clickButton(host, "− Row");
    clickButton(host, "− Column");
    expect(onChange).not.toHaveBeenCalled();
    host.remove();
  });

  it("transposes via the control button", () => {
    const host = document.createElement("div");
    let model = sheet();
    mountDatasheet(host, model, (m) => (model = m));
    clickButton(host, "⇄ Transpose");
    expect(model.cells[0]).toEqual(["", "S1", "S2"]);
    expect(model.cells[1]).toEqual(["A", "1", "3"]);
  });

  it("expands the grid on multi-cell TSV paste", () => {
    const host = document.createElement("div");
    let model = sheet();
    mountDatasheet(host, model, (m) => (model = m));
    const target = cell(host, 2, 2);
    const e = new Event("paste") as ClipboardEvent;
    Object.defineProperty(e, "clipboardData", {
      value: { getData: () => "9\t8\n7\t6" },
    });
    target.dispatchEvent(e);
    expect(model.cells[2]).toEqual(["S2", "3", "9", "8"]);
    expect(model.cells[3]).toEqual(["", "", "7", "6"]);
    // Header row was widened to match.
    expect(model.cells[0]).toHaveLength(4);
  });

  it("leaves single-cell pastes to the browser default", () => {
    const host = document.createElement("div");
    const onChange = vi.fn();
    mountDatasheet(host, sheet(), onChange);
    const e = new Event("paste") as ClipboardEvent;
    Object.defineProperty(e, "clipboardData", { value: { getData: () => "just text" } });
    cell(host, 1, 1).dispatchEvent(e);
    expect(onChange).not.toHaveBeenCalled();
  });

  it("setSheet re-renders with the new model", () => {
    const host = document.createElement("div");
    const api = mountDatasheet(host, sheet(), () => {});
    api.setSheet({
      cells: [
        ["", "X"],
        ["S", "5"],
      ],
    });
    expect(inputs(host)).toHaveLength(4);
    expect(cell(host, 1, 1).value).toBe("5");
  });
});

describe("localizePane", () => {
  const build = () => {
    document.body.innerHTML = `
      <h2>2 · Data</h2>
      <button>Insert into slide</button>
      <label>Gridlines<input type="checkbox"></label>
      <button>Not translated</button>`;
  };

  it("translates matching visible strings for German", () => {
    build();
    localizePane("de-DE");
    expect(document.querySelector("h2")!.textContent).toBe("2 · Daten");
    expect(document.querySelector("button")!.textContent).toBe("In Folie einfügen");
  });

  it("keeps child inputs and unknown strings intact", () => {
    build();
    localizePane("de");
    const label = document.querySelector("label")!;
    expect(label.querySelector("input")).not.toBeNull();
    expect(label.textContent).toContain("Gitterlinien");
    expect(document.querySelectorAll("button")[1].textContent).toBe("Not translated");
  });

  it("is a no-op for unsupported or missing languages", () => {
    build();
    localizePane("fr-FR");
    localizePane(undefined);
    expect(document.querySelector("h2")!.textContent).toBe("2 · Data");
  });

  it("translates the grouped picker, Format groups, search and datasheet help", () => {
    document.body.innerHTML = `
      <div class="group-label">Columns &amp; bars</div>
      <span class="fgroup-name">Axes &amp; scale</span>
      <p class="no-type-result">No chart type matches that search.</p>
      <details><summary>Paste straight from Excel — special data rows</summary></details>
      <input placeholder="Search chart types…" />`;
    localizePane("de-DE");
    expect(document.querySelector(".group-label")!.textContent).toBe("Säulen & Balken");
    expect(document.querySelector(".fgroup-name")!.textContent).toBe("Achsen & Skala");
    expect(document.querySelector(".no-type-result")!.textContent).toBe("Kein Diagrammtyp passt zu dieser Suche.");
    expect(document.querySelector("summary")!.textContent).toBe("Direkt aus Excel einfügen — besondere Datenzeilen");
    expect(document.querySelector("input")!.placeholder).toBe("Diagrammtypen suchen…");
  });

  it("localizeTree re-applies the active language to a freshly rendered subtree", () => {
    localizePane("de"); // sets the active language
    const root = document.createElement("div");
    root.innerHTML = `<div class="group-label">Distribution</div><span class="fgroup-name">Analysis</span>`;
    localizeTree(root);
    expect(root.querySelector(".group-label")!.textContent).toBe("Verteilung");
    expect(root.querySelector(".fgroup-name")!.textContent).toBe("Analyse");
  });
});
