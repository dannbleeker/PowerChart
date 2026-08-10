// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { mountDatasheet, sheetToData, type SheetModel } from "../src/taskpane/datasheet";
import { EN, localizePane, localizeTree, registerLanguage, type StringKey } from "../src/taskpane/i18n";

// A synthetic, fully-populated language: every catalogue key mapped to a marker.
// The app ships English-only, so this is how the DOM-sweep mechanism (which only
// runs when a language is registered) stays exercised. `«key»` makes a translated
// node obvious and proves the sweep matched the right source string.
const MARKER = Object.fromEntries((Object.keys(EN) as StringKey[]).map((k) => [k, `«${k}»`])) as Record<
  StringKey,
  string
>;
registerLanguage("xx", MARKER);

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

  it("does not turn a pasted gap into a stack separator", () => {
    // A row the paste APPENDS and leaves fully blank is `sheetToData`'s
    // clustered-stacked separator, so pasting a column that runs past the end of
    // the sheet with an ordinary Excel spacer row in it split one stacked column
    // into two side-by-side sub-stacks — no error, no cue. "+ Row" is guarded
    // against creating exactly this row and says so; paste grows rows too and
    // never was.
    const host = document.createElement("div");
    let model = sheet();
    mountDatasheet(host, model, (m) => (model = m));
    const e = new Event("paste") as ClipboardEvent;
    Object.defineProperty(e, "clipboardData", { value: { getData: () => ["10", "20", "", "40"].join("\n") } });
    cell(host, 1, 1).dispatchEvent(e);
    const appended = model.cells.slice(3);
    expect(appended.length, "the paste did not grow the sheet, so this proves nothing").toBeGreaterThan(0);
    for (const row of appended) {
      expect(
        row.every((c) => c.trim() === ""),
        `a fully blank row survived: ${JSON.stringify(row)}`,
      ).toBe(false);
    }
    // And the gap itself is still there — the row is what keeps every later
    // value on the category it was pasted against.
    expect(model.cells.length).toBe(5);
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

  /**
   * Excel appends a trailing newline to a copy. Filtering EVERY empty line
   * dropped interior ones too — and in a single-column copy a blank cell IS an
   * empty line, so a gap pulled every later value up a row and onto the wrong
   * category. No error, no visual cue, just wrong numbers.
   */
  it("keeps the gap when a single-column paste has a blank cell", () => {
    const host = document.createElement("div");
    let model = sheet();
    mountDatasheet(host, model, (m) => (model = m));
    const e = new Event("paste") as ClipboardEvent;
    // 10, blank, 30 — with Excel's trailing newline.
    Object.defineProperty(e, "clipboardData", { value: { getData: () => "10\n\n30\n" } });
    cell(host, 1, 1).dispatchEvent(e);
    expect(model.cells[1][1]).toBe("10");
    expect(model.cells[2][1]).toBe("");
    expect(model.cells[3][1]).toBe("30");
  });

  /**
   * seriesColors / seriesMeta in the pane are POSITIONAL. The grid buttons
   * spliced rows in and out without reporting it, so every series below the edit
   * inherited its neighbour's colour and combo type.
   */
  describe("structural edits are reported so positional state can follow", () => {
    const threeSeries = (): SheetModel => ({
      cells: [
        ["", "A", "B"],
        ["100%=", "10", "10"],
        ["S1", "1", "2"],
        ["S2", "3", "4"],
        ["S3", "5", "6"],
      ],
    });

    const events = (focus: { row: number; col: number }, label: string) => {
      // Attached: a detached element cannot take focus in jsdom, and the grid's
      // row/column operations act at the focused cell.
      const host = document.body.appendChild(document.createElement("div"));
      const seen: unknown[] = [];
      mountDatasheet(
        host,
        threeSeries(),
        () => {},
        (c) => seen.push(c),
      );
      cell(host, focus.row, focus.col).focus();
      clickButton(host, label);
      host.remove();
      return seen;
    };

    it("reports a row insert as a SERIES index, skipping the 100%= row", () => {
      // Cursor on S1 (sheet row 2) → "+ Row" inserts after it, at series 1.
      expect(events({ row: 2, col: 1 }, "+ Row")).toEqual([{ kind: "series-insert", index: 1 }]);
    });

    it("reports a row delete as a series index", () => {
      // Cursor on S2 (sheet row 3) → series 1.
      expect(events({ row: 3, col: 1 }, "− Row")).toEqual([{ kind: "series-remove", index: 1 }]);
    });

    it("says nothing when the deleted row is not a series", () => {
      expect(events({ row: 1, col: 1 }, "− Row")).toEqual([]);
    });

    it("reports column edits as category indices", () => {
      expect(events({ row: 1, col: 1 }, "+ Column")).toEqual([{ kind: "category-insert", index: 1 }]);
      expect(events({ row: 1, col: 2 }, "− Column")).toEqual([{ kind: "category-remove", index: 1 }]);
    });

    it("reports a transpose as a full reset", () => {
      expect(events({ row: 1, col: 1 }, "⇄ Transpose")).toEqual([{ kind: "reset" }]);
    });
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

  it("translates matching visible strings for a registered language", () => {
    build();
    localizePane("xx");
    expect(document.querySelector("h2")!.textContent).toBe("«2 · Data»");
    expect(document.querySelector("button")!.textContent).toBe("«Insert into slide»");
  });

  it("keeps child inputs and unknown strings intact", () => {
    build();
    localizePane("xx");
    const label = document.querySelector("label")!;
    expect(label.querySelector("input")).not.toBeNull();
    expect(label.textContent).toContain("«Gridlines»");
    expect(document.querySelectorAll("button")[1].textContent).toBe("Not translated");
  });

  it("is a no-op for unsupported or missing languages (English default)", () => {
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
    localizePane("xx");
    expect(document.querySelector(".group-label")!.textContent).toBe("«Columns & bars»");
    expect(document.querySelector(".fgroup-name")!.textContent).toBe("«Axes & scale»");
    expect(document.querySelector(".no-type-result")!.textContent).toBe("«No chart type matches that search.»");
    expect(document.querySelector("summary")!.textContent).toBe("«Paste straight from Excel — special data rows»");
    expect(document.querySelector("input")!.placeholder).toBe("«Search chart types…»");
  });

  it("localizeTree re-applies the active language to a freshly rendered subtree", () => {
    localizePane("xx"); // sets the active language
    const root = document.createElement("div");
    root.innerHTML = `<div class="group-label">Distribution</div><span class="fgroup-name">Analysis</span>`;
    localizeTree(root);
    expect(root.querySelector(".group-label")!.textContent).toBe("«Distribution»");
    expect(root.querySelector(".fgroup-name")!.textContent).toBe("«Analysis»");
  });
});
