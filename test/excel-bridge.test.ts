// @vitest-environment jsdom
import { beforeAll, describe, expect, it, vi } from "vitest";
import type { ChartConfig } from "../src/core/types";

/**
 * Drives the Excel companion pane end-to-end with stubbed Excel/Office
 * globals: selected range → generated PowerChart JSON in the output box.
 */

let rangeValues: unknown[][] = [];

beforeAll(async () => {
  document.body.innerHTML = `
    <select id="kind"><option value="stacked">stacked</option><option value="waterfall">waterfall</option></select>
    <input id="title" value="">
    <input type="checkbox" id="transpose">
    <button id="generate"></button>
    <button id="copy"></button>
    <textarea id="output"></textarea>
    <p id="note"></p>`;

  vi.stubGlobal("Office", { onReady: (cb: () => void) => cb(), context: { host: "Excel" } });
  vi.stubGlobal("Excel", {
    run: async (cb: (ctx: unknown) => Promise<void>) =>
      cb({
        workbook: {
          getSelectedRange: () => ({
            load() {},
            get values() {
              return rangeValues;
            },
            address: "Sheet1!A1:C3",
          }),
        },
        sync: async () => {},
      }),
  });
  await import("../src/excel/excel");
});

const generate = () => {
  document.getElementById("generate")!.click();
  // The click handler is async; flush the microtask queue.
  return new Promise((r) => setTimeout(r, 0));
};

const output = (): ChartConfig => JSON.parse((document.getElementById("output") as HTMLTextAreaElement).value);

describe("Excel data bridge", () => {
  it("enables the Generate button inside the Excel host", () => {
    expect((document.getElementById("generate") as HTMLButtonElement).disabled).toBe(false);
  });

  it("converts the selected range into a chart config", async () => {
    rangeValues = [
      ["", "Q1", "Q2"],
      ["North", 10, 20],
      ["South", 5, null],
    ];
    (document.getElementById("title") as HTMLInputElement).value = "Regional sales";
    await generate();
    const cfg = output();
    expect(cfg.kind).toBe("stacked");
    expect(cfg.title).toBe("Regional sales");
    expect(cfg.data.categories).toEqual(["Q1", "Q2"]);
    expect(cfg.data.series).toEqual([
      { name: "North", values: [10, 20] },
      { name: "South", values: [5, null] },
    ]);
    expect(document.getElementById("note")!.textContent).toContain("Sheet1!A1:C3");
  });

  it("transposes when the series run across the top row", async () => {
    (document.getElementById("kind") as HTMLSelectElement).value = "stacked";
    (document.getElementById("title") as HTMLInputElement).value = "";
    (document.getElementById("transpose") as HTMLInputElement).checked = true;
    rangeValues = [
      ["", "North", "South"], // series across the top
      ["Q1", 10, 5],
      ["Q2", 20, 8],
    ];
    await generate();
    const cfg = output();
    expect(cfg.data.categories).toEqual(["Q1", "Q2"]);
    expect(cfg.data.series).toEqual([
      { name: "North", values: [10, 20] },
      { name: "South", values: [5, 8] },
    ]);
    (document.getElementById("transpose") as HTMLInputElement).checked = false;
  });

  it("collects waterfall total markers into totalIndices", async () => {
    rangeValues = [
      ["", "Start", "Delta", "End"],
      ["Bridge", 50, 20, "e"],
    ];
    (document.getElementById("kind") as HTMLSelectElement).value = "waterfall";
    await generate();
    const cfg = output();
    expect(cfg.kind).toBe("waterfall");
    expect(cfg.waterfall).toEqual({ totalIndices: [2] });
  });

  it("reports Excel API failures in the note", async () => {
    vi.stubGlobal("Excel", {
      run: async () => {
        throw new Error("range is protected");
      },
    });
    await generate();
    expect(document.getElementById("note")!.textContent).toContain("range is protected");
  });
});
