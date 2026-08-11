// @vitest-environment jsdom
import { describe, it, expect } from "vitest";
import { mountDatasheet, sheetToData, dataToSheet } from "../src/taskpane/datasheet";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";

describe("ragged sheet from an ordinary imported config", () => {
  it("series with fewer values than categories", () => {
    const cfg = {
      kind: "clustered",
      width: 400,
      height: 300,
      data: { categories: ["Q1", "Q2", "Q3"], series: [{ name: "Sales", values: [10] }] },
    } as ChartConfig;
    // The engine accepts it and pads.
    console.log("engine values:", JSON.stringify(buildChart(cfg).nodes.length));
    const sheet = dataToSheet(cfg.data);
    console.log("sheet cells:", JSON.stringify(sheet.cells));
    const host = document.createElement("div");
    mountDatasheet(host, sheet, () => {});
    const rows = [...host.querySelectorAll("tr")].map((tr) => tr.querySelectorAll("input").length);
    console.log("inputs per rendered row:", JSON.stringify(rows));
    expect(rows[1]).toBe(rows[0]);
  });

  it("series with no values at all — what the engine happily renders", () => {
    const data = { categories: ["Q1", "Q2"], series: [{ name: "Sales" }] } as unknown as ChartConfig["data"];
    console.log("engine ok:", buildChart({ kind: "clustered", width: 400, height: 300, data } as ChartConfig).nodes.length > 0);
    let threw = "";
    try {
      dataToSheet(data);
    } catch (e) {
      threw = String(e);
    }
    console.log("dataToSheet:", threw || "ok");
    expect(threw).toBe("");
  });

  it("sheetToData tolerates ragged rows?", () => {
    const sheet = { cells: [["", "a", "b"], ["S", "1"], []] };
    let threw = "";
    try {
      console.log(JSON.stringify(sheetToData(sheet)));
    } catch (e) {
      threw = String(e);
    }
    console.log("sheetToData ragged:", threw || "ok");
  });
});
