import { describe, expect, it } from "vitest";
import { sheetToData, dataToSheet } from "../src/taskpane/datasheet";
import { parseDateToken } from "../src/core/format";

/**
 * The datasheet formula evaluator treated a blank cell as 0 for every aggregate.
 * That's Excel's convention for SUM, but MIN/MAX/AVG ignore empty cells — so a
 * gap in a range used to pin MIN at ≤0, MAX at ≥0, and skew AVG toward zero.
 *
 * Drive the evaluator through sheetToData: put a formula in a series cell and read
 * the value it resolves to (row 1 = categories, column A = series name).
 */
const evalFormula = (grid: string[][], formula: string): number | null => {
  // One category so the series row has exactly one value cell (B2) carrying the formula.
  const cells = [["", "R"], ["S", formula], ...grid.map((r) => r)];
  return sheetToData({ cells }).series[0].values[0];
};

// A data block placed at rows 3+ (A3.. = row index 2) that the formula references.
// We reference an explicit A1 range in the grid the formula sees.
const withRow = (values: string[]) => {
  // Row 3 (index 2): A3 label, B3..= values. Range B3:<end>3.
  return [["row", ...values]];
};

describe("MIN/MAX/AVG ignore blank cells; SUM still counts them as 0", () => {
  it("AVG skips a blank instead of averaging in a phantom 0", () => {
    // B3=10, C3="", D3=20 → range B3:D3
    expect(evalFormula(withRow(["10", "", "20"]), "=AVG(B3:D3)")).toBe(15);
  });

  it("MIN ignores a blank (does not return 0)", () => {
    expect(evalFormula(withRow(["10", "", "20"]), "=MIN(B3:D3)")).toBe(10);
  });

  it("MAX of an all-negative range with a gap is not 0", () => {
    expect(evalFormula(withRow(["-5", "", "-8"]), "=MAX(B3:D3)")).toBe(-5);
  });

  it("SUM still treats a blank as 0 (Excel convention, unchanged)", () => {
    expect(evalFormula(withRow(["10", "", "20"]), "=SUM(B3:D3)")).toBe(30);
  });

  it("aggregates over an all-blank range resolve to no-data, not 0", () => {
    // MAX over only blanks → NaN → the datasheet turns it into a null (gap).
    expect(evalFormula(withRow(["", "", ""]), "=MAX(B3:D3)")).toBeNull();
  });
});

/**
 * Bugs found by an adversarial hunt over the datasheet/date layer.
 */
describe("datasheet + date parsing hardening", () => {
  it("does not read a percentage cell as a calendar date", () => {
    // `Date.parse("50% UTC")` returns a finite garbage instant, so an ordinary
    // "50%" cell became epoch day -7305 AND flipped the chart into date mode.
    expect(parseDateToken("50%")).toBeNull();
    expect(parseDateToken("5%")).toBeNull();
    const d = sheetToData({
      cells: [
        ["", "Q1", "Q2"],
        ["A", "50%", "60%"],
      ],
    });
    expect(d.series[0].values).toEqual([50, 60]); // read as shares, not epoch days
    expect((d as { dates?: boolean }).dates).toBeFalsy();
  });

  it("reads a pasted Excel percent column as numbers (Excel copies the DISPLAYED text)", () => {
    // A share table is the canonical source for a 100%/stacked chart, and Excel
    // puts "35%" on the clipboard — not 0.35. Dropping it to a blank gap would
    // render an empty chart, so the % is stripped like the thousands separator.
    const d = sheetToData({
      cells: [
        ["", "2024", "2025"],
        ["Online", "35%", "42%"],
        ["Wholesale", "65%", "58%"],
      ],
    });
    expect(d.series.map((s) => s.values)).toEqual([
      [35, 42],
      [65, 58],
    ]);
    expect((d as { dates?: boolean }).dates).toBeFalsy();
    // The user's typed value survives the round trip — it used to be overwritten
    // in the datasheet with the garbage epoch day.
    expect(dataToSheet(d).cells[1]).toEqual(["Online", "35", "42"]);
    // Signed / spaced / fractional percents parse too.
    const misc = sheetToData({
      cells: [
        ["", "a", "b", "c"],
        ["r", "50 %", "1.5%", "-20%"],
      ],
    });
    expect(misc.series[0].values).toEqual([50, 1.5, -20]);
  });

  it("parses a full ISO-8601 date-time, not just a bare date", () => {
    // Appending " UTC" to a date-time made Date.parse NaN, so every task in a
    // pasted ISO export was silently dropped.
    const bare = parseDateToken("2026-01-05");
    expect(bare).not.toBeNull();
    expect(parseDateToken("2026-01-05T00:00:00.000Z")).toBe(bare);
    expect(parseDateToken("2026-01-05T09:30:00Z")).toBe(bare);
  });

  it("floors a date-time to its calendar day instead of rounding up", () => {
    const day = parseDateToken("2026-01-15")!;
    expect(parseDateToken("2026-01-15T18:00:00Z")).toBe(day); // was day + 1
    expect(parseDateToken("2026-01-15T00:00:00Z")).toBe(day);
  });

  it("ignores a blank cell in comma-separated aggregate args, as the range form does", () => {
    // =MIN(B2,C2,D2) counted the blank C2 as a real 0 while =MIN(B2:D2) ignored it.
    const cells = [
      ["", "c1", "c2", "c3"],
      ["r", "10", "", "20"],
      ["out", "=MIN(B2,C2,D2)", "=MIN(B2:D2)", "=AVG(B2,C2,D2)"],
    ];
    const out = sheetToData({ cells }).series[1].values;
    expect(out[0]).toBe(10); // comma form — was 0
    expect(out[1]).toBe(10); // range form (already correct)
    expect(out[2]).toBe(15); // AVG ignores the blank — was 10
  });
});
