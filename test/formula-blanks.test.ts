import { describe, expect, it } from "vitest";
import { sheetToData } from "../src/taskpane/datasheet";

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
