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

  it("refuses a comma-decimal number instead of mangling it by 1000x", () => {
    // Excel in de-DE/da-DK copies a number to the clipboard as "1.234,5".
    // Stripping EVERY comma read that as 1.2345 — a value off by a factor of
    // 1000 with nothing on the slide to reveal it — and "1,5" as 15. A comma we
    // cannot read as US grouping now yields a visible gap instead.
    const d = sheetToData({
      cells: [
        ["", "a", "b", "c", "d"],
        ["r", "1.234,5", "1,5", "1,234", "1,234,567.5"],
      ],
    });
    expect(d.series[0].values).toEqual([null, null, 1234, 1234567.5]);
    expect((d as { dates?: boolean }).dates).toBeFalsy(); // and not a date either
    // The formula layer reads its operands through the same parser, so a cell
    // reference cannot inherit the mangled value behind the datasheet's back.
    const f = sheetToData({
      cells: [
        ["", "a"],
        ["r", "1.234,5"],
        ["out", "=B2*2"],
      ],
    });
    expect(f.series[1].values[0]).toBeNull();
  });

  it("round-trips every calendar Gantt row as an ISO date, not a raw epoch day", () => {
    // Only Start/End/Milestone used to be re-serialised as dates, so re-opening
    // your own calendar Gantt showed "20494" where you had typed "2026-02-10"
    // and the row could then only be edited in epoch days. layoutGantt puts
    // today / holidays / baseline / bracket rows through the same time scale.
    const cells = [
      ["", "Task A", "Task B"],
      ["Start", "2026-01-05", "2026-02-02"],
      ["End", "2026-01-30", "2026-03-06"],
      ["Today", "2026-02-10", ""],
      ["Holiday", "2026-01-01", "2026-04-03"],
      ["Baseline start", "2026-01-01", "2026-01-22"],
      ["Bracket Phase 1", "2026-01-05", "2026-02-20"],
      ["% complete", "40", "10"], // not a date — stays a number
    ];
    expect(dataToSheet(sheetToData({ cells })).cells).toEqual(cells);
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

/**
 * `parseRow` and `cellNumeric` each coerced a cell to a number their own way and
 * had drifted apart: only parseRow stripped a trailing "%" and only parseRow read
 * a date. So the same cell was a value to the chart and an error to a formula
 * over it — `=SUM` across a pasted share table came back blank, and so did a
 * Gantt duration `=B3-B2` over two ISO dates.
 */
describe("a formula reads the same cell the chart does", () => {
  const valueOf = (cells: string[][], series = 0) => sheetToData({ cells }).series[series].values;

  it("sums percent-formatted cells", () => {
    expect(
      valueOf(
        [
          ["", "A"],
          ["Share", "35%"],
          ["Rest", "65%"],
          ["Total", "=SUM(B2:B3)"],
        ],
        2,
      ),
    ).toEqual([100]);
  });

  it("subtracts two ISO date cells into a duration", () => {
    expect(
      valueOf(
        [
          ["", "Task"],
          ["Start", "2026-01-05"],
          ["End", "2026-01-15"],
          ["Days", "=B3-B2"],
        ],
        2,
      ),
    ).toEqual([10]);
  });

  it("still refuses a genuinely non-numeric cell", () => {
    expect(
      valueOf(
        [
          ["", "A"],
          ["Raw", "n/a"],
          ["Calc", "=B2+1"],
        ],
        1,
      ),
    ).toEqual([null]);
  });
});

/**
 * The cycle guard is added before a descent and removed after, so a formula
 * naming the same cell twice re-walked its whole subtree twice — a chain of such
 * cells cost 2^depth evaluations, on the UI thread, on every keystroke. A 24-row
 * chain took minutes; memoised it is instant.
 */
describe("formula evaluation does not blow up exponentially", () => {
  it("resolves a deep doubly-referencing chain quickly", () => {
    const depth = 24;
    const cells: string[][] = [["", "A"]];
    for (let r = 1; r <= depth; r++) cells.push([`R${r}`, `=B${r + 2}+B${r + 2}`]);
    cells.push([`R${depth + 1}`, "1"]);
    const started = performance.now();
    const values = sheetToData({ cells }).series.map((s) => s.values[0]);
    expect(performance.now() - started).toBeLessThan(2000);
    // Each level doubles the one below it: the top is 2^depth.
    expect(values[0]).toBe(2 ** depth);
    expect(values[depth]).toBe(1);
  });

  it("still returns null for a genuine cycle", () => {
    expect(
      sheetToData({
        cells: [
          ["", "A"],
          ["X", "=B3"],
          ["Y", "=B2"],
        ],
      }).series[0].values,
    ).toEqual([null]);
  });
});

/**
 * `new Date(v * 86400000).toISOString()` THROWS past the Date range, and a Gantt
 * date row can hold anything a cell can — an epoch-second timestamp pasted from
 * an export, or a typo. The exception escaped through applyConfig into the
 * template/selection listeners: the chart never loaded and nothing was reported.
 */
describe("dataToSheet survives an unrepresentable date value", () => {
  const gantt = (v: number) => ({
    categories: ["Task"],
    series: [{ name: "Start", values: [v] }],
    dates: true,
  });

  it.each([1e12, 1.7e9, 20260105, -1e12, Number.MAX_SAFE_INTEGER])("writes %s back as a number", (v) => {
    const sheet = dataToSheet(gantt(v));
    expect(sheet.cells[1][1]).toBe(String(v));
  });

  it("still writes a real epoch day as an ISO date", () => {
    expect(dataToSheet(gantt(20494)).cells[1][1]).toBe("2026-02-10");
  });
});
