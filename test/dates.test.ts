import { describe, expect, it } from "vitest";
import { parseDateToken } from "../src/core/format";

/** Date-token parsing and gantt holiday brackets. */

describe("parseDateToken numeric ranges", () => {
  it("rejects hyphenated numeric ranges as category labels", () => {
    expect(parseDateToken("3-5")).toBeNull();
    expect(parseDateToken("10-20")).toBeNull();
    expect(parseDateToken("18–24")).toBeNull(); // en dash
  });
  it("still parses real dates", () => {
    expect(parseDateToken("2026-01-15")).toBeTypeOf("number");
    expect(parseDateToken("2026-01")).toBeTypeOf("number");
  });
});

/**
 * `Date.parse` is far more lenient than a date cell has any right to be. Excel's
 * `#DIV/0!` came back as 2000-01-01 — so an error cell became epoch day 10957
 * AND flipped the whole ChartData into date mode — and so did a threshold
 * ("<0.1") and any "<word> <number>" category ("Store 5" → 2001-05-01).
 */
describe("parseDateToken rejects things that only LOOK parseable", () => {
  it.each(["#DIV/0!", "#N/A", "#REF!", "#VALUE!", "#NUM!", "#NAME?", "#SPILL!"])("rejects the Excel error %s", (t) => {
    expect(parseDateToken(t)).toBeNull();
  });

  it.each(["<0.1", ">1000", "~5", "n/a*", "50%"])("rejects the non-date token %s", (t) => {
    expect(parseDateToken(t)).toBeNull();
  });

  it.each(["Store 5", "Top 10", "Region 3", "Cat 1", "Week 12", "Rev 2024", "Q1 2026"])(
    "rejects the category label %s",
    (t) => {
      expect(parseDateToken(t)).toBeNull();
    },
  );

  it("still parses every real date shape", () => {
    for (const t of [
      "2026-01-15",
      "2026-01-15T10:30:00Z",
      "2026-01-15T10:30:00+02:00",
      "15.01.2026",
      "Jan 2026",
      "January 2026",
      "15 Jan 2026",
      "Mon, 05 Jan 2026",
      "2026-01",
    ]) {
      expect(parseDateToken(t), t).not.toBeNull();
    }
  });
});
