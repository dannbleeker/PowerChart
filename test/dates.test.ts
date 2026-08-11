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

/**
 * `Jan-24` is Excel's `mmm-yy` and the commonest monthly category label there
 * is. `Date.parse` read the 24 as a DAY in its own default year (2001), which
 * looks harmless inside one year — the month-to-month gaps come out right — and
 * is not harmless across a year boundary: January landed 333 days BEFORE the
 * previous December, so a line chart drew its newest two months at the far left
 * of the plot, in front of October.
 */
describe("parseDateToken reads mmm-yy as a month and a year", () => {
  const day = (iso: string) => Math.floor(Date.parse(iso + "T00:00:00Z") / 86400000);

  it("puts the month first and the two-digit year second", () => {
    expect(parseDateToken("Jan-24")).toBe(day("2024-01-01"));
    expect(parseDateToken("Dec-23")).toBe(day("2023-12-01"));
    expect(parseDateToken("Jan/24")).toBe(day("2024-01-01"));
    expect(parseDateToken("January-24")).toBe(day("2024-01-01"));
  });

  it("orders a series that crosses new year", () => {
    const months = ["Oct-23", "Nov-23", "Dec-23", "Jan-24", "Feb-24"].map((m) => parseDateToken(m)!);
    for (let i = 1; i < months.length; i++) {
      expect(months[i], `${i}`).toBeGreaterThan(months[i - 1]);
    }
  });

  it("pivots two-digit years at 30, as Excel does", () => {
    expect(parseDateToken("Jan-29")).toBe(day("2029-01-01"));
    expect(parseDateToken("Jan-30")).toBe(day("1930-01-01"));
    expect(parseDateToken("Jan-99")).toBe(day("1999-01-01"));
  });

  it("leaves the genuinely ambiguous and the non-months alone", () => {
    // A space is how a day-of-month is written; only the hyphen/slash forms are
    // claimed. And a word that is not a month still falls through as before.
    expect(parseDateToken("Jan 24")).toBe(parseDateToken("24 Jan"));
    expect(parseDateToken("Mon-24")).toBeNull();
    expect(parseDateToken("3-5")).toBeNull();
  });
});

/**
 * A cell is not reliably a string.
 *
 * This function's whole job is to decide whether a cell is a date, and its
 * input arrives from a pasted block, a JSON config and the skill's caller.
 * `null`/`undefined` threw `Cannot read properties of null (reading 'trim')`,
 * and a NUMBER threw `raw.trim is not a function` — the case that matters,
 * because a bare number is explicitly not a date (the guard below the trim says
 * so) and it could not reach that answer without crashing first.
 */
describe("parseDateToken is handed whatever the cell held", () => {
  it("says 'not a date' for an absent cell instead of throwing", () => {
    for (const v of [null, undefined, ""] as unknown as string[]) {
      expect(parseDateToken(v)).toBeNull();
    }
  });

  it("says 'not a date' for a bare number, the answer it already gives the string", () => {
    expect(parseDateToken(2024 as unknown as string)).toBeNull();
    expect(parseDateToken("2024")).toBeNull();
    expect(parseDateToken(0 as unknown as string)).toBeNull();
  });
});
