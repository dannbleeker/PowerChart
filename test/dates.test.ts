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
 * A slash date whose first two numbers could both be a month is TWO dates, and
 * nothing in a cell says which. `Date.parse` picked the American one silently
 * while the dotted form — the one `docs/MANUAL.md` documents — picks the
 * European one, so the same digits in one datasheet meant dates two months
 * apart.
 */
describe("parseDateToken refuses a slash date that means two things", () => {
  const iso = (d: number | null) => (d == null ? null : new Date(d * 86400000).toISOString().slice(0, 10));

  it("refuses the ambiguous ones rather than guessing a month", () => {
    // Each of these is a different date read American-first or European-first.
    for (const t of ["03/01/2026", "01/03/2026", "12/11/2026", "05/06/2026", "06/07/2026"]) {
      expect(parseDateToken(t), `${t} was guessed at instead of refused`).toBeNull();
    }
    // The mirror image, so the size of the mistake is on record: the SAME digits
    // with a dot are the documented European form and still parse that way.
    expect(iso(parseDateToken("03.01.2026"))).toBe("2026-01-03");
    expect(iso(parseDateToken("01.03.2026"))).toBe("2026-03-01");
  });

  it("leaves every unambiguous date alone", () => {
    // Only one reading exists (15 is not a month), so nothing is being guessed.
    expect(iso(parseDateToken("01/15/2026"))).toBe("2026-01-15");
    // Both readings agree.
    expect(iso(parseDateToken("03/03/2026"))).toBe("2026-03-03");
    expect(iso(parseDateToken("1/1/2026"))).toBe("2026-01-01");
    expect(iso(parseDateToken("12/12/2026"))).toBe("2026-12-12");
    // A four-digit year first is Y/M/D and never ambiguous — the rule must not
    // reach it.
    expect(iso(parseDateToken("2026/01/15"))).toBe("2026-01-15");
    // …nor any of the forms that do not use slashes at all.
    expect(iso(parseDateToken("2026-01-15"))).toBe("2026-01-15");
    expect(iso(parseDateToken("15.01.2026"))).toBe("2026-01-15");
    expect(iso(parseDateToken("15 Jan 2026"))).toBe("2026-01-15");
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

/**
 * A date cell must mean the same day on every machine that opens the deck.
 *
 * `Date.parse` reads a bare ISO date as UTC and an offset-less ISO date-TIME as
 * LOCAL — the one inconsistency this parser handed straight through. So
 * `2026-01-15T20:00` became 2026-01-16T01:00Z for a reader in New York and
 * `2026-01-15T02:00` became 2026-01-14T17:00Z for one in Tokyo, and the
 * day-existence check then compared the 15 the token NAMES against the 16 or
 * the 14 the instant landed on and returned NULL. The cell did not shift, it
 * stopped being a date — so a Gantt row vanished, or a line chart's date
 * spacing fell back to even steps, for some readers and not others, depending
 * on the time of day in the cell.
 *
 * An explicit offset ended the same way by a different route: a perfectly legal
 * `2026-01-15T20:00-05:00` is 16 January in UTC, so the day check refused it in
 * every timezone on earth.
 *
 * The answer to both is the calendar date the token names. `process.env.TZ` is
 * honoured by `Date` at runtime in Node, so these run the same token under
 * several zones in one process — which is the only way to catch this on a CI
 * runner that is always UTC.
 */
describe("a date token means the same day in every timezone", () => {
  const ZONES = ["UTC", "America/New_York", "Asia/Tokyo", "Pacific/Kiritimati", "Pacific/Pago_Pago"];
  const inZones = (token: string): (number | null)[] => {
    const was = process.env.TZ;
    try {
      return ZONES.map((tz) => {
        process.env.TZ = tz;
        return parseDateToken(token);
      });
    } finally {
      process.env.TZ = was;
    }
  };

  it("reads an ISO date-time with no zone as the day it names", () => {
    const utcDay = parseDateToken("2026-01-15")!;
    for (const token of ["2026-01-15T20:00", "2026-01-15T02:00", "2026-01-15 20:00", "2026-01-15T23:59:59"]) {
      const answers = inZones(token);
      expect(new Set(answers).size, `${token} answered ${answers.join("/")} across ${ZONES.join(", ")}`).toBe(1);
      expect(answers[0], token).toBe(utcDay);
    }
  });

  it("accepts an ISO date-time carrying an explicit offset", () => {
    // 20:00−05:00 is the next day in UTC; the cell still names the 15th.
    expect(inZones("2026-01-15T20:00-05:00")).toEqual(Array(ZONES.length).fill(parseDateToken("2026-01-15")));
    expect(inZones("2026-01-15T02:00+09:00")).toEqual(Array(ZONES.length).fill(parseDateToken("2026-01-15")));
    expect(inZones("2026-01-15T20:00:00Z")).toEqual(Array(ZONES.length).fill(parseDateToken("2026-01-15")));
  });

  it("still refuses a day that does not exist, in every zone", () => {
    // The negative control: the guard the timezone drift was tripping is still
    // there for the mistake it was written for.
    expect(inZones("2026-02-29")).toEqual(Array(ZONES.length).fill(null));
    expect(inZones("2026-04-31T09:00")).toEqual(Array(ZONES.length).fill(null));
  });

  it("refuses a month outside the year, rather than rolling into the next one", () => {
    // `Date.UTC` normalises, so a 13th month is 1 January of the year after —
    // a Gantt row a year out, silently. The dotted form has always been able to
    // do this; reading the ISO form ourselves would newly let it.
    expect(parseDateToken("2026-13-01")).toBeNull();
    expect(parseDateToken("15.13.2026")).toBeNull();
    expect(parseDateToken("2026-00-05")).toBeNull();
    expect(parseDateToken("00.01.2026")).toBeNull();
    // …and the last real month of the year still parses, both spellings.
    expect(parseDateToken("2026-12-31")).toBe(parseDateToken("31.12.2026"));
  });
});
