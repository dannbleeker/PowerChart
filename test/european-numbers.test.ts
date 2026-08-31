import { describe, expect, it } from "vitest";
import { looksEuropean, fromEuropeanNumber } from "../src/taskpane/datasheet";

/**
 * WHEN THE PANE MAY DECIDE THAT A DOT IS A THOUSAND, AND WHEN IT MAY NOT.
 *
 * Per cell it cannot be decided and never will be: "1.234" is a perfectly good
 * en-US decimal. Across a PASTE it usually can, because the evidence arrives in
 * the same clipboard block — a comma that is not a thousands group is a decimal
 * comma, and a decimal comma means the dots are groups.
 *
 * So the rule is asymmetric on purpose: European evidence only counts when
 * nothing contradicts it. A block that says both, or neither, keeps the reading
 * it has always had. Getting this wrong in the eager direction would corrupt US
 * data, which is worse than leaving European data as it was — so every
 * ambiguous case below is a case where the pane must NOT act.
 */
describe("deciding a paste is written in the European convention", () => {
  const block = (...cells: string[]) => [["Region", "N"], ...cells.map((c) => ["x", c])];

  it("needs evidence a comma is a decimal, not a thousands group", () => {
    // "987,5" cannot be a US grouping — a group is exactly three digits.
    expect(looksEuropean(block("1.234", "987,5"))).toBe(true);
    // "1.234,56" says it outright: dot groups AND a comma decimal.
    expect(looksEuropean(block("1.234,56"))).toBe(true);
    // "1,5" likewise — one digit after the comma.
    expect(looksEuropean(block("1,5"))).toBe(true);
  });

  it("refuses when the block contradicts itself", () => {
    // A US grouping is decisive the other way. Both present = a mixed paste, or
    // a category that happens to look like a number; either way the pane has no
    // business guessing.
    expect(looksEuropean(block("987,5", "1,234"))).toBe(false);
    // A dot decimal that CANNOT be a group ("0.5") is also decisive.
    expect(looksEuropean(block("1,5", "0.5"))).toBe(false);
  });

  it("refuses when there is nothing decisive at all", () => {
    // The case that matters most: bare dot-grouped numbers with no comma
    // anywhere. "1.234" and "2.500" are exactly as likely to be US decimals,
    // and acting here would corrupt them.
    expect(looksEuropean(block("1.234", "2.500"))).toBe(false);
    expect(looksEuropean(block("1234", "2500"))).toBe(false);
    expect(looksEuropean(block("Nord", "Syd"))).toBe(false);
    expect(looksEuropean([])).toBe(false);
  });

  it("is not fooled by a US paste", () => {
    expect(looksEuropean(block("1,234", "2,500", "987.5"))).toBe(false);
  });
});

describe("rewriting a European number into the form the sheet stores", () => {
  it("turns dot groups into a plain number", () => {
    expect(fromEuropeanNumber("1.234")).toBe("1234");
    expect(fromEuropeanNumber("1.234.567")).toBe("1234567");
    expect(fromEuropeanNumber("1.234,56")).toBe("1234.56");
    expect(fromEuropeanNumber("-1.234,5")).toBe("-1234.5");
  });

  it("turns a decimal comma into a decimal point", () => {
    expect(fromEuropeanNumber("987,5")).toBe("987.5");
    expect(fromEuropeanNumber("0,25")).toBe("0.25");
  });

  it("leaves anything that is not a number in that convention alone", () => {
    // Text, dates, formulas and plain machine numbers all pass through — the
    // rewrite runs over every cell of a European block, so what it declines to
    // touch matters as much as what it changes.
    for (const s of ["Nord", "", "=SUM(B2:B4)", "2026-01-05", "1234", "987.5", "35%", "1.2"]) {
      expect(fromEuropeanNumber(s), s).toBe(s);
    }
  });

  it("reads a comma as a decimal even at three digits — which is why the gate matters", () => {
    // "1,234" is a thousand two hundred in Copenhagen only if you read the comma
    // as a group, and a European reader does not: to them it is one point two
    // three four, a perfectly ordinary three-decimal number. The rewrite is
    // right to convert it.
    expect(fromEuropeanNumber("1,234")).toBe("1.234");
    // And that is exactly why `looksEuropean` treats a US grouping as decisive
    // evidence AGAINST: this one conversion, applied to a US block, is the same
    // 1000x error pointing the other way. The rewrite cannot tell; the gate can.
    expect(looksEuropean([["1,234"], ["2,500"]])).toBe(false);
  });
});
