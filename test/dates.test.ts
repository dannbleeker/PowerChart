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
