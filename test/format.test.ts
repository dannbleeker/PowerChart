import { describe, expect, it } from "vitest";
import { cagr, formatNumber, formatPercent, niceTicks } from "../src/core/format";

describe("formatNumber", () => {
  it("picks decimals by magnitude", () => {
    expect(formatNumber(1234)).toBe("1,234");
    expect(formatNumber(5.25)).toBe("5.3");
    expect(formatNumber(0.123)).toBe("0.12");
  });
  it("forces sign and suffix", () => {
    expect(formatNumber(12, { forceSign: true })).toBe("+12");
    expect(formatNumber(-12, { forceSign: true })).toBe("-12");
    expect(formatNumber(12, { suffix: "%" })).toBe("12%");
  });
});

describe("formatPercent", () => {
  it("formats ratios", () => {
    expect(formatPercent(0.256)).toBe("26%");
    expect(formatPercent(0.081, 1, true)).toBe("+8.1%");
  });
});

describe("niceTicks", () => {
  it("covers the data range with round steps", () => {
    const t = niceTicks(0, 97);
    expect(t[0]).toBe(0);
    expect(t[t.length - 1]).toBeGreaterThanOrEqual(97);
    const step = t[1] - t[0];
    for (let i = 1; i < t.length; i++) expect(t[i] - t[i - 1]).toBeCloseTo(step);
  });
  it("spans zero for mixed-sign data", () => {
    const t = niceTicks(-30, 80);
    expect(t[0]).toBeLessThanOrEqual(-30);
    expect(t).toContain(0);
  });
  it("handles degenerate ranges", () => {
    expect(niceTicks(5, 5).length).toBeGreaterThan(1);
    expect(niceTicks(0, 0)).toEqual([0, 1]);
  });
});

describe("cagr", () => {
  it("computes compound growth", () => {
    expect(cagr(100, 121, 2)).toBeCloseTo(0.1);
  });
  it("rejects non-positive endpoints and zero periods", () => {
    expect(cagr(0, 10, 2)).toBeNull();
    expect(cagr(10, -5, 2)).toBeNull();
    expect(cagr(10, 20, 0)).toBeNull();
  });
});
