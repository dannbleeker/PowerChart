import { describe, expect, it } from "vitest";
import { cagr, formatNumber, formatPercent, niceTicks, polyTrend, trendStats } from "../src/core/format";

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
  it("normalises a negative that rounds to zero, in every locale", () => {
    // en/de/fr/da use ASCII "-", but sv/nb/fi/lt/et emit U+2212 MINUS SIGN.
    for (const locale of ["en-US", "de-DE", "da-DK", "sv-SE", "nb-NO", "fi-FI", "lt-LT"]) {
      expect(formatNumber(-0.004, { decimals: 0, locale }), locale).toBe("0");
    }
    // And the ones a pattern match on the formatted string could never cover:
    // RTL locales prefix an invisible directional mark (U+200E/U+061C) ahead of
    // the sign, and these render digits outside ASCII.
    const zeroOf: Record<string, string> = {
      "ar-AE": "0",
      "ur-PK": "0",
      "he-IL": "0", // ASCII digits behind a mark
      "ar-EG": "٠",
      "fa-IR": "۰",
      "bn-BD": "০",
      "my-MM": "၀",
    };
    for (const [locale, zero] of Object.entries(zeroOf)) {
      expect(formatNumber(-0.4, { decimals: 0, locale }), locale).toBe(zero);
      expect(formatNumber(-0.004, { decimals: 2, locale }), locale).not.toMatch(/[-−]/);
    }
  });

  it("keeps the sign on a genuine negative in those same locales", () => {
    // The normalisation must key off the value, not strip a leading character.
    for (const locale of ["ar-EG", "fa-IR", "bn-BD", "my-MM", "he-IL", "sv-SE"]) {
      expect(formatNumber(-5.2, { decimals: 1, locale }), locale).toMatch(/[-−]/);
    }
  });
});

describe("formatPercent", () => {
  it("formats ratios", () => {
    expect(formatPercent(0.256)).toBe("26%");
    expect(formatPercent(0.081, 1, true)).toBe("+8.1%");
  });
  it("normalises a negative that rounds to zero", () => {
    expect(formatPercent(-0.001)).toBe("0%");
    expect(formatPercent(-0.00004, 2)).toBe("0.00%");
  });
});

describe("trendStats", () => {
  // For df = 1 the two-tailed Student-t p-value has the closed form
  // p = 1 - (2/pi)·atan(|t|), which pins the regularized incomplete beta.
  it("matches the exact p-value for df = 1", () => {
    const r = trendStats([
      { x: 1, y: 1 },
      { x: 2, y: 2.5 },
      { x: 3, y: 2.9 },
    ])!;
    const t = Math.sqrt((r.r2 * 1) / (1 - r.r2));
    expect(r.p).toBeCloseTo(1 - (2 / Math.PI) * Math.atan(t), 6);
    expect(r.p).toBeCloseTo(0.20536, 4); // NOT ~0.014 — a weak 3-point fit is not significant
  });
  it("still finds a strong fit significant", () => {
    const pts = Array.from({ length: 12 }, (_, i) => ({ x: i, y: 2 * i + (i % 2 ? 0.2 : -0.2) }));
    const r = trendStats(pts)!;
    expect(r.r2).toBeGreaterThan(0.99);
    expect(r.p!).toBeLessThan(0.001);
  });
  it("returns p in [0,1] across a range of fits", () => {
    for (const n of [3, 5, 9, 20]) {
      const pts = Array.from({ length: n }, (_, i) => ({ x: i, y: Math.sin(i) * 3 + i }));
      const r = trendStats(pts)!;
      expect(r.p!).toBeGreaterThanOrEqual(0);
      expect(r.p!).toBeLessThanOrEqual(1);
    }
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

describe("polyTrend — least-squares polynomial fit", () => {
  it("recovers an exact quadratic (R² = 1, correct evaluation)", () => {
    // y = 2x² − 3x + 1 sampled exactly.
    const pts = [-2, -1, 0, 1, 2, 3].map((x) => ({ x, y: 2 * x * x - 3 * x + 1 }));
    const fit = polyTrend(pts, 2)!;
    expect(fit.degree).toBe(2);
    expect(fit.r2).toBeCloseTo(1, 6);
    expect(fit.at(4)).toBeCloseTo(2 * 16 - 12 + 1, 4); // = 21
    expect(fit.at(-3)).toBeCloseTo(2 * 9 + 9 + 1, 4); // = 28
  });

  it("clamps the degree to points − 1 so it never interpolates noise", () => {
    // 3 points, ask for quartic → degree clamps to 2.
    const fit = polyTrend(
      [
        { x: 0, y: 0 },
        { x: 1, y: 1 },
        { x: 2, y: 3 },
      ],
      4,
    )!;
    expect(fit.degree).toBe(2);
  });

  it("returns a lower R² for an underfit (line through a parabola)", () => {
    const pts = [-2, -1, 0, 1, 2].map((x) => ({ x, y: x * x }));
    expect(polyTrend(pts, 1)!.r2).toBeLessThan(polyTrend(pts, 2)!.r2);
    expect(polyTrend(pts, 2)!.r2).toBeCloseTo(1, 6);
  });

  it("returns null for fewer than two points", () => {
    expect(polyTrend([{ x: 1, y: 1 }], 2)).toBeNull();
  });
});
