import { describe, expect, it } from "vitest";
import {
  cagr,
  formatNumber,
  formatPercent,
  niceTicks,
  polyTrend,
  resolveAxisFormat,
  trendStats,
} from "../src/core/format";

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
    // ...in the locales whose sign glyph / digits a string match would miss.
    for (const locale of ["de-DE", "sv-SE", "ar-EG"]) {
      expect(formatPercent(-0.001, 0, false, locale), locale).not.toMatch(/[-−]/);
    }
  });
  it("honours the chart's locale, so one chart prints one number system", () => {
    // The funnel's conversion label sits beside formatNumber stage values; with
    // an en-US-only percent, a de-DE chart printed "12.000" next to "35.8%".
    expect(formatPercent(0.358, 1, false, "de-DE")).toBe("35,8 %");
    expect(formatPercent(0.149, 1, true, "de-DE")).toBe("+14,9 %");
    expect(formatNumber(12000, { decimals: 0, locale: "de-DE" })).toBe("12.000");
    // An unusable tag falls back to en-US rather than throwing (as formatNumber does).
    expect(formatPercent(0.358, 1, false, "not a locale")).toBe("35.8%");
  });
  it("groups a percentage past 1000%", () => {
    expect(formatPercent(35.5)).toBe("3,550%"); // was "3550%" beside formatNumber's "3,550"
  });
});

describe("resolveAxisFormat", () => {
  it("takes precision from the tick STEP, not the tick magnitude", () => {
    // Magnitude alone gave one decimal for a 0.01 step: ["7.4","7.5","7.5",…].
    expect(resolveAxisFormat([7.44, 7.45, 7.46, 7.47, 7.48]).decimals).toBe(2);
    expect(resolveAxisFormat([99, 99.5, 100, 100.5, 101]).decimals).toBe(1);
    expect(resolveAxisFormat([9.9, 9.95, 10, 10.05, 10.1]).decimals).toBe(2);
    // A coarse step keeps the magnitude precision (never coarser than before).
    expect(resolveAxisFormat([0, 0.1, 0.2, 0.3]).decimals).toBe(2);
    expect(resolveAxisFormat([0, 500, 1000]).decimals).toBe(0);
  });
  it("labels every tick distinctly and correctly", () => {
    for (const ticks of [
      [7.44, 7.45, 7.46, 7.47, 7.48],
      [88, 88.2, 88.4, 88.6, 88.8, 89],
      [1.02, 1.04, 1.06, 1.08, 1.1],
      [0.1, 1, 10, 100, 1000], // a log axis: 0.1 must not print as "0"
      niceTicks(1000.5, 1002.5, 5),
    ]) {
      const fmt = resolveAxisFormat(ticks);
      const labels = ticks.map((t) => formatNumber(t, fmt));
      expect(new Set(labels).size, labels.join("|")).toBe(new Set(ticks).size);
      for (let i = 0; i < ticks.length; i++) {
        expect(Number(labels[i].replace(/,/g, "")), labels[i]).toBeCloseTo(ticks[i], 9);
      }
    }
  });
  it("leaves an authored decimals count alone", () => {
    expect(resolveAxisFormat([7.44, 7.45], { decimals: 0 }).decimals).toBe(0);
  });
  it("falls back to magnitude when there is nothing to step between", () => {
    expect(resolveAxisFormat([]).decimals).toBe(2);
    expect(resolveAxisFormat([5, 5]).decimals).toBe(1);
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
  it("keeps ticks distinct past 12 significant digits", () => {
    // The FP-drift cleanup used to round to a fixed 12 significant digits, so an
    // axis around 1e13 collapsed onto ONE value: every gridline drawn on top of
    // the others and a top tick below the data max.
    for (const [min, max] of [
      [1e13, 1e13 + 4],
      [123456789012345, 123456789012350],
    ]) {
      const t = niceTicks(min, max, 5);
      expect(new Set(t).size, t.join("|")).toBe(t.length);
      expect(t[t.length - 1]).toBeGreaterThanOrEqual(max);
      expect(t[0]).toBeLessThanOrEqual(min);
    }
  });
  it("still cleans the FP drift of a fractional step", () => {
    expect(niceTicks(0, 0.5, 5)).toEqual([0, 0.2, 0.4, 0.6]);
    expect(niceTicks(-0.06, 0.06, 5)).toContain(0);
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

  it("clamps the degree to points − 2 so a residual dof always remains", () => {
    // A degree n−1 polynomial interpolates n points exactly (R²=1, meaningless),
    // so the fit must leave at least one degree of freedom: max degree = n−2.
    expect(
      polyTrend(
        [
          { x: 0, y: 0 },
          { x: 1, y: 1 },
          { x: 2, y: 3 },
          { x: 3, y: 4 },
        ],
        4,
      )!.degree,
    ).toBe(2); // 4 pts
    expect(
      polyTrend(
        [
          { x: 0, y: 0 },
          { x: 1, y: 1 },
          { x: 2, y: 3 },
        ],
        4,
      )!.degree,
    ).toBe(1); // 3 pts → line
    // With the clamp, a genuinely noisy set is NOT interpolated to R²=1.
    const noisy = [
      { x: 0, y: 1.1 },
      { x: 1, y: 1.9 },
      { x: 2, y: 3.2 },
      { x: 3, y: 3.8 },
      { x: 4, y: 5.3 },
    ];
    expect(polyTrend(noisy, 4)!.r2).toBeLessThan(1);
  });

  it("fits a small x-span instead of collapsing the pivot to null", () => {
    // Power sums scale as span^k, so a small span drove S[2d] under the solver's
    // 1e-12 pivot floor and the trendline silently vanished. Unit-scaling fixes it.
    const small = Array.from({ length: 8 }, (_, i) => ({ x: i * 0.001, y: (i * 0.001) ** 2 + 0.5 }));
    const fit = polyTrend(small, 3);
    expect(fit).not.toBeNull();
    expect(fit!.at(0.004)).toBeCloseTo(0.004 ** 2 + 0.5, 6);
    // A large span must still fit (centering intact).
    const large = Array.from({ length: 8 }, (_, i) => ({ x: 2e6 + i * 1000, y: i * i }));
    expect(polyTrend(large, 3)).not.toBeNull();
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
