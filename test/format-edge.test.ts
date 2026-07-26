import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { formatNumber, formatPercent, segmentLabel } from "../src/core/format";
import type { TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { resolveFormat } from "../src/core/format";

/** Number/label formatting edge cases — exact zero, locale, out-of-range decimals. */

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

const byName = (nodes: { name?: string }[], p: string) => nodes.filter((n) => n.name?.startsWith(p));

describe("formatNumber negative-zero", () => {
  it("normalises a rounded -0 to 0", () => {
    expect(formatNumber(-0.4, { decimals: 0 })).toBe("0");
    expect(formatNumber(-0.001, { decimals: 0 })).toBe("0");
  });
  it("keeps the sign on genuine negatives", () => {
    expect(formatNumber(-5.2)).toBe("-5.2");
    expect(formatNumber(-0.4, { decimals: 0, forceSign: true })).toBe("0");
  });
});

/**
 * Bugs surfaced by the codebase bug-hunt. Engine/layout tier.
 */
describe("formatNumber / formatPercent tolerate out-of-range decimals", () => {
  it("repairs a negative decimals count instead of throwing RangeError", () => {
    // `toFixed(-1)` throws; the whole render used to abort on a hand-edited config.
    expect(() => formatNumber(5, { decimals: -1 })).not.toThrow();
    expect(formatNumber(5, { decimals: -1 })).toBe("5");
  });

  it("repairs an absurdly large decimals count", () => {
    expect(() => formatNumber(5, { decimals: 500 })).not.toThrow();
    expect(() => formatPercent(0.5, 500)).not.toThrow();
  });

  it("clamps a negative percent decimals count", () => {
    expect(() => formatPercent(0.5, -3)).not.toThrow();
    expect(formatPercent(0.5, -3)).toBe("50%");
  });
});

describe("formatNumber renders an exact zero without a fractional digit", () => {
  it("prints 0, not 0.0, under auto decimals", () => {
    expect(formatNumber(0)).toBe("0");
  });
  it("still prints sub-unit magnitudes at 2 decimals", () => {
    expect(formatNumber(0.25)).toBe("0.25");
  });
});

describe("label content & locale", () => {
  it("builds multi-part labels", () => {
    const label = segmentLabel(["series", "value"], {
      value: 12,
      fraction: 0.5,
      series: "SMB",
      category: "2024",
      fmt: { decimals: 0 },
    });
    expect(label).toBe("SMB 12");
  });
  it("applies segment label content in stacked charts", () => {
    const scene = buildChart(
      cfg({
        data: {
          categories: ["A"],
          series: [
            { name: "S1", values: [60] },
            { name: "S2", values: [40] },
          ],
        },
        decorations: { labelContent: ["value", "percent"] },
      }),
    );
    const label = byName(scene.nodes, "label-0-0")[0] as TextNode;
    expect(label.text).toBe("60 60%");
  });
  it("formats with a locale", () => {
    expect(formatNumber(1234.5, { decimals: 1, locale: "de-DE" })).toBe("1.234,5");
  });
});

describe("label content everywhere", () => {
  it("line point labels honor labelContent", () => {
    const scene = buildChart(
      cfg({
        kind: "line",
        data: { categories: ["A"], series: [{ name: "GM", values: [42] }] },
        decorations: { segmentLabels: true, labelContent: ["series", "value"] },
      }),
    );
    const label = byName(scene.nodes, "label-0-0")[0] as TextNode;
    expect(label.text).toBe("GM 42");
  });
  it("scatter labels can include coordinates", () => {
    const scene = buildChart(
      cfg({
        kind: "scatter",
        data: {
          categories: ["P"],
          series: [
            { name: "X", values: [10] },
            { name: "Y", values: [20] },
          ],
        },
        decorations: { labelContent: ["category", "value"] },
      }),
    );
    const label = byName(scene.nodes, "label-0")[0] as TextNode;
    expect(label.text).toBe("P (10, 20)");
  });
});

describe("segmentLabel percent without a fraction", () => {
  it("drops the percent part when no denominator exists", () => {
    const label = segmentLabel(["value", "percent"], {
      value: 5,
      fraction: null,
      series: "S",
      category: "A",
      fmt: resolveFormat([5]),
    });
    expect(label).toBe("5.0");
  });
});

/**
 * `formatPercent` gained a locale so a chart could not print two number systems
 * side by side ("12.000" beside "35.8%"), but only the axis/segment path passed
 * it. The CAGR and difference arrows and four layouts' own percent labels still
 * called it bare, so a localized chart mixed systems whenever those were on.
 *
 * Asserted by DIFFERENCE against the same chart with no locale, rather than
 * against a hard-coded German string: some of these labels carry no decimals
 * ("56 %"), where the only tell is the narrow space de-DE puts before the sign.
 * Comparing the two renders catches every such difference and stays honest if
 * a label's decimal count changes.
 */
describe("every percent label follows the chart's locale", () => {
  const pcts = (c: ChartConfig) =>
    buildChart(c)
      .nodes.filter((n): n is TextNode => n.kind === "text" && n.text.includes("%"))
      .map((n) => `${n.name}=${n.text}`);

  const expectLocalized = (base: Partial<ChartConfig>, where: string) => {
    const build = (locale?: string) =>
      pcts({ ...base, ...DEFAULT_SIZE, numberFormat: { decimals: 1, locale } } as unknown as ChartConfig);
    const plain = build();
    const german = build("de-DE");
    expect(plain.length, `${where}: no percent label found — fixture is vacuous`).toBeGreaterThan(0);
    expect(german.length, where).toBe(plain.length);
    // Every percent label must have moved; one that did not is a bare call site.
    const unchanged = german.filter((g, i) => g === plain[i]);
    expect(unchanged, `${where}: not localized — ${unchanged.join(", ")}`).toEqual([]);
  };

  const trend = { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 20, 33.7] }] };

  it("threads it through the CAGR arrow", () => {
    expectLocalized({ kind: "stacked", data: trend, decorations: { cagr: { from: 0, to: 2 } } }, "cagr");
  });

  it("threads it through the difference arrow", () => {
    expectLocalized(
      { kind: "stacked", data: trend, decorations: { difference: { from: 0, to: 2, percent: true } } },
      "difference",
    );
  });

  it.each(["funnel", "cascade", "mekko", "waffle"] as const)("threads it through %s labels", (kind) => {
    expectLocalized(
      {
        kind,
        data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [55.5, 33.3, 11.2] }] },
        decorations: { segmentLabels: true },
      },
      kind,
    );
  });
});
