import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { formatNumber, formatPercent, segmentLabel } from "../src/core/format";
import type { TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

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
