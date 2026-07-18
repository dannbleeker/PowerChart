import { describe, expect, it } from "vitest";
import { formatNumber, formatPercent } from "../src/core/format";
import { paletteColor } from "../src/core/style";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { EllipseNode, WedgeNode } from "../src/core/scene";

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

describe("paletteColor wraps by the palette's actual length", () => {
  it("indexes a short custom palette by its length, never % 8", () => {
    const pal = ["#111111", "#222222", "#333333"];
    expect(paletteColor(pal, 4)).toBe("#222222"); // 4 % 3 = 1, not undefined
    expect(paletteColor(pal, 3)).toBe("#111111");
  });
  it("guards an empty palette", () => {
    expect(paletteColor([], 0)).toBe("#888888");
  });
});

describe("scatter group colouring honours a short palette", () => {
  it("gives a group beyond the palette length a real colour, not 'undefined'", () => {
    const cfg: ChartConfig = {
      kind: "scatter",
      width: 480,
      height: 320,
      style: { palette: ["#111111", "#222222", "#333333"] } as ChartConfig["style"],
      data: {
        categories: ["a", "b"],
        series: [
          { name: "X", values: [1, 2] },
          { name: "Y", values: [1, 2] },
          { name: "Group", values: [1, 5] }, // group id 5, past the 3-colour palette
        ],
      },
    };
    const fills = buildChart(cfg)
      .nodes.filter((n) => n.name?.startsWith("point-"))
      .map((n) => (n as { fill?: string }).fill);
    expect(fills.length).toBeGreaterThan(0);
    for (const f of fills) {
      expect(f).toBeTruthy();
      expect(f).not.toBe("undefined");
      expect(f).toMatch(/^#[0-9a-fA-F]{6}$/);
    }
  });
});

describe("pie radius never goes negative on a narrow frame", () => {
  it("floors the radius so wedges and the hole keep valid geometry", () => {
    const cfg: ChartConfig = {
      kind: "doughnut",
      width: 80, // width*0.5 - fs*7 = -30 before the floor
      height: 320,
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [3, 4, 5] }] },
    };
    const nodes = buildChart(cfg).nodes;
    for (const n of nodes) {
      if (n.kind === "wedge") expect((n as WedgeNode).r).toBeGreaterThanOrEqual(0);
      if (n.kind === "ellipse") {
        expect((n as EllipseNode).rx).toBeGreaterThanOrEqual(0);
        expect((n as EllipseNode).ry).toBeGreaterThanOrEqual(0);
      }
    }
  });
});
