import { describe, expect, it } from "vitest";
import { buildChart, normalizeConfig } from "../src/core/chart";
import { formatNumber, formatPercent, niceTicks } from "../src/core/format";
import type { ChartConfig, ChartKind } from "../src/core/types";

/** Every numeric coordinate a scene node can carry, for finiteness checks. */
function coordsOf(n: import("../src/core/scene").SceneNode): number[] {
  switch (n.kind) {
    case "rect":
    case "text":
    case "chevron":
      return [n.x, n.y, n.w, n.h];
    case "line":
      return [n.x1, n.y1, n.x2, n.y2];
    case "ellipse":
      return [n.cx, n.cy, n.rx, n.ry];
    case "wedge":
      return [n.cx, n.cy, n.r, n.innerR, n.startAngle, n.endAngle];
    case "polygon":
      return n.points.flatMap((p) => [p.x, p.y]);
    case "symbol":
      return [n.cx, n.cy, n.size];
    case "arrowhead":
      return [n.x, n.y, n.angle, n.size];
  }
}

function expectAllFinite(cfg: ChartConfig) {
  const scene = buildChart(cfg);
  for (const n of scene.nodes) {
    for (const c of coordsOf(n)) expect(Number.isFinite(c)).toBe(true);
  }
  return scene;
}

describe("format hardening", () => {
  it("niceTicks repairs an inverted range", () => {
    const ticks = niceTicks(5, 3);
    expect(ticks.every(Number.isFinite)).toBe(true);
    expect(ticks[0]).toBeLessThanOrEqual(ticks[ticks.length - 1]);
  });

  it("niceTicks returns a safe default for non-finite bounds", () => {
    expect(niceTicks(NaN, 10)).toEqual([0, 1]);
    expect(niceTicks(0, Infinity)).toEqual([0, 1]);
    expect(niceTicks(-Infinity, Infinity)).toEqual([0, 1]);
  });

  it("formatNumber suppresses non-finite values instead of printing NaN", () => {
    expect(formatNumber(NaN)).toBe("");
    expect(formatNumber(Infinity)).toBe("");
    expect(formatNumber(-Infinity)).toBe("");
    // finite values are unaffected
    expect(formatNumber(1234, { decimals: 0 })).toBe("1,234");
  });

  it("formatPercent suppresses non-finite values", () => {
    expect(formatPercent(NaN)).toBe("");
    expect(formatPercent(0.5)).toBe("50%");
  });

  it("formatNumber falls back to en-US on an invalid locale", () => {
    expect(() => formatNumber(1234, { decimals: 0, locale: "not-a-real-locale!!" })).not.toThrow();
    expect(formatNumber(1234, { decimals: 0, locale: "not-a-real-locale!!" })).toBe("1,234");
  });
});

describe("normalizeConfig", () => {
  it("clamps zero/negative/NaN dimensions to the default size", () => {
    const base = { kind: "clustered", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } };
    expect(normalizeConfig({ ...base, width: 0, height: -5 } as ChartConfig).width).toBe(480);
    expect(normalizeConfig({ ...base, width: NaN, height: 300 } as ChartConfig).width).toBe(480);
    expect(normalizeConfig({ ...base, width: 640, height: 400 } as ChartConfig).width).toBe(640);
  });

  it("orders a reversed manual scale and drops non-finite ends", () => {
    const base = { kind: "clustered", width: 480, height: 300, data: { categories: ["A"], series: [{ name: "S", values: [1] }] } };
    expect(normalizeConfig({ ...base, scale: { min: 5, max: 3 } } as ChartConfig).scale).toEqual({ min: 3, max: 5 });
    expect(normalizeConfig({ ...base, scale: { min: NaN, max: 10 } } as ChartConfig).scale).toEqual({ max: 10 });
  });

  it("pads short series rows and nulls non-finite cells", () => {
    const cfg = normalizeConfig({
      kind: "clustered",
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B", "C"],
        series: [{ name: "S", values: [1, NaN as unknown as number] }],
      },
    } as ChartConfig);
    expect(cfg.data.series[0].values).toEqual([1, null, null]);
  });

  it("leaves a well-formed config's data content unchanged", () => {
    const cfg: ChartConfig = {
      kind: "stacked",
      width: 480,
      height: 300,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
    };
    const norm = normalizeConfig(cfg);
    expect(norm.data.categories).toEqual(["A", "B"]);
    expect(norm.data.series[0].values).toEqual([1, 2]);
    expect(norm.width).toBe(480);
  });
});

describe("buildChart survives malformed configs", () => {
  const good = (over: Partial<ChartConfig>): ChartConfig => ({
    kind: "clustered",
    width: 480,
    height: 300,
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 20, 30] }] },
    ...over,
  });

  it("reversed scale renders finite geometry", () => {
    expectAllFinite(good({ scale: { min: 100, max: 0 } }));
  });

  it("zero / negative / NaN size renders finite geometry", () => {
    expectAllFinite(good({ width: 0, height: 0 }));
    expectAllFinite(good({ width: -100, height: -100 }));
    expectAllFinite(good({ width: NaN, height: NaN }));
  });

  it("non-finite and null data cells render finite geometry", () => {
    expectAllFinite(
      good({
        data: {
          categories: ["A", "B", "C"],
          series: [{ name: "S", values: [Infinity as unknown as number, null, NaN as unknown as number] }],
        },
      }),
    );
  });

  it("series rows shorter/longer than categories do not throw", () => {
    expect(() =>
      buildChart(
        good({
          data: {
            categories: ["A", "B", "C", "D"],
            series: [
              { name: "Short", values: [1] },
              { name: "Long", values: [1, 2, 3, 4, 5, 6] },
            ],
          },
        }),
      ),
    ).not.toThrow();
  });

  it("an unknown kind falls back to columns without throwing", () => {
    expect(() => buildChart(good({ kind: "totally-bogus" as ChartKind }))).not.toThrow();
  });

  it("non-finite label offsets are dropped", () => {
    expectAllFinite(
      good({ decorations: { totals: true }, labelOffsets: { "total-0": { dx: NaN, dy: -8 } } }),
    );
  });

  it("empty data does not throw", () => {
    expect(() => buildChart(good({ data: { categories: [], series: [] } }))).not.toThrow();
  });
});

/**
 * Structural fuzz: the existing fuzz.test.ts randomizes VALUES inside a
 * well-formed shape. This randomizes the STRUCTURE — mismatched lengths,
 * missing rows, hostile cells — which is exactly what an authored/LLM-written
 * config or a hand-edited shape tag can carry. None may throw or emit NaN.
 */
function rng(seed: number) {
  return () => {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const KINDS: ChartKind[] = [
  "stacked", "clustered", "stacked100", "waterfall", "mekko", "line", "area",
  "butterfly", "combo", "pie", "doughnut", "boxplot", "radar", "heatmap",
  "scatter", "bubble", "cascade", "funnel", "waffle", "treemap", "sunburst",
];

describe("malformed-structure fuzz", () => {
  const rand = rng(20260718);
  for (let iter = 0; iter < 80; iter++) {
    const kind = KINDS[Math.floor(rand() * KINDS.length)];
    const nCats = Math.floor(rand() * 6); // may be 0
    const nSeries = Math.floor(rand() * 4); // may be 0
    const hostile = (): number | null => {
      const r = rand();
      if (r < 0.12) return null;
      if (r < 0.18) return NaN;
      if (r < 0.24) return Infinity;
      if (r < 0.3) return -Infinity;
      return Math.round((rand() - 0.4) * 300);
    };
    const cfg: ChartConfig = {
      kind,
      width: rand() < 0.15 ? 0 : 480,
      height: rand() < 0.15 ? -10 : 300,
      data: {
        categories: Array.from({ length: nCats }, (_, i) => (rand() < 0.2 ? "" : `C${i}`)),
        series: Array.from({ length: nSeries }, (_, s) => ({
          name: `S${s}`,
          // Row length deliberately mismatched against nCats.
          values: Array.from({ length: Math.floor(rand() * 8) }, hostile),
        })),
      },
      scale: rand() < 0.2 ? { min: 50, max: 10 } : undefined,
    };
    it(`#${iter} ${kind} ${nSeries}x${nCats}`, () => {
      let scene!: ReturnType<typeof buildChart>;
      expect(() => (scene = buildChart(cfg))).not.toThrow();
      for (const n of scene.nodes) {
        for (const c of coordsOf(n)) expect(Number.isFinite(c)).toBe(true);
      }
    });
  }
});
