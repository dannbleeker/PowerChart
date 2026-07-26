import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { RectNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Decoration layout — bands stay clipped to the plot, anchors, variance tiers, corners. */

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */
const W = 480;

const H = 300;

function cfg(partial: Partial<ChartConfig>): ChartConfig {
  return { kind: "stacked", width: 480, height: 300, data: { categories: [], series: [] }, ...partial };
}

const base = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  ...DEFAULT_SIZE,
  width: W_offframeguards,
  height: H,
  data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 20, 30] }] },
  ...partial,
});

/** Every rect stays within the canvas (a small overhang margin for edge labels). */
const inFrame = (nodes: { kind: string }[], margin = 40) => {
  for (const n of nodes as any[]) {
    if (n.kind !== "rect") continue;
    expect(n.y, `rect top ${n.name}`).toBeGreaterThanOrEqual(-margin);
    expect(n.y + n.h, `rect bottom ${n.name}`).toBeLessThanOrEqual(H + margin);
    expect(n.x, `rect left ${n.name}`).toBeGreaterThanOrEqual(-margin);
    expect(n.x + n.w, `rect right ${n.name}`).toBeLessThanOrEqual(W_offframeguards + margin);
  }
};

/**
 * Geometry that must stay inside the canvas. These were found by the review as
 * off-frame / wrong-size layout bugs the loose fuzz bound (|c| < 5000) hid.
 */
const W_offframeguards = 480;

describe("decoration anchors never fall back to a value the scale excludes", () => {
  it("candlestick: a blank High anchors on the period's own prices", () => {
    const cfg: ChartConfig = {
      kind: "candlestick",
      width: W,
      height: H,
      data: {
        categories: ["Mon", "Tue", "Wed"],
        series: [
          { name: "Open", values: [100, 102, 103] },
          { name: "High", values: [105, null, 106] },
          { name: "Low", values: [99, 101, 102] },
          { name: "Close", values: [102, 103, 105] },
        ],
      },
      decorations: { callouts: [{ text: "gap", category: 1 }] },
    };
    // Before: toY(0) on a zero-free OHLC scale put the callout box at y ≈ 3452.
    const box = buildChart(cfg).nodes.find((n) => n.name === "callout-box-0") as RectNode;
    expect(box.y).toBeGreaterThanOrEqual(0);
    expect(box.y + box.h).toBeLessThanOrEqual(H);
  });

  it("violin: a category with no observations anchors on the plot floor", () => {
    const cfg: ChartConfig = {
      kind: "violin",
      width: W,
      height: H,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "o1", values: [50, null] },
          { name: "o2", values: [60, null] },
          { name: "o3", values: [70, null] },
        ],
      },
      decorations: { valueAxis: true, callouts: [{ text: "x", category: 1 }] },
    };
    // Before: Math.max(...[0]) on a 50–70 domain put the callout at y ≈ 915.
    const box = buildChart(cfg).nodes.find((n) => n.name === "callout-box-0") as RectNode;
    expect(box.y).toBeGreaterThanOrEqual(0);
    expect(box.y + box.h).toBeLessThanOrEqual(H);
  });
});

describe("decoration corners", () => {
  it("computes CAGR on a single series", () => {
    const scene = buildChart(
      cfg({
        data: {
          categories: ["Y1", "Y2"],
          series: [
            { name: "A", values: [100, 121] },
            { name: "B", values: [500, 400] },
          ],
        },
        decorations: { cagr: { from: 0, to: 1, series: 0 } },
      }),
    );
    const label = scene.nodes.find((n) => n.name === "cagr-label") as TextNode;
    expect(label.text).toContain("+21.0%");
  });

  it("anchors difference arrows at a value line", () => {
    const scene = buildChart(
      cfg({
        data: { categories: ["A"], series: [{ name: "S", values: [150] }] },
        decorations: {
          valueLines: [{ mode: "value", value: 100 }],
          difference: { from: 0, to: 0, fromValueLine: 0, percent: false },
        },
      }),
    );
    const label = scene.nodes.find((n) => n.name === "diff-label") as TextNode;
    expect(label.text).toBe("+50");
  });
});

describe("decoration bands stay clipped to the plot", () => {
  it("a y-band whose range exceeds the value domain does not render off-frame", () => {
    // valueToY extrapolates past the axis; an unclamped band ran off the canvas.
    const scene = buildChart(base({ decorations: { bands: [{ axis: "y", from: -10_000, to: 10_000 }] } }));
    const band = scene.nodes.find((n) => n.name === "band-0") as any;
    expect(band).toBeTruthy();
    inFrame(scene.nodes);
  });

  it("a band entirely outside the domain is dropped, not drawn with 0/neg height", () => {
    const scene = buildChart(base({ decorations: { bands: [{ axis: "y", from: 1e6, to: 2e6 }] } }));
    // Off-plot band collapses to h<=0 and is filtered out.
    expect(scene.nodes.find((n) => n.name === "band-0")).toBeUndefined();
  });
});

describe("variance tier reserves height only where it is drawn", () => {
  it("a line chart does not lose plot height to a variance strip it never draws", () => {
    const off = buildChart(base({ kind: "line", decorations: {} }));
    const on = buildChart(base({ kind: "line", decorations: { variance: { actual: 0, reference: 0 } } }));
    // No variance nodes on a line chart either way…
    expect(on.nodes.some((n) => n.name?.startsWith("variance"))).toBe(false);
    // …and the markers occupy the same vertical extent (no phantom reservation).
    const yOf = (s: any) => s.nodes.filter((n: any) => n.name?.startsWith("marker-")).map((n: any) => n.y);
    expect(Math.max(...yOf(on))).toBeCloseTo(Math.max(...yOf(off)), 5);
  });

  it("a stacked column chart still reserves and draws the variance tier", () => {
    const cfg = base({ kind: "stacked", decorations: { variance: { actual: 0, reference: 0 } } });
    const scene = buildChart(cfg);
    expect(scene.nodes.some((n) => n.name?.startsWith("variance"))).toBe(true);
  });
});

/**
 * A callout bubble hovers 4.2 font sizes above its anchor. When the anchor is at
 * the plot ceiling there is no room there, and box, text and tail all landed off
 * the top of the canvas. 100% and Mekko charts expose no `valueToY`, so EVERY
 * callout on them falls back to `columnTop` — which on those kinds is the plot
 * ceiling for every column by construction, making the failure unconditional.
 */
describe("callouts stay on the canvas", () => {
  const withCallout = (over: Partial<ChartConfig>) =>
    buildChart({
      ...DEFAULT_SIZE,
      width: 480,
      height: 300,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "x", values: [3, 4] },
          { name: "y", values: [1, 2] },
        ],
      },
      decorations: { callouts: [{ category: 0, text: "worth a note" }] },
      ...over,
    } as unknown as ChartConfig);

  it.each(["stacked100", "mekko", "stacked"] as const)("fits the bubble on a %s chart", (kind) => {
    const s = withCallout({ kind });
    const off = (s.nodes as { name?: string; x?: number; y?: number; w?: number; h?: number }[])
      .filter((n) => n.name?.startsWith("callout-"))
      .filter(
        (n) =>
          (n.y ?? 0) < -0.5 ||
          (n.y ?? 0) + (n.h ?? 0) > s.height + 0.5 ||
          (n.x ?? 0) < -0.5 ||
          (n.x ?? 0) + (n.w ?? 0) > s.width + 0.5,
      )
      .map((n) => `${n.name}@${n.x?.toFixed(1)},${n.y?.toFixed(1)}`);
    expect(off, `off-canvas: ${off.join(", ")}`).toEqual([]);
  });

  it("still hovers above an anchor that has room", () => {
    const s = withCallout({ kind: "stacked", decorations: { callouts: [{ category: 0, text: "note" }] } });
    const box = (s.nodes as { name?: string; y?: number }[]).find((n) => n.name === "callout-box-0")!;
    const anchor = (s.nodes as { name?: string; y?: number }[]).find((n) => n.name === "seg-0-0")!;
    expect(box.y!).toBeLessThan(anchor.y!);
  });
});

describe("difference arrow anchored to a value line", () => {
  const c = (partial: Partial<ChartConfig>): ChartConfig => ({
    kind: "stacked",
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "S1", values: [10, 20, 30] },
        { name: "S2", values: [5, 5, 5] },
      ],
    },
    ...DEFAULT_SIZE,
    ...partial,
  });
  const texts = (s: ReturnType<typeof buildChart>) => s.nodes.filter((n): n is TextNode => n.kind === "text");

  const pair = { from: 0, to: 2 };
  const hasDiffLabel = (decorations: ChartConfig["decorations"]) =>
    texts(buildChart(c({ decorations }))).some((t) => t.name === "diff-label");

  it("anchors at a mean value line", () => {
    expect(
      hasDiffLabel({
        segmentLabels: true,
        valueLines: [{ mode: "mean" }],
        difference: { ...pair, fromValueLine: 0 },
      }),
    ).toBe(true);
  });

  it("anchors at a fixed value line", () => {
    expect(
      hasDiffLabel({
        segmentLabels: true,
        valueLines: [{ mode: "value", value: 20 }],
        difference: { ...pair, fromValueLine: 0 },
      }),
    ).toBe(true);
  });

  it("normalizes the deprecated single valueLine field", () => {
    const s = buildChart(
      c({
        decorations: {
          segmentLabels: true,
          valueLine: { mode: "mean" },
          difference: { ...pair, fromValueLine: 0 },
        },
      }),
    );
    expect(s.nodes.some((n) => n.name?.startsWith("value-line"))).toBe(true);
  });
});
