import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Funnel — band geometry, short-frame fit, conversion markers. */

/**
 * Degenerate-frame / degenerate-scale guards found by a layout bug-hunt. Each is
 * byte-identical for normal charts (snapshots unchanged) and only repairs an edge
 * that previously emitted negative geometry or NaN coordinates.
 */
const negDims = (nodes: ReturnType<typeof buildChart>["nodes"]) =>
  nodes.filter(
    (n) =>
      ((n.kind === "rect" || n.kind === "text") && ((n as RectNode).w < 0 || (n as RectNode).h < 0)) ||
      Object.entries(n).some(([k, v]) => ["x", "y", "w", "h"].includes(k) && typeof v === "number" && Number.isNaN(v)),
  );

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */
const W = 480;

const H = 300;

const texts = (nodes: SceneNode[], namePrefix: string) =>
  nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith(namePrefix)).map((n) => n.text);

const base = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  ...DEFAULT_SIZE,
  width: W_offframeguards,
  height: H,
  data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 20, 30] }] },
  ...partial,
});

/**
 * Geometry that must stay inside the canvas. These were found by the review as
 * off-frame / wrong-size layout bugs the loose fuzz bound (|c| < 5000) hid.
 */
const W_offframeguards = 480;

/** Backlog batch B: funnel kind + lollipop/dot/range bar styles. */
describe("funnel", () => {
  const s = buildChart(sampleConfig("funnel"));

  it("draws centered bands with width proportional to value", () => {
    const bands = [0, 1, 4].map((c) => s.nodes.find((n) => n.name === `stage-${c}`) as RectNode);
    // Widths proportional: 720/1200, 120/1200.
    expect(bands[1].w / bands[0].w).toBeCloseTo(720 / 1200, 2);
    expect(bands[2].w / bands[0].w).toBeCloseTo(120 / 1200, 2);
    // Centered: all bands share the same center x.
    const cx = (r: RectNode) => r.x + r.w / 2;
    expect(cx(bands[1])).toBeCloseTo(cx(bands[0]), 5);
    expect(cx(bands[2])).toBeCloseTo(cx(bands[0]), 5);
  });

  it("labels conversion vs the previous stage between bands", () => {
    const conv = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("conversion-"));
    expect(conv).toHaveLength(4);
    expect(conv[0].text).toContain("60.0%"); // 720/1200
    // Stage names on the left, values on/beside the bands.
    expect(s.nodes.some((n) => n.name === "category-0")).toBe(true);
    expect((s.nodes.find((n) => n.name === "stage-value-4") as TextNode).text).toBe("120");
  });

  it("narrow bands put their value beside the band, wide ones inside", () => {
    const wide = s.nodes.find((n) => n.name === "stage-value-0") as TextNode;
    const band0 = s.nodes.find((n) => n.name === "stage-0") as RectNode;
    expect(wide.x).toBeGreaterThanOrEqual(band0.x); // inside
    const tiny = buildChart({
      ...sampleConfig("funnel"),
      data: { categories: ["All", "Won"], series: [{ name: "Deals", values: [10000, 12] }] },
    });
    const narrow = tiny.nodes.find((n) => n.name === "stage-value-1") as TextNode;
    const band = tiny.nodes.find((n) => n.name === "stage-1") as RectNode;
    expect(narrow.x).toBeGreaterThan(band.x + band.w); // beside
  });
});

describe("funnel bands never go negative on a short, crowded frame", () => {
  it("floors band height", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: 640,
      height: 60, // 20 stages + 1.5em gaps can't fit → bands went negative
      data: {
        categories: Array.from({ length: 20 }, (_, i) => `S${i}`),
        series: [{ name: "v", values: Array.from({ length: 20 }, (_, i) => 20 - i) }],
      },
    };
    expect(negDims(buildChart(cfg).nodes)).toHaveLength(0);
  });
});

describe("funnel conversion marker follows the direction of the step", () => {
  it("marks a rise with ▴ and a fall with ▾", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: W,
      height: H,
      // Ascending — the pyramid ordering funnel.ts recommends.
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 50, 0] }] },
    };
    // Before: "▾ 500.0%" — a down arrow on a 5x increase.
    expect(texts(buildChart(cfg).nodes, "conversion-")).toEqual(["▴ 500.0%", "▾ 0.0%"]);
  });

  it("drops the marker when the stage is unchanged", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: W,
      height: H,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 10] }] },
    };
    expect(texts(buildChart(cfg).nodes, "conversion-")).toEqual(["100.0%"]);
  });
});

describe("funnel bands fit a short frame", () => {
  it("many stages on a short frame stay inside the canvas", () => {
    const scene = buildChart(
      base({
        kind: "funnel",
        height: 60, // deliberately short: a fixed gap used to overshoot the bottom
        data: { categories: ["a", "b", "c", "d", "e", "f"], series: [{ name: "S", values: [60, 50, 40, 30, 20, 10] }] },
      }),
    );
    const stages = scene.nodes.filter((n) => /^stage-\d+$/.test(n.name ?? "")); // the bands, not stage-value-*
    expect(stages.length).toBe(6);
    for (const s of stages as any[]) expect(s.y + s.h).toBeLessThanOrEqual(60 + 1);
  });
});
