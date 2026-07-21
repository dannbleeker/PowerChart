import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { divergingScale, noDataFill, sequentialScale } from "../src/core/color";
import { logFloor } from "../src/core/layout/frame";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";

const DARK = { background: "#1b1b1b", text: "#f2f1ec" };

const axisLabels = (cfg: ChartConfig) =>
  buildChart(cfg)
    .nodes.filter((n): n is TextNode => n.kind === "text" && n.name === "value-axis")
    .map((n) => n.text);

const rects = (cfg: ChartConfig, prefix: string) =>
  buildChart(cfg).nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith(prefix));

describe("log axis floors on the data, not three decades below its own max", () => {
  // Both callers derive dataMin with a zero seed, so valueScale's `dataMin > 0`
  // branch was unreachable and every log axis fell back to dataMax / 1000.
  const bars: ChartConfig = {
    kind: "clustered",
    logScale: true,
    decorations: { valueAxis: true },
    data: { categories: ["a", "b", "c"], series: [{ name: "S", values: [200, 260, 300] }] },
  } as ChartConfig;

  it("gives a column chart only the decades its data occupies", () => {
    expect(axisLabels(bars)).toEqual(["100", "1,000"]);
  });

  it("gives a line chart the same floor", () => {
    expect(axisLabels({ ...bars, kind: "line" } as ChartConfig)).toEqual(["100", "1,000"]);
  });

  it("spends the plot on the data instead of on empty decades", () => {
    // The phantom decades squashed 200–300 into the top sliver of the plot: the
    // shortest and tallest column differed by ~4% of the plot height.
    const tops = rects(bars, "seg-").map((r) => r.y);
    const spread = Math.max(...tops) - Math.min(...tops);
    expect(spread).toBeGreaterThan(30);
  });

  it("falls back to dataMin when nothing is positive", () => {
    expect(logFloor([-3, 0, -1], -3)).toBe(-3);
    expect(logFloor([], 0)).toBe(0);
    expect(logFloor([0.5, 4, 90], 0)).toBe(0.5);
  });

  it("leaves a linear axis alone", () => {
    expect(axisLabels({ ...bars, logScale: false } as ChartConfig)).toEqual(["0", "100", "200", "300"]);
  });
});

describe("value scales resolve against the canvas, not a hardcoded white", () => {
  const heat: ChartConfig = {
    kind: "heatmap",
    data: { categories: ["x", "y"], series: [{ name: "r1", values: [1, 9] }] },
  } as ChartConfig;

  it("repaints heatmap cells for a dark canvas", () => {
    const light = rects(heat, "cell-").map((r) => r.fill);
    const dark = rects({ ...heat, style: DARK } as ChartConfig, "cell-").map((r) => r.fill);
    expect(light.length).toBeGreaterThan(0);
    expect(dark).not.toEqual(light);
    // The low end must sink INTO the dark slide, not glare off it.
    expect(dark[0]).not.toBe(light[0]);
    expect(Number.parseInt(dark[0].slice(1, 3), 16)).toBeLessThan(0x60);
  });

  it("repaints tilemap tiles for a dark canvas", () => {
    const tiles: ChartConfig = {
      kind: "tilemap",
      data: { categories: ["CA", "TX", "NY"], series: [{ name: "v", values: [10, 90, 50] }] },
    } as ChartConfig;
    const light = rects(tiles, "tile-").map((r) => r.fill);
    const dark = rects({ ...tiles, style: DARK } as ChartConfig, "tile-").map((r) => r.fill);
    expect(light.length).toBeGreaterThan(0);
    expect(dark).not.toEqual(light);
  });

  it("sinks a diverging zero into the canvas", () => {
    expect(divergingScale(-1, 1, "#2a78d6", "#e34948", "#1b1b1b")(0)).toBe("#1b1b1b");
    expect(divergingScale(-1, 1, "#2a78d6", "#e34948")(0)).toBe("#ffffff");
  });

  it("darkens the no-data fill on a dark canvas and keeps it on a light one", () => {
    expect(noDataFill("#ffffff")).toBe("#e6e6e6");
    const dark = noDataFill("#1b1b1b");
    expect(dark).not.toBe("#e6e6e6");
    expect(Number.parseInt(dark.slice(1, 3), 16)).toBeLessThan(0x40);
  });

  it("leaves light-canvas output byte-identical", () => {
    expect(sequentialScale(0, 10, "#2a78d6")(0)).toBe(sequentialScale(0, 10, "#2a78d6", "#ffffff")(0));
    expect(sequentialScale(0, 10, "#2a78d6")(0)).toBe("#f1f4fb");
  });
});
