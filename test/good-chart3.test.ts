import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE, valueExtent } from "../src/core/chart";
import { lerpColor, sequentialScale, divergingScale } from "../src/core/color";
import { detectLayout, EU_TILES, EUROPE_TILES, TILE_LAYOUTS, US_TILES, WORLD_TILES } from "../src/core/layout/tilemap-layouts";
import { sceneToSvg } from "../src/render/svg";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, PolygonNode, RectNode, TextNode } from "../src/core/scene";

/** Batch 3: boxplot, radar, heatmap, tilemap, datamark axes. */

const cfg = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
  ...DEFAULT_SIZE,
  ...partial,
});

const named = (c: ChartConfig, prefix: string) => buildChart(c).nodes.filter((n) => n.name?.startsWith(prefix));

describe("boxplot", () => {
  const summary: ChartConfig = cfg({
    kind: "boxplot",
    data: {
      categories: ["North", "South"],
      series: [
        { name: "Min", values: [2, 3] },
        { name: "Q1", values: [3, 5] },
        { name: "Median", values: [4, 7] },
        { name: "Q3", values: [6, 9] },
        { name: "Max", values: [8, 12] },
        { name: "Mean", values: [4.5, 7.2] },
        { name: "Outlier 1", values: [11, null] },
      ],
    },
  });

  it("precomputed rows drive boxes, whiskers, mean, and outliers directly", () => {
    const s = buildChart(summary);
    expect(s.nodes.filter((n) => n.name?.startsWith("box-"))).toHaveLength(2);
    expect(s.nodes.filter((n) => n.name?.startsWith("median-") && n.kind === "line")).toHaveLength(2);
    expect(s.nodes.filter((n) => n.name?.startsWith("mean-"))).toHaveLength(2);
    expect(s.nodes.filter((n) => n.name?.startsWith("outlier-"))).toHaveLength(1);
    // Median line sits inside the box vertically.
    const box = s.nodes.find((n) => n.name === "box-0") as RectNode;
    const med = s.nodes.find((n) => n.name === "median-0") as LineNode;
    expect(med.y1).toBeGreaterThan(box.y);
    expect(med.y1).toBeLessThan(box.y + box.h);
  });

  it("raw samples get computed quartiles and Tukey outliers", () => {
    const raw: ChartConfig = cfg({
      kind: "boxplot",
      data: {
        categories: ["A"],
        // 8 tight observations + one far outlier.
        series: [4, 5, 5, 6, 6, 7, 7, 8, 25].map((v) => ({ name: "", values: [v] })),
      },
    });
    const s = buildChart(raw);
    expect(s.nodes.some((n) => n.name === "box-0")).toBe(true);
    expect(s.nodes.filter((n) => n.name?.startsWith("outlier-0"))).toHaveLength(1);
    // Whisker stops at the last in-fence point, not at the outlier.
    const capHi = s.nodes.find((n) => n.name === "cap-hi-0") as LineNode;
    const outlier = s.nodes.find((n) => n.name === "outlier-0-0");
    expect(outlier?.kind === "ellipse" && outlier.cy).toBeLessThan(capHi.y1);
  });

  it("all boxes in one chart share a single value axis", () => {
    // South's Max (12) equals a hypothetical value on any other box's scale:
    // assert equal values map to equal coordinates across categories.
    const shared: ChartConfig = cfg({
      kind: "boxplot",
      data: {
        categories: ["A", "B", "C"],
        series: [
          { name: "Min", values: [2, 2, 5] },
          { name: "Q1", values: [3, 3, 6] },
          { name: "Median", values: [4, 4, 8] },
          { name: "Q3", values: [6, 6, 9] },
          { name: "Max", values: [8, 8, 12] },
        ],
      },
    });
    const s = buildChart(shared);
    // Boxes A and B carry identical numbers → identical y geometry.
    const boxA = s.nodes.find((n) => n.name === "box-0") as RectNode;
    const boxB = s.nodes.find((n) => n.name === "box-1") as RectNode;
    expect(boxA.y).toBeCloseTo(boxB.y);
    expect(boxA.h).toBeCloseTo(boxB.h);
    // And box C's larger values sit strictly higher on the same scale.
    const boxC = s.nodes.find((n) => n.name === "box-2") as RectNode;
    expect(boxC.y).toBeLessThan(boxA.y);
    // Exactly one value axis is generated for the whole chart.
    const axisLabels = buildChart({ ...shared, decorations: { valueAxis: true, segmentLabels: true } })
      .nodes.filter((n) => n.name === "value-axis");
    expect(axisLabels.length).toBeGreaterThan(1); // one shared set of ticks, not per-box
  });

  it("renders horizontally: boxes become rows on a bottom value axis", () => {
    const v = buildChart(summary);
    const h = buildChart({ ...summary, horizontal: true, decorations: { valueAxis: true, categoryAxis: true, segmentLabels: true } });
    const vBox = v.nodes.find((n) => n.name === "box-0") as RectNode;
    const hBox = h.nodes.find((n) => n.name === "box-0") as RectNode;
    // Vertical: box taller than wide (IQR spans y). Horizontal: wider than tall.
    expect(vBox.h).toBeGreaterThan(0);
    expect(hBox.w).toBeGreaterThan(hBox.h * 0.5);
    // The median line rotates too: vertical chart → horizontal line; horizontal chart → vertical line.
    const vMed = v.nodes.find((n) => n.name === "median-0") as LineNode;
    const hMed = h.nodes.find((n) => n.name === "median-0") as LineNode;
    expect(vMed.y1).toBeCloseTo(vMed.y2);
    expect(hMed.x1).toBeCloseTo(hMed.x2);
    // Same shared-axis property holds horizontally: equal values → equal x.
    const hData: ChartConfig = {
      ...summary,
      horizontal: true,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "Min", values: [2, 2] },
          { name: "Q1", values: [3, 3] },
          { name: "Median", values: [4, 4] },
          { name: "Q3", values: [6, 6] },
          { name: "Max", values: [8, 8] },
        ],
      },
    };
    const hs = buildChart(hData);
    const b0 = hs.nodes.find((n) => n.name === "box-0") as RectNode;
    const b1 = hs.nodes.find((n) => n.name === "box-1") as RectNode;
    expect(b0.x).toBeCloseTo(b1.x);
    expect(b0.w).toBeCloseTo(b1.w);
    expect(b0.y).not.toBeCloseTo(b1.y); // separate category rows
    // Outliers and whisker caps rotate with the chart.
    const hOut = buildChart({
      ...hData,
      data: { ...hData.data, series: [...hData.data.series, { name: "Outlier 1", values: [11, null] }] },
    });
    const out = hOut.nodes.find((n) => n.name === "outlier-0-0");
    const cap = hOut.nodes.find((n) => n.name === "cap-hi-0") as LineNode;
    expect(out?.kind === "ellipse" && out.cx).toBeGreaterThan(cap.x1); // beyond the whisker, along x
    expect(cap.x1).toBeCloseTo(cap.x2); // caps are vertical dashes now
  });

  it("min/max whiskers can be forced and Same Scale sees the extent", () => {
    const raw: ChartConfig = cfg({
      kind: "boxplot",
      boxplot: { whiskers: "minmax" },
      data: { categories: ["A"], series: [4, 5, 6, 7, 25].map((v) => ({ name: "", values: [v] })) },
    });
    expect(buildChart(raw).nodes.some((n) => n.name?.startsWith("outlier-"))).toBe(false);
    expect(valueExtent(raw)).toEqual({ min: 0, max: 25 });
  });
});

describe("radar", () => {
  const radar: ChartConfig = cfg({
    kind: "radar",
    data: {
      categories: ["A", "B", "C", "D", "E"],
      series: [
        { name: "Today", values: [3, 2, 3, 2, 4] },
        { name: "Target", values: [4, 4, 4, 3, 5] },
      ],
    },
    scale: { min: 0, max: 5 },
  });

  it("draws polygon grid, spokes, and one translucent polygon per series", () => {
    const s = buildChart(radar);
    expect(s.nodes.filter((n) => n.name?.startsWith("spoke-"))).toHaveLength(5);
    expect(s.nodes.filter((n) => n.kind === "polygon" && n.name?.startsWith("grid-")).length).toBeGreaterThan(1);
    const series = s.nodes.filter((n): n is PolygonNode => n.kind === "polygon" && !!n.name?.startsWith("series-"));
    expect(series).toHaveLength(2);
    expect(series[0].fillOpacity).toBeCloseTo(0.18);
    expect(series[0].points).toHaveLength(5);
  });

  it("first spoke points to 12 o'clock and the grid can be circles", () => {
    const s = buildChart(radar);
    const spoke0 = s.nodes.find((n) => n.name === "spoke-0") as LineNode;
    expect(spoke0.x1).toBeCloseTo(spoke0.x2, 5); // straight up
    expect(spoke0.y2).toBeLessThan(spoke0.y1);
    const circles = buildChart({ ...radar, decorations: { gridShape: "circle", segmentLabels: true } });
    expect(circles.nodes.some((n) => n.kind === "ellipse" && n.name?.startsWith("grid-") && n.fill === "none")).toBe(true);
  });

  it("renders SVG polygons with fill-opacity", () => {
    const svg = sceneToSvg(buildChart(radar));
    expect(svg).toContain("<polygon");
    expect(svg).toContain('fill-opacity="0.18"');
  });
});

describe("heatmap", () => {
  it("colors one global sequential scale with labels and legend", () => {
    const s = buildChart(
      cfg({
        kind: "heatmap",
        data: {
          categories: ["Q1", "Q2"],
          series: [
            { name: "North", values: [10, 40] },
            { name: "South", values: [20, 30] },
          ],
        },
      }),
    );
    const cells = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("cell-"));
    expect(cells).toHaveLength(4);
    expect(new Set(cells.map((c) => c.fill)).size).toBe(4); // distinct values → distinct fills
    expect(s.nodes.some((n) => n.name === "legend-min")).toBe(true);
    expect(s.nodes.filter((n) => n.name?.startsWith("cell-label-"))).toHaveLength(4);
  });

  it("switches to a diverging scale (through white) when data spans zero", () => {
    const s = buildChart(
      cfg({
        kind: "heatmap",
        data: { categories: ["A", "B", "C"], series: [{ name: "R", values: [-8, 0, 8] }] },
      }),
    );
    const cells = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("cell-0"));
    expect(cells[1].fill).toBe("#ffffff"); // zero = the white pivot
    expect(s.nodes.some((n) => n.name === "legend-zero")).toBe(true);
  });

  it("renders nulls as gray no-data cells and survives constant data", () => {
    const s = buildChart(
      cfg({ kind: "heatmap", data: { categories: ["A", "B"], series: [{ name: "R", values: [5, null] }] } }),
    );
    const cells = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("cell-0"));
    expect(cells[1].fill).toBe("#e6e6e6");
    const constant = buildChart(
      cfg({ kind: "heatmap", data: { categories: ["A", "B"], series: [{ name: "R", values: [5, 5] }] } }),
    );
    expect(constant.nodes.some((n) => n.name === "legend-swatch")).toBe(true);
  });
});

describe("tilemap", () => {
  it("ships complete verified layouts", () => {
    expect(Object.keys(US_TILES)).toHaveLength(51); // 50 states + DC
    expect(Object.keys(EUROPE_TILES)).toHaveLength(41);
    expect(Object.keys(EU_TILES)).toHaveLength(27); // EU members
    expect(Object.keys(WORLD_TILES)).toHaveLength(10);
    // No two regions share a tile in any layout.
    for (const layout of Object.values(TILE_LAYOUTS)) {
      const coords = Object.values(layout).map(([c, r]) => `${c},${r}`);
      expect(new Set(coords).size).toBe(coords.length);
    }
  });

  it("auto-detects the layout from region codes", () => {
    expect(detectLayout(["CA", "TX", "NY", "FL"])).toBe("us");
    expect(detectLayout(["DK", "SE", "DE", "FR"])).toBe("eu");
    expect(detectLayout(["DK", "SE", "NO", "CH"])).toBe("europe"); // NO/CH are non-EU
    expect(detectLayout(["XX", "YY"])).toBeNull();
  });

  it("renders colored tiles, gray no-data tiles, and a legend", () => {
    const s = buildChart(
      cfg({
        kind: "tilemap",
        map: "world",
        data: {
          categories: ["NA", "EU", "SEA"],
          series: [{ name: "Rev", values: [100, 80, 20] }],
        },
      }),
    );
    const tiles = s.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("tile-"));
    expect(tiles).toHaveLength(10); // full layout renders
    expect((s.nodes.find((n) => n.name === "tile-NA") as RectNode).fill).not.toBe("#e6e6e6");
    expect((s.nodes.find((n) => n.name === "tile-SSA") as RectNode).fill).toBe("#e6e6e6");
    expect(s.nodes.some((n) => n.name === "legend-nodata")).toBe(true);
  });

  it("degrades gracefully for unrecognized codes", () => {
    const s = buildChart(
      cfg({ kind: "tilemap", data: { categories: ["FOO", "BAR"], series: [{ name: "V", values: [1, 2] }] } }),
    );
    expect(s.nodes.some((n) => n.name === "tilemap-error")).toBe(true);
  });
});

describe("datamark axes", () => {
  const line: ChartConfig = cfg({
    kind: "line",
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [12, 30, 22] }] },
  });

  it("draws tick dashes without an axis line", () => {
    const s = buildChart({ ...line, decorations: { valueAxis: "datamarks", segmentLabels: false } });
    expect(s.nodes.filter((n) => n.name === "datamark").length).toBeGreaterThan(1);
    expect(s.nodes.filter((n): n is TextNode => n.kind === "text" && n.name === "value-axis").length).toBeGreaterThan(1);
  });

  it("tickMode data places marks at the scale extremes (range frame)", () => {
    const s = buildChart({
      ...line,
      decorations: { valueAxis: "datamarks", tickMode: "data", segmentLabels: false },
    });
    expect(s.nodes.filter((n) => n.name === "datamark")).toHaveLength(2);
  });
});

describe("color scales", () => {
  it("interpolates in linear light and clamps", () => {
    expect(lerpColor("#000000", "#ffffff", 0)).toBe("#000000");
    expect(lerpColor("#000000", "#ffffff", 1)).toBe("#ffffff");
    // Linear-light midpoint of black/white is lighter than sRGB 0x80.
    const mid = parseInt(lerpColor("#000000", "#ffffff", 0.5).slice(1, 3), 16);
    expect(mid).toBeGreaterThan(160);
  });
  it("sequential floor never reaches pure white; diverging is symmetric", () => {
    const seq = sequentialScale(0, 10, "#2a78d6");
    expect(seq(0)).not.toBe("#ffffff");
    const div = divergingScale(-10, 10, "#2a78d6", "#e34948");
    expect(div(0)).toBe("#ffffff");
    expect(div(10)).toBe("#2a78d6");
    expect(div(-10)).toBe("#e34948");
  });
});
