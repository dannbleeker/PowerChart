import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { LineNode, RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Line / area — slope, stepped, smoothed, confidence bands, missing-data, profiles. */

describe("forecast styling on lines", () => {
  const line: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["2023", "2024", "2025", "2026"], series: [{ name: "Rev", values: [40, 46, 52, 60] }] },
    decorations: { forecastFrom: 2, segmentLabels: false },
  };

  it("dashes segments and hollows markers from the boundary on", () => {
    const s = buildChart(line);
    const segs = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    expect(segs.find((l) => l.name === "line-0-1")!.dash).toBeUndefined(); // actuals
    expect(segs.find((l) => l.name === "line-0-2")!.dash).toEqual([4, 3]); // into forecast
    expect(segs.find((l) => l.name === "line-0-3")!.dash).toEqual([4, 3]);
    const solid = s.nodes.find((n) => n.name === "marker-0-1") as RectNode;
    const hollow = s.nodes.find((n) => n.name === "marker-0-2") as RectNode;
    expect(solid.fill).not.toBe("#ffffff");
    expect(hollow.fill).toBe("#ffffff");
    expect(s.nodes.some((n) => n.name === "forecast-divider")).toBe(true);
  });

  it("is inert without the option", () => {
    const s = buildChart({ ...line, decorations: { segmentLabels: false } });
    expect(s.nodes.some((n) => n.name === "forecast-divider")).toBe(false);
    expect(
      s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-")).every((l) => !l.dash),
    ).toBe(true);
  });
});

describe("line confidence bands and fill-between", () => {
  const line: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: {
      categories: ["2024", "2025", "2026"],
      series: [
        { name: "Forecast", values: [50, 56, 63] },
        { name: "Band low", values: [50, 52, 55] },
        { name: "Band high", values: [50, 60, 72] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("shades a ribbon between Band low/high without drawing them as lines", () => {
    const s = buildChart(line);
    const slabs = s.nodes.filter((n) => n.name?.startsWith("band-ribbon"));
    expect(slabs.length).toBeGreaterThan(20);
    // Only the forecast line renders; band rows are not series.
    expect(s.nodes.some((n) => n.name === "line-1-1")).toBe(false);
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    // Ribbon renders behind the line (earlier in the node list).
    expect(s.nodes.findIndex((n) => n.name?.startsWith("band-ribbon"))).toBeLessThan(
      s.nodes.findIndex((n) => n.name === "line-0-1"),
    );
    // The band widens the scale: max tick covers 72.
    const slab = s.nodes.find((n) => n.name?.startsWith("band-ribbon-1")) as RectNode;
    expect(slab.y).toBeGreaterThan(0);
  });

  it("fillBetween shades the gap between two series", () => {
    const s = buildChart({
      kind: "line",
      ...DEFAULT_SIZE,
      data: {
        categories: ["Q1", "Q2", "Q3"],
        series: [
          { name: "Plan", values: [40, 50, 60] },
          { name: "Actual", values: [38, 46, 52] },
        ],
      },
      decorations: { fillBetween: [0, 1], segmentLabels: false },
    });
    expect(s.nodes.filter((n) => n.name?.startsWith("fill-between")).length).toBeGreaterThan(20);
    // Both series still draw as lines.
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    expect(s.nodes.some((n) => n.name === "line-1-1")).toBe(true);
  });
});

/** Backlog batch E: slope chart option, waffle kind, KPI tile element. */
describe("slope chart (decorations.slope on line)", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: {
      categories: ["2020", "2025"],
      series: [
        { name: "Brand A", values: [48, 66] },
        { name: "Brand B", values: [40, 49] },
      ],
    },
    decorations: { slope: true },
  };
  const s = buildChart(cfg);

  it("draws end rails, straight lines, and 'Name value' labels at both ends", () => {
    expect(s.nodes.filter((n) => n.name?.startsWith("slope-rail"))).toHaveLength(2);
    expect((s.nodes.find((n) => n.name === "slope-left-0") as TextNode).text).toBe("Brand A 48");
    expect((s.nodes.find((n) => n.name === "slope-right-0") as TextNode).text).toBe("Brand A 66");
    const line = s.nodes.find((n) => n.name === "line-0-1") as LineNode;
    expect(line.y1).toBeGreaterThan(line.y2); // 48 → 66 rises
    // No value axis or gridlines in slope mode.
    expect(s.nodes.some((n) => n.name?.startsWith("grid"))).toBe(false);
    expect(s.nodes.some((n) => n.name?.startsWith("vaxis"))).toBe(false);
  });

  it("labels take the series color and de-overlap when ends are close", () => {
    const crowded = buildChart({
      ...cfg,
      data: {
        categories: ["2020", "2025"],
        series: [
          { name: "Alpha", values: [50, 60] },
          { name: "Beta", values: [50.5, 60.5] },
          { name: "Gamma", values: [51, 61] },
        ],
      },
    });
    const ys = [0, 1, 2]
      .map((si) => (crowded.nodes.find((n) => n.name === `slope-left-${si}`) as TextNode).y)
      .sort((a, b) => a - b);
    expect(ys[1] - ys[0]).toBeGreaterThanOrEqual(14);
    expect(ys[2] - ys[1]).toBeGreaterThanOrEqual(14);
    const label = s.nodes.find((n) => n.name === "slope-left-0") as TextNode;
    const line = s.nodes.find((n) => n.name === "line-0-1") as LineNode;
    expect(label.color).toBe(line.stroke);
  });

  it("handles >2 categories as a polyline and ignores slope on area charts", () => {
    const three = buildChart({
      ...cfg,
      data: {
        categories: ["2020", "2022", "2025"],
        series: [{ name: "A", values: [40, 55, 66] }],
      },
    });
    expect(three.nodes.filter((n) => n.name?.startsWith("line-0-"))).toHaveLength(2);
    expect(three.nodes.filter((n) => n.name?.startsWith("slope-rail"))).toHaveLength(2); // ends only
    const area = buildChart({ ...cfg, kind: "area" });
    expect(area.nodes.some((n) => n.name?.startsWith("slope-rail"))).toBe(false);
  });
});

/**
 * Backlog batch G — gaps within existing kinds: stepped line/area,
 * Excel-style column gap width & clustered overlap, butterfly value ticks.
 */
describe("stepped line", () => {
  const base: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "v", values: [1, 2, 3] }] },
    decorations: { segmentLabels: false },
  };

  it("default draws one sloped connector per interval", () => {
    const plain = buildChart(base);
    const seg = plain.nodes.find((n) => n.name === "line-0-1") as LineNode;
    expect(seg).toBeTruthy();
    // Diagonal: neither horizontal nor vertical.
    expect(seg.y1).not.toBeCloseTo(seg.y2, 5);
    expect(seg.x1).not.toBeCloseTo(seg.x2, 5);
  });

  it('"after" holds then jumps (HV elbow)', () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, stepped: "after" } });
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(false);
    const a = s.nodes.find((n) => n.name === "line-0-1a") as LineNode;
    const b = s.nodes.find((n) => n.name === "line-0-1b") as LineNode;
    expect(a.y1).toBeCloseTo(a.y2, 5); // horizontal at the previous value
    expect(b.x1).toBeCloseTo(b.x2, 5); // then a vertical jump
    expect(b.x1).toBeCloseTo(a.x2, 5); // meeting at the next category x
  });

  it('"before" jumps immediately, "center" steps at the midpoint', () => {
    const before = buildChart({ ...base, decorations: { segmentLabels: false, stepped: "before" } });
    const ba = before.nodes.find((n) => n.name === "line-0-1a") as LineNode;
    expect(ba.x1).toBeCloseTo(ba.x2, 5); // vertical first

    const center = buildChart({ ...base, decorations: { segmentLabels: false, stepped: "center" } });
    const ca = center.nodes.find((n) => n.name === "line-0-1a") as LineNode;
    const cb = center.nodes.find((n) => n.name === "line-0-1b") as LineNode;
    const cc = center.nodes.find((n) => n.name === "line-0-1c") as LineNode;
    expect([ca, cb, cc].every(Boolean)).toBe(true);
    expect(cb.x1).toBeCloseTo((ca.x1 + cc.x2) / 2, 5); // riser at the midpoint
  });
});

describe("stepped area", () => {
  const base: ChartConfig = {
    kind: "area",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "v", values: [10, 20, 30] }] },
    decorations: { segmentLabels: false },
  };

  it("default interpolates the slab top; stepped holds it flat", () => {
    const slabs = (scene: ReturnType<typeof buildChart>) =>
      scene.nodes.filter((n) => n.name?.startsWith("area-0-0-")) as RectNode[];

    const plain = slabs(buildChart(base));
    // A sloped (interpolated) top: the first and last slab of the segment differ.
    expect(plain.length).toBeGreaterThan(1);
    expect(Math.abs(plain[0].y - plain[plain.length - 1].y)).toBeGreaterThan(0.5);

    // A stepped area has a flat top across the interval, so the slab-fill needs
    // no tessellation at all — the segment collapses to a single slab.
    const stepped = slabs(buildChart({ ...base, decorations: { segmentLabels: false, stepped: "before" } }));
    expect(stepped.length).toBe(1);
  });
});

/**
 * Backlog batch H — more §2 within-kind gaps: area with negative values,
 * scatter/bubble trajectory trail, boxplot jittered raw-data dots.
 */
describe("area with negative values", () => {
  const areaRects = (sc: { nodes: SceneNode[] }) =>
    sc.nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("area-"));
  const hasNegAxisLabel = (sc: { nodes: SceneNode[] }) =>
    sc.nodes.some((n) => n.kind === "text" && n.name === "value-axis" && /^-/.test((n.text ?? "").trim()));
  // The value axis "0" tick sits at toY(0) — the zero baseline in screen y.
  const zeroY = (sc: { nodes: SceneNode[] }) => {
    const z = sc.nodes.find(
      (n): n is TextNode => n.kind === "text" && n.name === "value-axis" && (n.text ?? "").trim() === "0",
    )!;
    return z.y + z.h / 2;
  };
  const mk = (values: (number | null)[], cats = ["A", "B", "C"]): ChartConfig => ({
    kind: "area",
    ...DEFAULT_SIZE,
    data: { categories: cats, series: [{ name: "v", values }] },
    decorations: { segmentLabels: false, valueAxis: true, gridlines: true },
  });

  it("fills both above and below the zero baseline", () => {
    const s = buildChart(mk([8, -8], ["A", "B"]));
    const y0 = zeroY(s);
    const rects = areaRects(s);
    expect(rects.some((r) => r.y + r.h > y0 + 1)).toBe(true); // fill extending below zero
    expect(rects.some((r) => r.y < y0 - 1)).toBe(true); // fill extending above zero
  });

  it("extends the value axis into the negatives (not clamped to 0)", () => {
    expect(hasNegAxisLabel(buildChart(mk([10, -6, 8])))).toBe(true);
  });

  it("positive-only area stays entirely above the baseline", () => {
    const pos = buildChart(mk([10, 6, 8]));
    expect(hasNegAxisLabel(pos)).toBe(false);
    const y0 = zeroY(pos);
    expect(areaRects(pos).every((r) => r.y + r.h <= y0 + 1)).toBe(true); // nothing below zero
  });
});

describe("smoothed lines", () => {
  const base: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "v", values: [1, 4, 2, 5] }] },
    decorations: { segmentLabels: false },
  };

  it("default draws straight connectors", () => {
    const s = buildChart(base);
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
    expect(s.nodes.some((n) => n.name?.startsWith("line-0-1-s"))).toBe(false);
  });

  it("smooth replaces straight segments with a sampled spline polyline", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, smooth: true } });
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(false);
    const sampled = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.match(/^line-0-\d+-s\d+$/));
    // 3 segments × 16 samples.
    expect(sampled.length).toBe(48);
  });

  it("stepped wins over smooth (mutually exclusive)", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, smooth: true, stepped: "after" } });
    expect(s.nodes.some((n) => n.name?.startsWith("line-0-1-s"))).toBe(false);
    expect(s.nodes.some((n) => n.name === "line-0-1a")).toBe(true);
  });
});

describe("transparent floating segments", () => {
  const cfg: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "base", color: "transparent", values: [10, 14] },
        { name: "Range", values: [20, 18] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("draws no rect for the transparent segment but still stacks above it", () => {
    const s = buildChart(cfg);
    expect(s.nodes.some((n) => n.name === "seg-0-0")).toBe(false); // transparent base not drawn
    const range = s.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    expect(range).toBeTruthy();
    const solid = buildChart({
      ...cfg,
      data: {
        ...cfg.data,
        series: [
          { name: "base", values: [10, 14] },
          { name: "Range", values: [20, 18] },
        ],
      },
    });
    const solidRange = solid.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    // Same height, but the floating one sits higher (smaller y) is false — same y
    // since the level is identical; the difference is the missing base rect.
    expect(range.y).toBeCloseTo(solidRange.y, 3);
    expect(range.h).toBeCloseTo(solidRange.h, 3);
    // The floating build is missing the base segment the solid build draws.
    expect(solid.nodes.some((n) => n.name === "seg-0-0")).toBe(true);
  });

  it("floating bar does not reach the baseline", () => {
    const s = buildChart(cfg);
    const range = s.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    // Baseline is the bottom of the plot; the bar's bottom (y+h) is above it
    // by the (undrawn) base segment. Compare Q1 (base 10) vs a base-0 build.
    const grounded = buildChart({
      ...cfg,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "base", color: "transparent", values: [0, 0] },
          { name: "Range", values: [20, 18] },
        ],
      },
    });
    const gRange = grounded.nodes.find((n): n is RectNode => n.name === "seg-1-0")!;
    expect(range.y + range.h).toBeLessThan(gRange.y + gRange.h - 1); // floats above the grounded bar's base
  });
});

describe("line missing-data bridge", () => {
  const base: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C", "D"], series: [{ name: "v", values: [10, null, 30, 40] }] },
    decorations: { segmentLabels: false },
  };

  it("breaks the line at nulls by default", () => {
    const s = buildChart(base);
    const lines = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    // Only C→D connects (A is isolated by the null B, B has no point).
    expect(lines).toHaveLength(1);
    expect(s.nodes.filter((n) => n.name?.startsWith("marker-0-"))).toHaveLength(3); // A, C, D
  });

  it("bridges across nulls when bridgeGaps is set", () => {
    const s = buildChart({ ...base, decorations: { segmentLabels: false, bridgeGaps: true } });
    const lines = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("line-0-"));
    // A→C (bridging B) and C→D → two segments.
    expect(lines).toHaveLength(2);
    // The bridge segment joins A's point to C's point.
    const bridge = lines.find((l) => l.name === "line-0-2")!;
    const cx = (i: number) => s.nodes.find((n): n is RectNode => n.name === `marker-0-${i}`)!.x;
    expect(bridge.x1).toBeCloseTo(cx(0) + 2.4, 0); // near A's marker
    expect(bridge.x2).toBeCloseTo(cx(2) + 2.4, 0); // to C's marker
  });
});

/**
 * Backlog batch P — the deferred §2-tail items: horizontal profile chart
 * (line/area) and radar vertex markers (verified in the scene → add-in path).
 */
describe("horizontal profile chart — line", () => {
  const cfg: ChartConfig = {
    kind: "line",
    ...DEFAULT_SIZE,
    horizontal: true,
    data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [20, 80, 50] }] },
    decorations: { segmentLabels: false, categoryAxis: true, valueAxis: true },
  };
  const s = buildChart(cfg);
  const marker = (c: number) => s.nodes.find((n): n is RectNode => n.name === `marker-0-${c}`)!;

  it("runs categories down the y axis (top to bottom)", () => {
    expect(marker(0).y).toBeLessThan(marker(1).y);
    expect(marker(1).y).toBeLessThan(marker(2).y);
  });

  it("maps larger values further right (value along x)", () => {
    expect(marker(1).x).toBeGreaterThan(marker(0).x); // 80 right of 20
    expect(marker(2).x).toBeGreaterThan(marker(0).x); // 50 right of 20
    expect(marker(1).x).toBeGreaterThan(marker(2).x); // 80 right of 50
  });

  it("connects the points with line segments", () => {
    expect(s.nodes.some((n) => n.name === "line-0-1")).toBe(true);
  });

  it("the vertical line chart is unaffected (no horizontal dispatch)", () => {
    const v = buildChart({ ...cfg, horizontal: undefined });
    // In the vertical chart categories run along x, so markers share no single column.
    const m0 = v.nodes.find((n): n is RectNode => n.name === "marker-0-0")!;
    const m1 = v.nodes.find((n): n is RectNode => n.name === "marker-0-1")!;
    expect(m1.x).toBeGreaterThan(m0.x); // categories advance along x when vertical
  });
});

describe("horizontal profile chart — stacked area", () => {
  const cfg: ChartConfig = {
    kind: "area",
    ...DEFAULT_SIZE,
    horizontal: true,
    data: {
      categories: ["N", "S"],
      series: [
        { name: "A", values: [40, 55] },
        { name: "B", values: [25, 20] },
      ],
    },
    decorations: { seriesLabels: true, categoryAxis: true, valueAxis: true },
  };
  const s = buildChart(cfg);

  it("stacks the series along x (second series to the right of the first)", () => {
    // Compare the same y-strip (slab k=0) of each series: B stacks to the
    // right of A there (globally the x-ranges overlap across categories).
    const a0 = s.nodes.find((n): n is RectNode => n.name === "area-0-0-0")!;
    const b0 = s.nodes.find((n): n is RectNode => n.name === "area-1-0-0")!;
    expect(a0).toBeTruthy();
    expect(b0).toBeTruthy();
    expect(b0.x).toBeGreaterThanOrEqual(a0.x + a0.w - 1);
  });
});
