import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import type { EllipseNode, LineNode, RectNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Columns — clustered/stacked, gap width, overlap, stacked-100 negatives, bar styles. */

/** Regression tests for the deferred combo / stacked100 / small-multiples fixes. */
const hasNaN = (nodes: { [k: string]: unknown }[]) =>
  nodes.some((n) => Object.values(n).some((v) => typeof v === "number" && Number.isNaN(v)));

describe("bar styles on clustered", () => {
  const base: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B", "C"],
      series: [
        { name: "2024", values: [40, 55, 48] },
        { name: "2025", values: [52, 60, 45] },
      ],
    },
  };
  const styled = (barStyle: "lollipop" | "dot" | "range") =>
    buildChart({ ...base, decorations: { barStyle, segmentLabels: true } });

  it("lollipop: stem from the baseline + dot at the value", () => {
    const s = styled("lollipop");
    const stems = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("stem-"));
    const dots = s.nodes.filter((n): n is EllipseNode => n.kind === "ellipse" && !!n.name?.startsWith("seg-"));
    expect(stems).toHaveLength(6);
    expect(dots).toHaveLength(6);
    // Stem ends at the dot; no rectangles drawn for the data.
    expect(stems[0].y2).toBeCloseTo(dots[0].cy, 5);
    expect(s.nodes.some((n) => n.kind === "rect" && n.name?.startsWith("seg-"))).toBe(false);
  });

  it("dot: dots only, no stems", () => {
    const s = styled("dot");
    expect(s.nodes.some((n) => n.name?.startsWith("stem-"))).toBe(false);
    expect(s.nodes.filter((n) => n.kind === "ellipse" && n.name?.startsWith("seg-"))).toHaveLength(6);
  });

  it("range: two series' dots joined by a connector on one shared x", () => {
    const s = styled("range");
    const ranges = s.nodes.filter((n): n is LineNode => n.kind === "line" && !!n.name?.startsWith("range-"));
    expect(ranges).toHaveLength(3);
    const d0 = s.nodes.find((n) => n.name === "seg-0-0") as EllipseNode;
    const d1 = s.nodes.find((n) => n.name === "seg-1-0") as EllipseNode;
    expect(d0.cx).toBeCloseTo(d1.cx, 5); // dumbbell: same x per category
    expect(ranges[0].y1).toBeCloseTo(d0.cy, 5);
    expect(ranges[0].y2).toBeCloseTo(d1.cy, 5);
  });

  it("stays plain bars by default and on stacked charts", () => {
    expect(buildChart(base).nodes.some((n) => n.kind === "rect" && n.name === "seg-0-0")).toBe(true);
    const stacked = buildChart({
      ...base,
      kind: "stacked",
      decorations: { barStyle: "lollipop", segmentLabels: true },
    });
    expect(stacked.nodes.some((n) => n.name?.startsWith("stem-"))).toBe(false);
  });
});

describe("column gap width", () => {
  const base: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B", "C"], series: [{ name: "v", values: [5, 8, 3] }] },
    decorations: { segmentLabels: false },
  };

  it("gapWidth 0 makes columns touch (1.5× the default width)", () => {
    const def = (buildChart(base).nodes.find((n) => n.name === "seg-0-0") as RectNode).w;
    const touch = (buildChart({ ...base, gapWidth: 0 }).nodes.find((n) => n.name === "seg-0-0") as RectNode).w;
    // Default gapWidth 50 → colThick = slot·2/3; gapWidth 0 → colThick = slot.
    expect(touch / def).toBeCloseTo(1.5, 2);
  });

  it("large gapWidth thins the columns", () => {
    const def = (buildChart(base).nodes.find((n) => n.name === "seg-0-0") as RectNode).w;
    const thin = (buildChart({ ...base, gapWidth: 300 }).nodes.find((n) => n.name === "seg-0-0") as RectNode).w;
    expect(thin).toBeLessThan(def);
  });
});

describe("clustered overlap", () => {
  const base: ChartConfig = {
    kind: "clustered",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "s1", values: [5, 6] },
        { name: "s2", values: [7, 8] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  const stride = (cfg: ChartConfig) => {
    const s = buildChart(cfg);
    const a = s.nodes.find((n) => n.name === "seg-0-0") as RectNode;
    const b = s.nodes.find((n) => n.name === "seg-1-0") as RectNode;
    return { d: b.x - a.x, w: a.w, ax: a.x, bx: b.x };
  };

  it("overlap 0 reproduces the historical edge-to-edge layout", () => {
    const r = stride(base);
    // Two bars filling the column: stride equals a bar's full width (w+gap).
    expect(r.d).toBeGreaterThan(0);
    expect(r.d).toBeCloseTo(r.w + 1, 5);
  });

  it("positive overlap widens bars and shrinks the stride; 100 fully overlaps", () => {
    const zero = stride(base);
    const forty = stride({ ...base, overlap: 40 });
    expect(forty.w).toBeGreaterThan(zero.w);
    expect(forty.d).toBeLessThan(zero.d);
    const full = stride({ ...base, overlap: 100 });
    expect(full.ax).toBeCloseTo(full.bx, 5); // same position
  });

  it("negative overlap opens a gap between bars", () => {
    const zero = stride(base);
    const neg = stride({ ...base, overlap: -50 });
    expect(neg.d).toBeGreaterThan(zero.d);
  });
});

describe("stacked100 negative values", () => {
  const cfg: ChartConfig = {
    kind: "stacked100",
    ...DEFAULT_SIZE,
    data: {
      categories: ["Q1", "Q2"],
      series: [
        { name: "New", values: [60, 70] },
        { name: "Renewal", values: [40, 45] },
        { name: "Returns", values: [-15, -10] },
      ],
    },
    decorations: { segmentLabels: false },
  };

  it("renders the negative series below the zero line (not clamped away)", () => {
    const s = buildChart(cfg);
    const seg = (nm: string) => s.nodes.find((n): n is RectNode => n.kind === "rect" && n.name === nm)!;
    const returns = seg("seg-2-0");
    expect(returns).toBeTruthy(); // the negative segment is drawn
    const newSeg = seg("seg-0-0"); // bottom of the positive stack, sits at the zero line
    // Returns starts at/below the zero line and extends further down.
    expect(returns.y).toBeGreaterThanOrEqual(newSeg.y + newSeg.h - 1);
    expect(returns.y + returns.h).toBeGreaterThan(newSeg.y + newSeg.h + 1);
  });

  it("positive-only stacked100 is unchanged (fills exactly to the top)", () => {
    const pos = buildChart({
      ...cfg,
      data: {
        categories: ["Q1"],
        series: [
          { name: "A", values: [60] },
          { name: "B", values: [40] },
        ],
      },
    });
    // Two segments, no negative region.
    expect(pos.nodes.filter((n) => n.name?.match(/^seg-\d+-0$/))).toHaveLength(2);
  });
});

describe("stacked100 with an all-negative category", () => {
  it("fills the segments downward instead of collapsing to zero", () => {
    const cfg: ChartConfig = {
      kind: "stacked100",
      ...DEFAULT_SIZE,
      data: {
        categories: ["A"],
        series: [
          { name: "P", values: [-30] },
          { name: "Q", values: [-20] },
        ],
      },
    };
    const scene = buildChart(cfg);
    const segs = scene.nodes.filter(
      (n) => n.kind === "rect" && (n.name ?? "").startsWith("seg") && (n as { h: number }).h > 0.5,
    );
    expect(segs.length).toBe(2); // both shares visible (were 0-height before)
    expect(hasNaN(scene.nodes as never)).toBe(false);
  });
});

describe("clustered level-anchored decorations (regression)", () => {
  // seriesLevels only advanced in the STACKED branch, so on a clustered chart it
  // was an array of zeros — while decor.ts still opted into "stack level" mode
  // because the array existed. The difference arrow silently collapsed onto the
  // baseline and read "0".
  const base = {
    width: 480,
    height: 300,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "S1", values: [100, 150] },
        { name: "S2", values: [50, 50] },
      ],
    },
    decorations: { difference: { from: 0, to: 1, series: 0 } },
  };

  it.each(["stacked", "clustered"])("%s anchors the arrow to the series mark, not the axis", (kind) => {
    const s = buildChart({ ...base, kind } as unknown as ChartConfig);
    const label = s.nodes.find((n) => n.name?.startsWith("diff-label")) as { text?: string } | undefined;
    const line = s.nodes.find((n) => n.name?.startsWith("diff-line")) as { y1?: number; y2?: number } | undefined;
    expect(label?.text, `${kind} difference label`).toBe("+50%");
    expect(line, `${kind} difference line`).toBeTruthy();
    expect(Math.abs(line!.y1! - line!.y2!), `${kind} arrow collapsed to zero length`).toBeGreaterThan(1);
  });
});

describe("stacked100 value axis (regression)", () => {
  const cfg = (negatives: boolean): ChartConfig =>
    ({
      kind: "stacked100",
      width: 480,
      height: 300,
      data: {
        categories: ["Q1", "Q2"],
        series: [
          { name: "New", values: [60, 70] },
          { name: "Renewal", values: [40, 45] },
          ...(negatives ? [{ name: "Returns", values: [-40, -10] }] : []),
        ],
      },
      decorations: { valueAxis: true, gridlines: true },
    }) as unknown as ChartConfig;

  it.each([false, true])("keeps every tick and gridline on the canvas (negatives=%s)", (negatives) => {
    const s = buildChart(cfg(negatives));
    const offCanvas = s.nodes
      .map((n) => ({ n, y: (n as { y?: number }).y ?? (n as { y1?: number }).y1 }))
      .filter(({ y }) => typeof y === "number" && (y < -0.5 || y > s.height + 0.5))
      .map(({ n, y }) => `${n.name}@${y!.toFixed(1)}`);
    // niceTicks rounds OUTWARD from pctNegMin, and nothing filtered it back into
    // the domain — a gridline landed at y=301.6 on a 300pt canvas.
    expect(offCanvas, `off-canvas: ${offCanvas.join(", ")}`).toEqual([]);
  });

  it("labels the share axis in percent, each label naming its own tick", () => {
    const labels = (buildChart(cfg(false)).nodes as { name?: string; text?: string }[])
      .filter((n) => n.name === "value-axis")
      .map((n) => n.text);
    // Was 0.0 / 0.3 / 0.5 / 0.8 / 1.0 — fractions on a chart whose segments read
    // "60%", and 0.25/0.75 rounded to labels that name no tick on the axis.
    expect(labels).toEqual(["0%", "25%", "50%", "75%", "100%"]);
  });

  it.each([false, true])("keeps the unit suffix off the share axis (horizontal=%s)", (horizontal) => {
    // A share is unitless. numberFormat.suffix is how an author says "millions"
    // — routed through the axis it labelled the 100% strip "25 m%", while the
    // segments beside it correctly read "25%".
    const labels = (
      buildChart({
        ...cfg(false),
        horizontal,
        numberFormat: { suffix: " m", decimals: "auto" },
      } as unknown as ChartConfig).nodes as { name?: string; text?: string }[]
    )
      .filter((n) => n.name === "value-axis")
      .map((n) => n.text);
    expect(labels).toEqual(["0%", "25%", "50%", "75%", "100%"]);
  });
});

/**
 * An authored "100% =" denominator can be SMALLER than the column it normalises
 * (a multi-select survey — "100% = 500 respondents", answers totalling 180% — or
 * a typo). The 100% scale was hard-pinned to max: 1, so those segments were
 * painted outside the plot AND outside the chart frame. `waffle` already handles
 * the same input (test/spatial-layouts.test.ts).
 */
describe("stacked100 honors an inconsistent 100%= denominator", () => {
  const s = buildChart({
    kind: "stacked100",
    ...DEFAULT_SIZE,
    width: 480,
    height: 300,
    data: {
      categories: ["A"],
      series: [
        { name: "x", values: [300] },
        { name: "y", values: [600] },
      ],
      hundredPercent: [500],
    },
  } as unknown as ChartConfig);

  it("keeps the over-100% stack inside the canvas", () => {
    const off = (s.nodes as { name?: string; x?: number; y?: number; w?: number; h?: number }[])
      .filter((n) => n.name?.startsWith("seg-"))
      .filter((n) => (n.y ?? 0) < -0.5 || (n.y ?? 0) + (n.h ?? 0) > s.height + 0.5)
      .map((n) => `${n.name}@${n.y?.toFixed(1)}`);
    expect(off, `off-canvas: ${off.join(", ")}`).toEqual([]);
  });

  it("keeps the segments proportional to their shares", () => {
    const seg = (name: string) => (s.nodes as { name?: string; h?: number }[]).find((n) => n.name === name)!;
    // 600 : 300 — the taller segment is exactly twice the shorter one.
    expect(seg("seg-1-0").h! / seg("seg-0-0").h!).toBeCloseTo(2, 1);
  });

  it("leaves a well-formed 100% chart on the plain 0/25/50/75/100 strip", () => {
    const ok = buildChart({
      kind: "stacked100",
      ...DEFAULT_SIZE,
      width: 480,
      height: 300,
      data: {
        categories: ["A"],
        series: [
          { name: "x", values: [3] },
          { name: "y", values: [7] },
        ],
      },
      decorations: { valueAxis: true },
    } as unknown as ChartConfig);
    const labels = (ok.nodes as { name?: string; text?: string }[])
      .filter((n) => n.name === "value-axis")
      .map((n) => n.text);
    expect(labels).toEqual(["0%", "25%", "50%", "75%", "100%"]);
  });
});
