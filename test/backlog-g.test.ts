import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { LineNode, RectNode, TextNode } from "../src/core/scene";

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

describe("butterfly value ticks", () => {
  const base: ChartConfig = {
    kind: "butterfly",
    ...DEFAULT_SIZE,
    data: {
      categories: ["A", "B"],
      series: [
        { name: "Left", values: [10, 20] },
        { name: "Right", values: [15, 25] },
      ],
    },
  };

  it("adds no axis chrome by default", () => {
    const plain = buildChart(base);
    expect(plain.nodes.some((n) => n.name?.startsWith("tick-"))).toBe(false);
    expect(plain.nodes.some((n) => n.name?.startsWith("gridline-"))).toBe(false);
  });

  it("valueAxis + gridlines draw mirrored ticks and gridlines on both flanks", () => {
    const axed = buildChart({ ...base, decorations: { valueAxis: true, gridlines: true } });
    expect(axed.nodes.some((n) => n.name?.startsWith("gridline-"))).toBe(true);
    const ticks = axed.nodes.filter((n): n is TextNode => !!n.name?.startsWith("tick-"));
    const lefts = ticks.filter((n) => n.name!.endsWith("-l"));
    const rights = ticks.filter((n) => n.name!.endsWith("-r"));
    expect(lefts.length).toBeGreaterThan(1);
    expect(lefts.length).toBe(rights.length); // mirrored
    expect(axed.nodes.some((n) => n.name === "tick-0-l")).toBe(true);
  });
});
