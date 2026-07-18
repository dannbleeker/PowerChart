import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";

/**
 * "The good chart" batch: connector lines, per-cell highlight colors,
 * callouts, background bands, exploding pie slices, footnote + 100% note.
 */

const base: ChartConfig = {
  kind: "stacked",
  ...DEFAULT_SIZE,
  data: {
    categories: ["A", "B", "C"],
    series: [
      { name: "S1", values: [10, 20, 30] },
      { name: "S2", values: [5, 6, 7] },
    ],
  },
};

const byPrefix = (cfg: ChartConfig, prefix: string) => buildChart(cfg).nodes.filter((n) => n.name?.startsWith(prefix));

describe("connector lines", () => {
  it("joins each segment boundary between adjacent stacked columns", () => {
    const conns = byPrefix({ ...base, decorations: { connectors: true, segmentLabels: true } }, "connector-");
    // 2 series → 2 positive boundaries × 2 gaps = 4 lines.
    expect(conns).toHaveLength(4);
    for (const c of conns) expect(c.kind).toBe("line");
  });

  it("connects negative-side boundaries separately", () => {
    const cfg: ChartConfig = {
      ...base,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S", values: [10, 12] },
          { name: "T", values: [-4, -6] },
        ],
      },
      decorations: { connectors: true, segmentLabels: true },
    };
    const names = byPrefix(cfg, "connector-").map((n) => n.name);
    // The trailing index is the series index on both sides — S is 0, T is 1 —
    // so the negative boundary is "-1n". (It was the push-order of the negative
    // segments, which drifted out of step with the series whenever one was zero.)
    expect(names).toContain("connector-0-0");
    expect(names).toContain("connector-0-1n");
  });

  it("works in horizontal orientation and stays off by default", () => {
    expect(
      byPrefix({ ...base, horizontal: true, decorations: { connectors: true, segmentLabels: true } }, "connector-"),
    ).toHaveLength(4);
    expect(byPrefix(base, "connector-")).toHaveLength(0);
  });

  it("is skipped for clustered-stacked (multiple stack groups)", () => {
    const cfg: ChartConfig = {
      ...base,
      data: {
        categories: ["A", "B"],
        series: [
          { name: "S1", values: [10, 12], stack: 0 },
          { name: "S2", values: [5, 6], stack: 1 },
        ],
      },
      decorations: { connectors: true, segmentLabels: true },
    };
    expect(byPrefix(cfg, "connector-")).toHaveLength(0);
  });
});

describe("per-cell highlight colors", () => {
  it("recolors a single stacked segment", () => {
    const cfg: ChartConfig = {
      ...base,
      data: { ...base.data, series: [{ name: "S1", values: [10, 20, 30], colors: [null, "#e34948", null] }] },
    };
    const segs = byPrefix(cfg, "seg-0-") as RectNode[];
    expect(segs[1].fill).toBe("#e34948");
    expect(segs[0].fill).not.toBe("#e34948");
  });

  it("recolors a pie slice and a line marker (larger, bold label)", () => {
    const pie: ChartConfig = {
      ...base,
      kind: "pie",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1], colors: [null, "#e34948"] }] },
    };
    const slices = byPrefix(pie, "slice-");
    expect(slices[1].kind === "wedge" && slices[1].fill).toBe("#e34948");

    const line: ChartConfig = {
      ...base,
      kind: "line",
      data: {
        categories: ["A", "B", "C"],
        series: [{ name: "S", values: [1, 5, 3], colors: [null, "#e34948", null] }],
      },
    };
    const markers = byPrefix(line, "marker-0-") as RectNode[];
    expect(markers[1].fill).toBe("#e34948");
    expect(markers[1].w).toBeGreaterThan(markers[0].w);
  });
});

describe("callouts", () => {
  it("draws a tail, box, and centered text anchored above the column", () => {
    const cfg: ChartConfig = {
      ...base,
      decorations: { segmentLabels: true, callouts: [{ text: "One-off", category: 1 }] },
    };
    const s = buildChart(cfg);
    const box = s.nodes.find((n) => n.name === "callout-box-0") as RectNode;
    const tail = s.nodes.find((n) => n.name === "callout-tail-0");
    const text = s.nodes.find((n) => n.name === "callout-text-0") as TextNode;
    expect(box).toBeDefined();
    expect(tail?.kind).toBe("line");
    expect(text.text).toBe("One-off");
    // Bubble hovers above the anchor (smaller y = higher).
    const seg = s.nodes.find((n) => n.name === "seg-0-1") as RectNode;
    expect(box.y + box.h).toBeLessThanOrEqual(seg.y);
  });

  it("anchors at a series level and accepts nudges", () => {
    const cfg: ChartConfig = {
      ...base,
      decorations: { segmentLabels: true, callouts: [{ text: "x", category: 0, series: 0, dx: 20, dy: 5 }] },
    };
    const plain = buildChart({
      ...base,
      decorations: { segmentLabels: true, callouts: [{ text: "x", category: 0, series: 0 }] },
    });
    const nudged = buildChart(cfg);
    const b0 = plain.nodes.find((n) => n.name === "callout-box-0") as RectNode;
    const b1 = nudged.nodes.find((n) => n.name === "callout-box-0") as RectNode;
    expect(b1.x - b0.x).toBeCloseTo(20);
    expect(b1.y - b0.y).toBeCloseTo(5);
  });
});

describe("background bands", () => {
  it("value-range band spans the plot width behind the columns", () => {
    const cfg: ChartConfig = {
      ...base,
      decorations: { segmentLabels: true, bands: [{ axis: "y", from: 10, to: 20, label: "Target" }] },
    };
    const s = buildChart(cfg);
    const band = s.nodes.find((n) => n.name === "band-0") as RectNode;
    expect(band).toBeDefined();
    // Bands render before segments (behind them).
    expect(s.nodes.indexOf(band)).toBeLessThan(s.nodes.findIndex((n) => n.name === "seg-0-0"));
    expect(s.nodes.some((n) => n.name === "band-label-0")).toBe(true);
  });

  it("category-range band covers the given category slots", () => {
    const cfg: ChartConfig = { ...base, decorations: { segmentLabels: true, bands: [{ axis: "x", from: 1, to: 2 }] } };
    const s = buildChart(cfg);
    const band = s.nodes.find((n) => n.name === "band-0") as RectNode;
    const seg1 = s.nodes.find((n) => n.name === "seg-0-1") as RectNode;
    const seg0 = s.nodes.find((n) => n.name === "seg-0-0") as RectNode;
    expect(band.x).toBeLessThanOrEqual(seg1.x);
    expect(band.x).toBeGreaterThan(seg0.x);
  });

  it("draws value-unit bands on scatter plots", () => {
    const cfg: ChartConfig = {
      ...base,
      kind: "scatter",
      data: {
        categories: ["P1", "P2"],
        series: [
          { name: "X", values: [10, 60] },
          { name: "Y", values: [10, 60] },
        ],
      },
      decorations: { bands: [{ axis: "x", from: 30, to: 50, label: "Focus" }] },
    };
    const s = buildChart(cfg);
    expect(s.nodes.some((n) => n.name === "band-0")).toBe(true);
    expect(s.nodes.some((n) => n.name === "band-label-0")).toBe(true);
  });
});

describe("exploding pie slice", () => {
  it("offsets the exploded wedge radially from the center", () => {
    const data = { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] };
    const plain = buildChart({ ...base, kind: "pie", data });
    const burst = buildChart({ ...base, kind: "pie", data, pie: { explode: [1] } });
    const w0 = plain.nodes.find((n) => n.name === "slice-1");
    const w1 = burst.nodes.find((n) => n.name === "slice-1");
    const keep = burst.nodes.find((n) => n.name === "slice-0");
    const k0 = plain.nodes.find((n) => n.name === "slice-0");
    expect(w0?.kind === "wedge" && w1?.kind === "wedge" && (w0.cx !== w1.cx || w0.cy !== w1.cy)).toBe(true);
    expect(keep?.kind === "wedge" && k0?.kind === "wedge" && keep.cx === k0.cx && keep.cy === k0.cy).toBe(true);
  });
});

describe("footnote and 100% note", () => {
  it("renders the source line bottom-left in muted small text", () => {
    const s = buildChart({ ...base, footnote: "Source: Statistics Denmark, 2024" });
    const fn = s.nodes.find((n) => n.name === "footnote") as TextNode;
    expect(fn.text).toBe("Source: Statistics Denmark, 2024");
    expect(fn.y).toBeGreaterThan(base.height * 0.9);
  });

  it("reserves space so the plot does not collide with the footnote", () => {
    const plain = buildChart(base);
    const withFn = buildChart({ ...base, footnote: "src" });
    const bottom = (sc: typeof plain) =>
      Math.max(...sc.nodes.filter((n) => n.name?.startsWith("seg-")).map((n) => (n as RectNode).y + (n as RectNode).h));
    expect(bottom(withFn)).toBeLessThan(bottom(plain));
  });

  it("renders '100% = N' for pies and uniform 100% charts, skips mixed denominators", () => {
    const pie = buildChart({
      ...base,
      kind: "pie",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [30, 70] }] },
      decorations: { hundredPercentNote: true },
    });
    expect((pie.nodes.find((n) => n.name === "footnote") as TextNode).text).toContain("100% = 100");

    const uniform = buildChart({
      ...base,
      kind: "stacked100",
      data: { ...base.data, hundredPercent: [50, 50, 50] },
      decorations: { hundredPercentNote: true, segmentLabels: true },
    });
    expect((uniform.nodes.find((n) => n.name === "footnote") as TextNode).text).toContain("100% = 50");

    const mixed = buildChart({
      ...base,
      kind: "stacked100",
      decorations: { hundredPercentNote: true, segmentLabels: true },
    });
    // Column sums differ (15/26/37) → no honest single note → nothing rendered.
    expect(mixed.nodes.some((n) => n.name === "footnote")).toBe(false);
  });

  it("joins the note and the source into one line", () => {
    const s = buildChart({
      ...base,
      kind: "pie",
      footnote: "Source: X",
      data: { categories: ["A"], series: [{ name: "S", values: [10] }] },
      decorations: { hundredPercentNote: true },
    });
    const fn = s.nodes.find((n) => n.name === "footnote") as TextNode;
    expect(fn.text).toContain("100% = 10");
    expect(fn.text).toContain("Source: X");
  });
});
