import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { TextNode, WedgeNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Sunburst rings. */

describe("sunburst", () => {
  const cfg: ChartConfig = {
    kind: "sunburst",
    ...DEFAULT_SIZE,
    data: {
      categories: ["G1 | A", "G1 | B", "G2 | C"],
      series: [{ name: "V", values: [30, 10, 40] }],
    },
    decorations: { segmentLabels: false },
  };
  const s = buildChart(cfg);

  it("nests items on an outer ring inside group wedges on the inner ring", () => {
    const g0 = s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === "group-0")!;
    const item = s.nodes.find((n): n is WedgeNode => n.kind === "wedge" && n.name === "slice-0")!;
    expect(g0).toBeTruthy();
    expect(item).toBeTruthy();
    // Inner ring (groups) sits inside the outer ring (items).
    expect(g0.r).toBeLessThan(item.r);
    expect(item.innerR).toBeGreaterThanOrEqual(g0.r - 0.01);
    // Group spans are proportional: G1 (40) vs G2 (40) → equal; item A (30) > B (10).
    const a = s.nodes.find((n): n is WedgeNode => n.name === "slice-0")!;
    const b = s.nodes.find((n): n is WedgeNode => n.name === "slice-1")!;
    expect(a.endAngle - a.startAngle).toBeGreaterThan(b.endAngle - b.startAngle);
  });

  it("flat data (no groups) makes a single-ring doughnut", () => {
    const flat = buildChart({ ...cfg, data: { categories: ["A", "B"], series: [{ name: "V", values: [1, 1] }] } });
    expect(flat.nodes.some((n) => n.name?.startsWith("group-"))).toBe(false);
    const slices = flat.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    expect(slices).toHaveLength(2);
    expect(slices.every((w) => w.innerR > 0)).toBe(true);
  });

  it("mixes a loose (ungrouped) label into the grouped ring without throwing", () => {
    // One "|"-less label makes the chart grouped (some label has a "|") yet sends
    // that item through groupOf/labelOf's no-"|" branch — an unnamed "" group.
    const mixed = buildChart({
      ...cfg,
      data: { categories: ["G1 | A", "loose", "G2 | B"], series: [{ name: "V", values: [20, 15, 25] }] },
    });
    // Still a two-ring sunburst: group wedges on the inner ring, items on the outer.
    expect(mixed.nodes.some((n) => n.name?.startsWith("group-"))).toBe(true);
    const slices = mixed.nodes.filter((n): n is WedgeNode => n.kind === "wedge" && !!n.name?.startsWith("slice-"));
    expect(slices).toHaveLength(3);
    // Every wedge angle is finite (the loose item did not produce NaN spans).
    for (const w of mixed.nodes.filter((n): n is WedgeNode => n.kind === "wedge")) {
      expect(Number.isFinite(w.startAngle)).toBe(true);
      expect(Number.isFinite(w.endAngle)).toBe(true);
    }
  });
});

describe("outer labels stay on the canvas (regression)", () => {
  // The radius reserved a fixed fs*0.5 vertical / fs*4 horizontal margin, but
  // outer labels sit at r + fs*0.7 in an fs*1.4-tall box as wide as their text.
  // The SHIPPED sample overflowed: label-3 ran y 295 -> 309 on a 300pt canvas.
  const overflowing = (s: ReturnType<typeof buildChart>) =>
    (s.nodes as { name?: string; x?: number; y?: number; w?: number; h?: number }[])
      .filter((n) => n.name?.startsWith("label-") || n.name?.startsWith("group-label"))
      .filter(
        (n) =>
          (n.y ?? 0) < -0.5 ||
          (n.y ?? 0) + (n.h ?? 0) > s.height + 0.5 ||
          (n.x ?? 0) < -0.5 ||
          (n.x ?? 0) + (n.w ?? 0) > s.width + 0.5,
      )
      .map((n) => `${n.name} y ${n.y?.toFixed(1)}->${((n.y ?? 0) + (n.h ?? 0)).toFixed(1)}`);

  it("the shipped sample fits", () => {
    const s = buildChart(sampleConfig("sunburst"));
    expect(overflowing(s), `overflowing: ${overflowing(s).join(" | ")}`).toEqual([]);
  });

  it("long labels fit too", () => {
    const s = buildChart({
      kind: "sunburst",
      width: 480,
      height: 300,
      data: {
        categories: [
          "EMEA | Enterprise segment",
          "EMEA | SMB segment",
          "APAC | Enterprise segment",
          "APAC | Consumer segment",
        ],
        series: [{ name: "V", values: [40, 30, 20, 10] }],
      },
      decorations: { segmentLabels: true },
    } as unknown as ChartConfig);
    expect(overflowing(s), `overflowing: ${overflowing(s).join(" | ")}`).toEqual([]);
  });
});

/**
 * Adjacent OUTSIDE labels were the last overlapping text left at the default
 * font once every other shape had been fixed — one pair at 120x90 and two at
 * 300x60, grazing by 0.3 to 1.0pt.
 *
 * Each label was bounded by its own wedge's ARC, which says nothing about where
 * the NEIGHBOUR's midpoint falls: a wide wedge beside a narrow one earns a tall
 * label and still sits close to it. Every outside label is on one circle
 * (`r + fs*0.7`), so the gap between two of them is a fact the layout can read
 * once they are all placed — no `collide.ts` pass needed, which matters because
 * these are named `label-N`, the same name scatter and bubble use for POINT
 * labels, and moving those was measured as a worse trade twice.
 */
describe("sunburst — adjacent outside labels", () => {
  const outside = (cfg: ChartConfig) =>
    buildChart(cfg)
      .nodes.filter((n): n is TextNode => n.kind === "text" && /^label-\d+$/.test(n.name ?? "") && !!n.text.trim())
      // Only the ring's own labels: `align: "center"` is an INSIDE label, which
      // is bounded by its chord and takes no part in this.
      .filter((n) => n.align !== "center");

  const inkY = (t: TextNode) => {
    const base =
      t.valign === "top"
        ? t.y + t.fontSize
        : t.valign === "bottom"
          ? t.y + t.h - t.fontSize * 0.25
          : t.y + t.h / 2 + t.fontSize * 0.36;
    return { y0: base - t.fontSize * 0.8, y1: base + t.fontSize * 0.21 };
  };

  it.each([
    [120, 90],
    [300, 60],
    [200, 150],
    [480, 300],
  ])("no two of them overlap at %ix%i", (w, h) => {
    const ls = outside({ ...sampleConfig("sunburst"), width: w, height: h } as ChartConfig);
    expect(ls.length, "no outside labels were drawn — the check would be vacuous").toBeGreaterThan(1);
    const bad: string[] = [];
    for (let i = 0; i < ls.length; i++) {
      for (let j = i + 1; j < ls.length; j++) {
        // Side matters: a left label and a right label never meet however close
        // their y values are, so only same-side pairs are a collision.
        if (ls[i].align !== ls[j].align) continue;
        const a = inkY(ls[i]);
        const b = inkY(ls[j]);
        if (Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0) > 0.01) bad.push(`${ls[i].name} over ${ls[j].name}`);
      }
    }
    expect(bad, bad.join(" | ")).toEqual([]);
  });

  it("leaves a roomy chart completely alone", () => {
    // The fit must be a LAST RESORT. A shrink that fires where nothing needed
    // shrinking moves showcase slides — the radar's ticks did exactly that once.
    const ls = outside({ ...sampleConfig("sunburst"), width: 480, height: 300 } as ChartConfig);
    for (const l of ls) expect(l.fontSize, `${l.name} was shrunk on a 480x300 chart`).toBeCloseTo(8.5, 5);
  });

  it("drops only ONE of a pair it cannot separate, not both", () => {
    // Past the font floor shrinking cannot help, but dropping both loses a label
    // the survivor's room could have carried. At 300x60 the narrow wedge's label
    // goes and its neighbour stays.
    const ls = outside({ ...sampleConfig("sunburst"), width: 300, height: 60 } as ChartConfig);
    const names = ls.map((l) => l.name);
    expect(names).toContain("label-5");
    expect(names).not.toContain("label-6");
  });
});
