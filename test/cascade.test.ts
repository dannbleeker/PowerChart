import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig } from "../src/core/types";
import type { RectNode, TextNode } from "../src/core/scene";

/** Cascade / decomposition chart: each stage is a subset of the previous. */

const cfg: ChartConfig = {
  kind: "cascade",
  ...DEFAULT_SIZE,
  data: {
    categories: [
      "Total contacts | | Contacts",
      "Answered | Dropped contacts | Contacts",
      "With a case | Without a case | Incidents",
      "Solved | Not solved | Incidents",
    ],
    series: [{ name: "Volume", values: [4986, 4616, 3405, 2685] }],
  },
};

const rect = (s: ReturnType<typeof buildChart>, name: string) => s.nodes.find((n) => n.name === name) as RectNode;
const text = (s: ReturnType<typeof buildChart>, name: string) => s.nodes.find((n) => n.name === name) as TextNode;

describe("cascade chart", () => {
  const s = buildChart(cfg);

  it("draws top-aligned stage bars with heights proportional to volume", () => {
    const bars = [0, 1, 2, 3].map((c) => rect(s, `stage-${c}`));
    for (const b of bars) expect(b.y).toBeCloseTo(bars[0].y); // top-aligned
    expect(bars[1].h / bars[0].h).toBeCloseTo(4616 / 4986, 2);
    expect(bars[3].h / bars[0].h).toBeCloseTo(2685 / 4986, 2);
    // Distinct per-stage colors.
    expect(new Set(bars.map((b) => b.fill)).size).toBe(4);
  });

  it("labels each stage with its value and % of the previous stage", () => {
    expect(text(s, "stage-label-1-0").text).toBe("Answered");
    expect(text(s, "stage-label-1-1").text).toBe("4,616");
    expect(text(s, "stage-label-1-2").text).toBe("(92.6%)");
    // First stage is the 100% base — no percentage line.
    expect(s.nodes.some((n) => n.name === "stage-label-0-2")).toBe(false);
  });

  it("kept bar + drop box exactly equal the previous column — never longer", () => {
    // Invariant: column c's total extent (bar + remainder box) bottoms out
    // at PRECISELY the previous bar's bottom edge, for every split.
    for (const c of [1, 2, 3]) {
      const prev = rect(s, `stage-${c - 1}`);
      const drop = rect(s, `drop-${c}`);
      expect(drop.y + drop.h).toBeCloseTo(prev.y + prev.h, 5);
    }
    // Even when the remainder is tiny and the box needs label height, it
    // grows upward over the bar — not downward past the total.
    const tiny = buildChart({
      ...cfg,
      data: { categories: ["A", "B"], series: [{ name: "V", values: [1000, 995] }] },
    });
    const prev = tiny.nodes.find((n) => n.name === "stage-0") as RectNode;
    const drop = tiny.nodes.find((n) => n.name === "drop-1") as RectNode;
    expect(drop.y + drop.h).toBeCloseTo(prev.y + prev.h, 5);
    expect(drop.h).toBeGreaterThan(2); // still tall enough to exist visibly
  });

  it("hangs a labeled remainder box at each split", () => {
    const drop1 = rect(s, "drop-1");
    const bar1 = rect(s, "stage-1");
    expect(drop1.y).toBeGreaterThan(bar1.y); // below the bar's top, at the split
    expect(drop1.h / rect(s, "stage-0").h).toBeCloseTo(370 / 4986, 1);
    // Long captions wrap: caption line + numbers line.
    expect(text(s, "drop-label-1").text).toBe("Dropped contacts");
    expect(text(s, "drop-value-1").text).toBe("370 (7.4%)");
    expect(text(s, "drop-label-2").text).toBe("Without a case");
    expect(text(s, "drop-value-2").text).toContain("1,211");
    // No remainder before the first stage.
    expect(s.nodes.some((n) => n.name === "drop-0")).toBe(false);
  });

  it("spans group header bands over consecutive same-group stages", () => {
    const g0 = rect(s, "group-0");
    const g2 = rect(s, "group-2");
    expect(text(s, "group-label-0").text).toBe("Contacts");
    expect(text(s, "group-label-2").text).toBe("Incidents");
    // "Contacts" spans stages 0–1: wider than one bar, ends before group 2 starts.
    expect(g0.w).toBeGreaterThan(rect(s, "stage-0").w);
    expect(g0.x + g0.w).toBeLessThanOrEqual(g2.x + 1);
  });

  it("defaults the drop caption, skips growth stages, and survives plain categories", () => {
    const plain = buildChart({
      ...cfg,
      data: { categories: ["A", "B", "C"], series: [{ name: "V", values: [100, 60, 80] }] },
    });
    expect((plain.nodes.find((n) => n.name === "drop-label-1") as TextNode).text).toBe("Other: 40 (40.0%)");
    // B→C grows, so there is nothing dropped to draw.
    expect(plain.nodes.some((n) => n.name === "drop-2")).toBe(false);
    // No group parts → no header band.
    expect(plain.nodes.some((n) => n.name?.startsWith("group-"))).toBe(false);
  });

  it("renders the sample and respects the footnote reservation", () => {
    const sample = buildChart(sampleConfig("cascade"));
    expect(sample.nodes.some((n) => n.name === "footnote")).toBe(true);
    const bar = sample.nodes.find((n) => n.name === "stage-0") as RectNode;
    const fn = sample.nodes.find((n) => n.name === "footnote") as TextNode;
    expect(bar.y + bar.h).toBeLessThan(fn.y);
  });
});
