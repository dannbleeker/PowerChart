import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

/**
 * Outside pie labels against their NEIGHBOURS, not just the frame edge.
 *
 * 46 of this project's 75 known text overlaps were one cause: a 24-slice pie has
 * more labels than it has ring. The only fit rule an outside label had was
 * horizontal clipping to the frame edge, which cannot see the label two points
 * above it, so adjacent labels were drawn over each other.
 *
 * De-collision is not the remedy and `collide.ts` says why: it nudges upward
 * only, and pie labels are deliberately absent from `MOVABLE` because a nudge
 * past a neighbour leaves a label beside the wrong wedge with its leader line
 * pointing at it. Shrink-then-drop moves nothing, so nothing can end up
 * mislabelled — and it is the order this layout already takes for the whole
 * ring, for inside labels, and that the funnel and butterfly take before it.
 */
const pie = (n: number, width = 480, height = 300): ChartConfig => ({
  kind: "pie",
  width,
  height,
  data: {
    categories: Array.from({ length: n }, (_, i) => `Category ${i + 1}`),
    series: [{ name: "S", values: Array.from({ length: n }, (_, i) => 10 + i) }],
  },
});

const labelsOf = (cfg: ChartConfig): TextNode[] =>
  buildChart(cfg).nodes.filter((n): n is TextNode => n.kind === "text" && /^label-/.test(n.name ?? "") && !!n.text);

/** Painted extent, the same approximation `tightBox` uses. */
const box = (n: TextNode) => ({
  x: n.x,
  y: n.y,
  w: Math.min(n.w, n.text.length * n.fontSize * 0.55),
  h: n.fontSize * 1.25,
});

function overlappingPairs(labels: TextNode[]): number {
  let n = 0;
  for (let i = 0; i < labels.length; i++) {
    for (let j = i + 1; j < labels.length; j++) {
      const a = box(labels[i]);
      const b = box(labels[j]);
      if (a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h) n++;
    }
  }
  return n;
}

describe("a pie with more labels than ring", () => {
  it("draws no label over another at 24 categories", () => {
    // Measured before this rule: 24 labels drawn, 7 overlapping pairs.
    const labels = labelsOf(pie(24));
    expect(overlappingPairs(labels), "outside labels are colliding again").toBe(0);
  });

  it("drops the ones that cannot fit rather than shrinking past legibility", () => {
    // The mechanism has to actually engage — a rule that kept all 24 and simply
    // made them tiny would pass the overlap check and fail the reader. Measured:
    // 21 of 24 survive at this size, and the three lost are the narrowest
    // slices, whose labels are worth least.
    const labels = labelsOf(pie(24));
    expect(labels.length, "nothing was dropped, so nothing was crowded out").toBeLessThan(24);
    expect(labels.length, "too much was dropped to still be a labelled chart").toBeGreaterThan(12);
    // Nothing below the 5pt floor this file already uses for the whole ring.
    for (const l of labels) expect(l.fontSize, `${l.text} is below the legibility floor`).toBeGreaterThanOrEqual(5);
  });

  it("leaves an uncrowded pie alone", () => {
    // THE REGRESSION THAT WOULD MATTER. Five slices have room, and a rule that
    // dropped or shrank any of them would be paying for the 24-slice case out of
    // the everyday one.
    const five = labelsOf(pie(5));
    expect(five.length, "a five-slice pie lost a label").toBe(5);
    expect(overlappingPairs(five)).toBe(0);
  });

  it("keeps every wedge, whatever happens to the labels", () => {
    // Dropping a LABEL loses a name; dropping a slice would lose data. The chart
    // still shows all 24 shares — only the crowded names go.
    const wedges = buildChart(pie(24)).nodes.filter((n) => n.kind === "wedge");
    expect(wedges.length).toBe(24);
  });
});
