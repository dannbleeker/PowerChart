import { describe, it, expect } from "vitest";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

/**
 * Radar spoke names against each other, in the direction the chord cannot see.
 *
 * The last of the "24 categories" family, after the pie's outside labels went
 * the same way. A radar already sized its perimeter names by the CHORD between
 * two spokes, which is the room a name has ALONG the ring — and near the top and
 * the bottom of the web the ring runs horizontally, two names sit almost side by
 * side, and what limits them there is their width, which the chord never looked
 * at. A 24-spoke web drew `category-9` over `category-10` for that reason.
 *
 * Shrink, then drop, and nothing moves — the same order the pie takes. On a web
 * it matters more than anywhere else: the chart IS the mapping from name to
 * spoke, so a name nudged onto its neighbour's axis does not look untidy, it
 * lies. A name that leaves is a name the reader looks up; a name that lies is
 * worse than no name at all.
 *
 * Measured across the frame sweep's eight sizes at both sweep fonts, before and
 * after:
 *
 *     before   288 names drawn   96 overlapping pairs
 *     after    236 names drawn    0 overlapping pairs
 *
 * A five-spoke web is untouched — the regression that would have mattered,
 * paying for the crowded case out of the everyday one. Its count across the same
 * sweep is identical before and after, and it is NOT forty (five names times
 * eight frames times two fonts): a small web drops its whole perimeter ring
 * long before any of this, which is `ringFits` and deliberate.
 */
const radar = (n: number, width = 480, height = 300, fontSize?: number): ChartConfig =>
  ({
    kind: "radar",
    width,
    height,
    ...(fontSize ? { style: { fontSize } } : {}),
    data: {
      categories: Array.from({ length: n }, (_, i) => `Category ${i + 1}`),
      series: [{ name: "S", values: Array.from({ length: n }, (_, i) => 1 + (i % 5)) }],
    },
  }) as ChartConfig;

const namesOf = (cfg: ChartConfig): TextNode[] =>
  buildChart(cfg).nodes.filter(
    (n): n is TextNode => n.kind === "text" && /^category-\d+$/.test(n.name ?? "") && !!n.text.trim(),
  );

/**
 * The ink a name paints, positioned by its own alignment.
 *
 * The NODE box is wider than the text for a centred or right-aligned label, and
 * measuring the box would report overlaps the reader never sees. This is what
 * the frame gate measures, down to the area threshold: two names count as
 * colliding when they share more than a square point, which is the rule the
 * layout now fits itself by.
 */
const ink = (t: TextNode) => {
  const w = Math.min(t.w, t.text.length * t.fontSize * 0.55);
  const x0 = t.align === "center" ? t.x + (t.w - w) / 2 : t.align === "right" ? t.x + t.w - w : t.x;
  return { x0, x1: x0 + w, y0: t.y, y1: t.y + t.h };
};

function overlappingPairs(labels: TextNode[]): number {
  let n = 0;
  for (let i = 0; i < labels.length; i++) {
    for (let j = i + 1; j < labels.length; j++) {
      const a = ink(labels[i]);
      const b = ink(labels[j]);
      const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0);
      const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
      if (w > 0 && h > 0 && w * h > 1) n++;
    }
  }
  return n;
}

/** The frame sweep's sizes, so this is checked where the overlaps were found. */
const FRAMES: [number, number][] = [
  [80, 60],
  [120, 90],
  [160, 120],
  [200, 150],
  [300, 60],
  [60, 300],
  [480, 300],
  [960, 540],
];

describe("a radar with more spokes than perimeter", () => {
  it("draws no name over another, at any frame the sweep uses", () => {
    // Frame by frame rather than in total: a single number hides WHICH size
    // regressed, and 200x150 at 10pt is the one that survived two earlier
    // versions of this rule.
    for (const [w, h] of FRAMES) {
      for (const fontSize of [10, 18]) {
        const labels = namesOf(radar(24, w, h, fontSize));
        expect(overlappingPairs(labels), `${w}x${h} fs=${fontSize}: spoke names are colliding`).toBe(0);
      }
    }
  });

  it("drops what will not fit rather than shrinking past legibility", () => {
    // The mechanism has to engage. A rule that kept all 24 and made them tiny
    // would pass the overlap check and fail the reader.
    const labels = namesOf(radar(24, 200, 150, 10));
    expect(labels.length, "nothing was dropped, so nothing was crowded out").toBeLessThan(24);
    for (const l of labels) expect(l.fontSize, `${l.text} is below the legibility floor`).toBeGreaterThanOrEqual(5);
  });

  it("leaves an uncrowded web alone", () => {
    // THE REGRESSION THAT WOULD MATTER. Five spokes have room, and a rule that
    // dropped or shrank one of them would be paying for the 24-spoke case out of
    // the everyday one.
    for (const [w, h] of [
      [480, 300],
      [960, 540],
    ]) {
      for (const fontSize of [10, 18]) {
        const five = namesOf(radar(5, w, h, fontSize));
        expect(five.length, `${w}x${h} fs=${fontSize}: a five-spoke web lost a name`).toBe(5);
      }
    }
  });

  it("draws exactly as many five-spoke names as it did before the rule", () => {
    // Written asserting all five at every sweep size, and that was wrong: a
    // radar at 80x60 drops its whole perimeter ring, deliberately and long
    // before this — `ringFits`, because a label past the frame is worse than no
    // label — and at 200x150 in 18pt the same thing happens with these long
    // names. The total across the sweep is the honest guard: 35 of a possible
    // 80, measured with the crowding pass removed. If this number moves, the
    // rule has started charging the uncrowded case.
    let drawn = 0;
    for (const [w, h] of FRAMES) for (const fontSize of [10, 18]) drawn += namesOf(radar(5, w, h, fontSize)).length;
    expect(drawn, "the crowding rule reached a five-spoke web").toBe(35);
  });

  it("keeps every spoke, whatever happens to the names", () => {
    // Dropping a NAME loses a label; dropping a spoke would lose data. The web
    // still has all 24 axes — only the crowded names go.
    const spokes = buildChart(radar(24)).nodes.filter((n) => n.kind === "line" && /^spoke-/.test(n.name ?? ""));
    expect(spokes.length).toBe(24);
  });

  it("still builds a web with hundreds of spokes", () => {
    // The rule walks each name against its two neighbours, so it is linear and
    // needs no spoke cap — an all-pairs version did, and was dropped when a
    // 3,024-combination search showed it changed no answer. This is the guard
    // that the linear one has not quietly grown quadratic again: the schema
    // allows 4096 categories, and a pane waits on this synchronously.
    const started = performance.now();
    const huge = buildChart(radar(400));
    expect(performance.now() - started, "building a 400-spoke web got slow").toBeLessThan(2000);
    expect(huge.nodes.some((n) => n.kind === "line" && /^spoke-/.test(n.name ?? ""))).toBe(true);
  });
});
