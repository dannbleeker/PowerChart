import { describe, expect, it } from "vitest";
import { resolveLabelCollisions } from "../src/core/collide";
import type { SceneNode, TextNode } from "../src/core/scene";

/** Label collision resolution — which labels may move, and how far. */

const label = (over: Partial<TextNode>): TextNode => ({
  kind: "text",
  x: 0,
  y: 0,
  w: 40,
  h: 12,
  text: "a",
  fontSize: 10,
  color: "#000",
  align: "center",
  valign: "middle",
  ...over,
});

describe("label collision resolution", () => {
  it("ignores unnamed and immovable nodes", () => {
    const overlapped: SceneNode[] = [label({ text: "a" }), label({ text: "b", name: "segment-label-0-0" })];
    const before = JSON.stringify(overlapped);
    resolveLabelCollisions(overlapped);
    expect(JSON.stringify(overlapped)).toBe(before);
  });

  it("nudges a movable total up off a fixed label when there is room", () => {
    const nodes: SceneNode[] = [
      label({ y: 40, text: "fixed", name: "segment-label-0-0" }),
      label({ y: 40, text: "42", name: "total-0" }),
    ];
    resolveLabelCollisions(nodes);
    const total = nodes[1] as TextNode;
    expect(total.y).toBeLessThan(40); // moved up, clear of the fixed label
    expect(total.y).toBeGreaterThanOrEqual(0); // but still on the canvas
  });

  /**
   * How far, not just which way.
   *
   * The three assertions above pin a DIRECTION — moved up, still on the canvas —
   * and a mutation run showed what that costs: changing the nudge from
   * `fontSize * 0.55` to `fontSize / 0.55`, a 3.3x longer jump, was invisible to
   * the whole suite. So were the loop bound and the rank sentinel. `collide.ts`
   * scored 53% against 47 surviving mutants, the worst file in `src/core`, and
   * every one of those survivors is a label landing somewhere nobody asserted.
   *
   * The numbers here are derived, not recorded: a 10pt label whose tight box is
   * 12 tall, starting at y=40 against a fixed box spanning 40..52, must clear to
   * y+12 <= 40. At 5.5 per nudge that is three nudges — 40 - 16.5 = 23.5. Change
   * the step and this number moves.
   */
  it("nudges by exactly the step it claims, and no further than it must", () => {
    const nodes: SceneNode[] = [
      label({ y: 40, text: "fixed", name: "segment-label-0-0" }),
      label({ y: 40, text: "42", name: "total-0" }),
    ];
    resolveLabelCollisions(nodes);
    expect((nodes[1] as TextNode).y).toBe(23.5);
  });

  /**
   * The nudge gives up after ten tries, and the cap is observable.
   *
   * Deleting it entirely (`tries < 10` to `true`) survived the whole suite,
   * including twenty-four hostile-input tests that run this exact line — their
   * bar is "does not throw, stays finite", which a runaway loop satisfies right
   * up until it doesn't.
   *
   * A 200pt fixed label spanning 0..250 can never be cleared by a label starting
   * inside it: every nudge moves further in. Capped, it stops after ten steps at
   * 240 - 55 = 185, still overlapping and honestly so. Uncapped, it would climb
   * until the off-canvas guard fired and restore to 240 — so the two are a
   * different number rather than a different feeling.
   */
  it("stops nudging after a bounded number of tries", () => {
    const nodes: SceneNode[] = [
      label({ y: 0, h: 250, fontSize: 200, text: "big", name: "segment-label-0-0" }),
      label({ y: 240, text: "42", name: "total-0" }),
    ];
    resolveLabelCollisions(nodes);
    expect((nodes[1] as TextNode).y, "the ten-nudge cap is not doing anything").toBe(185);
  });

  /**
   * An unnamed label is not movable, tested where the difference shows.
   *
   * The first case in this file asserts exactly this and cannot see it: both its
   * labels sit at y=0, so a label wrongly treated as movable is pushed off the
   * canvas, restored by the off-canvas guard, and ends where it started. Flipping
   * `movableRank`'s sentinel from -1 to +1 — which makes every unnamed label
   * movable — passed. Given room above, it does not.
   */
  it("leaves an unnamed label alone even when it has room to move", () => {
    const nodes: SceneNode[] = [
      label({ y: 100, text: "fixed", name: "segment-label-0-0" }),
      label({ y: 100, text: "unnamed" }),
    ];
    resolveLabelCollisions(nodes);
    expect((nodes[1] as TextNode).y, "an unnamed label was treated as movable").toBe(100);
  });

  /**
   * A solid band of IMMOVABLE labels between two y positions.
   *
   * One tall label will not do: `tightBox` clamps a box to the ink it actually
   * carries (`fontSize * 1.25`), so a label declared 66 points tall obstructs
   * 12.5 of them and the nudge simply steps over it. That is how the first
   * version of these tests passed against the unfixed file — nothing was ever
   * in the label's way.
   */
  const wall = (from: number, to: number, name: string): SceneNode[] =>
    Array.from({ length: Math.ceil((to - from) / 6) + 1 }, (_, i) =>
      label({ y: from + i * 6, text: name, fontSize: 10, name: "segment-label-0-0" }),
    );

  it("flips a point label BELOW its mark when everything above it is taken", () => {
    // The nudge only goes up, and a point label can be boxed in from above. It
    // used to climb the whole ten-nudge budget and come to rest inside whatever
    // was up there — on a 480x300 combo at a 26pt font, the chart's own TITLE.
    // The wall spans more than that budget (10 x 0.55 x fontSize = 55 points),
    // so no legal number of upward nudges can clear it.
    const point = label({ y: 60, text: "45", fontSize: 10, name: "combo-label-0-3" });
    resolveLabelCollisions([...wall(0, 60, "above"), point], 300);
    expect(point.y, "the point label did not move below its obstacle").toBeGreaterThan(60);
  });

  it("leaves a boxed-in point label where the layout put it rather than moving it somewhere worse", () => {
    // Blocked above AND below: the flip must put it back, not leave it wherever
    // the failed search stopped. A label that moved and still collides is worse
    // than one that never moved — it collides somewhere its author did not pick.
    const point = label({ y: 60, text: "45", fontSize: 10, name: "combo-label-0-3" });
    resolveLabelCollisions([...wall(0, 60, "above"), ...wall(66, 200, "below"), point], 300);
    expect(point.y).toBe(60);
  });

  it("does not flip a point label off the bottom of the canvas", () => {
    // The mirror of the off-the-top guard: the first clear spot below is past
    // the foot of the chart, so the flip refuses it and restores. An
    // overlapping label reads; one drawn off the canvas is lost.
    const point = label({ y: 60, text: "45", fontSize: 10, name: "combo-label-0-3" });
    resolveLabelCollisions([...wall(0, 60, "above"), point], 85);
    expect(point.y).toBe(60);
    expect(point.y + point.h).toBeLessThanOrEqual(85);
  });

  it("flips the CAGR caption below its arrow rather than leave it on the title", () => {
    // The other way of getting this caption off the title — flooring it at the
    // title's bottom — was tried and reverted: it turned five `title x
    // cagr-label` overlaps into eight against the column totals. A flip cannot
    // do that, because it only takes a spot already clear of everything
    // settled, and totals rank ahead of it.
    const cagr = label({ y: 60, text: "+15.2% p.a.", fontSize: 10, name: "cagr-label" });
    resolveLabelCollisions([...wall(0, 60, "above"), cagr], 300);
    expect(cagr.y).toBeGreaterThan(60);
  });

  it("does not flip the CAGR caption onto a total, which is what the floor did", () => {
    // Totals rank ahead of the caption, so they settle first and every position
    // below is already taken. The flip must find nothing and put the caption
    // back — not buy its way off the title by landing on a total, which is
    // exactly what the floor this replaces did.
    //
    // Spaced 14 apart against an ink height of 12.5 so they clear each OTHER:
    // a first version stacked them, the pass nudged them apart on their own
    // account, and the test read that as the caption shoving them.
    const cagr = label({ y: 60, text: "+15.2% p.a.", fontSize: 10, name: "cagr-label" });
    const totals = [80, 94, 108, 122, 136].map((y, i) => label({ y, text: "93", fontSize: 10, name: `total-${i}` }));
    const before = totals.map((t) => t.y);
    resolveLabelCollisions([...wall(0, 60, "above"), ...totals, cagr], 300);
    expect(cagr.y, "the caption moved onto the totals row").toBe(60);
    expect(
      totals.map((t) => t.y),
      "a total was pushed to make room for the caption",
    ).toEqual(before);
  });

  it("does not flip a series label below, because that would reorder it", () => {
    // Only labels anchored to a single point may go down. Moving a series label
    // down crosses the label beneath it, and the two then name each other's
    // lines — every label on the canvas, each one wrong.
    const series = label({ y: 60, text: "Brand A", fontSize: 10, name: "series-label-0" });
    resolveLabelCollisions([...wall(0, 60, "above"), series], 300);
    expect(series.y).toBeLessThanOrEqual(60);
  });

  it("does not nudge a movable total off the top of the canvas", () => {
    // A total pinned against the top with no room (e.g. sharing the totals row
    // with the fixed grand-total label) must NOT escape the canvas: an
    // overlapping label reads, an off-canvas one is lost. It gives up in place.
    const nodes: SceneNode[] = [
      label({ text: "fixed", name: "segment-label-0-0" }),
      label({ text: "42", name: "total-0" }),
    ];
    resolveLabelCollisions(nodes);
    expect((nodes[1] as TextNode).y).toBeGreaterThanOrEqual(0);
  });
});
