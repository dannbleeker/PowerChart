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
