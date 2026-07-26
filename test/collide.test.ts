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
