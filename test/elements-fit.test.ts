import { describe, expect, it } from "vitest";
import { buildKpiTile, buildProcessFlow, buildTableScene } from "../src/core/elements";
import { textWidth } from "../src/core/scene";
import type { ArrowheadNode, Scene, TextNode } from "../src/core/scene";

/**
 * Glyph placement and fit-to-width for the non-chart elements: the bug hunt
 * found in-cell glyphs orphaned from their (right-aligned) value, arrowheads
 * anchored by their tip on the text line, process-flow labels drawn over each
 * other and off-frame, and an empty table whose closing rule sat above y=0.
 */
const node = (s: Scene, name: string) => s.nodes.find((n) => n.name === name)!;

/** An arrowhead's (x, y) is its TIP; its body runs 1.8*size back along `angle`. */
const arrowInk = (a: ArrowheadNode) => {
  const rad = (a.angle * Math.PI) / 180;
  const back = { x: a.x - Math.cos(rad) * a.size * 1.8, y: a.y - Math.sin(rad) * a.size * 1.8 };
  return {
    cx: (a.x + back.x) / 2,
    cy: (a.y + back.y) / 2,
    left: Math.min(a.x, back.x) - Math.abs(Math.sin(rad)) * a.size * 0.7,
    right: Math.max(a.x, back.x) + Math.abs(Math.sin(rad)) * a.size * 0.7,
  };
};

/** Right-aligned cell text: the ink ends at the right edge of its box. */
const rightAlignedInkLeft = (t: TextNode) => t.x + t.w - textWidth(t.text, t.fontSize, t.bold);

describe("in-cell effect glyphs", () => {
  const table = buildTableScene(
    [
      ["Region", "Q1 revenue", "YoY"],
      ["North America", "1,204", "[up] +12%"],
      ["Europe", "986", "[down] -4%"],
    ],
    480,
  );

  it("sits next to its own right-aligned value, not at the cell's left edge", () => {
    const arrow = arrowInk(node(table, "cell-trend-1-2") as ArrowheadNode);
    const own = node(table, "cell-text-1-2") as TextNode;
    const neighbour = node(table, "cell-text-1-1") as TextNode;
    const ownInkLeft = rightAlignedInkLeft(own);
    const neighbourInkRight = neighbour.x + neighbour.w;
    expect(arrow.right).toBeLessThanOrEqual(ownInkLeft);
    // Adjacent to the value it annotates…
    expect(ownInkLeft - arrow.right).toBeLessThan(10);
    // …and further from the previous column's number than from that value.
    expect(ownInkLeft - arrow.right).toBeLessThan(arrow.left - neighbourInkRight);
  });

  it("straddles the text line, so an up row and a down row align", () => {
    const up = arrowInk(node(table, "cell-trend-1-2") as ArrowheadNode);
    const down = arrowInk(node(table, "cell-trend-2-2") as ArrowheadNode);
    const rowMid = (ri: number) => {
      const t = node(table, `cell-text-${ri}-2`) as TextNode;
      return t.y + t.h / 2;
    };
    expect(up.cy).toBeCloseTo(rowMid(1), 6);
    expect(down.cy).toBeCloseTo(rowMid(2), 6);
    expect(down.cy - up.cy).toBeCloseTo(rowMid(2) - rowMid(1), 6);
  });

  it("centres the KPI tile's delta arrow on the delta text line", () => {
    for (const opts of [
      { label: "Revenue", value: "€4.2M", delta: "+12% vs LY" },
      { value: "2.1%", delta: "-0.4pp", goodIsUp: false },
    ]) {
      const tile = buildKpiTile(opts);
      const arrow = arrowInk(node(tile, "kpi-arrow") as ArrowheadNode);
      const delta = node(tile, "kpi-delta") as TextNode;
      expect(arrow.cy).toBeCloseTo(delta.y + delta.h / 2, 6);
    }
  });
});

describe("process-flow labels fit their chevron", () => {
  const check = (steps: string[]) => {
    const scene = buildProcessFlow(steps, 0, 480, 40);
    const labels = scene.nodes.filter((n): n is TextNode => n.kind === "text");
    expect(labels).toHaveLength(steps.length);
    for (const l of labels) {
      const ink = textWidth(l.text, l.fontSize, l.bold);
      expect(ink, `${l.name} ink`).toBeLessThanOrEqual(l.w + 1e-9);
      // Centre-aligned inside its box, so the ink must also start on the canvas.
      expect(l.x + (l.w - ink) / 2, `${l.name} left`).toBeGreaterThanOrEqual(0);
    }
    return labels;
  };

  it("shrinks the font as steps crowd, and never draws a label off-frame", () => {
    check(["Requirements", "Design", "Development", "Testing", "Deployment", "Handover", "Benefits"]);
    check(Array(13).fill("Discover"));
    check(Array(20).fill("Discover"));
  });

  it("ellipsizes what no readable font size can fit", () => {
    const labels = check(Array(20).fill("Discovery workshop"));
    expect(labels[0].text.endsWith("…")).toBe(true);
  });

  it("leaves a comfortable flow at the full 11pt", () => {
    const labels = check(["Scope", "Build", "Launch"]);
    expect(labels.map((l) => [l.fontSize, l.text])).toEqual([
      [11, "Scope"],
      [11, "Build"],
      [11, "Launch"],
    ]);
  });
});

describe("empty table", () => {
  it("draws nothing rather than a closing rule above the scene origin", () => {
    const scene = buildTableScene([], 480);
    expect(scene.height).toBe(0);
    expect(scene.nodes).toEqual([]);
  });
});
