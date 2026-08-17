import { describe, expect, it } from "vitest";
import { buildKpiTile, buildProcessFlow, buildTableScene, clipToWidth } from "../src/core/elements";
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

/**
 * The type says `string`; the value came out of a file someone pasted.
 *
 * All five element builders are exported from `src/index.ts`, so the skill's
 * caller reaches them with hand-written JSON, and the pane builds its table
 * straight from datasheet cells. Every case below is something an author
 * plainly meant — a year in a table, a tile with no value yet, a blank line in
 * a pasted block — and every one of them used to throw a TypeError out of the
 * engine and take the whole element down.
 */
describe("elements survive the JSON they are actually handed", () => {
  it("renders a numeric table cell instead of throwing on it", () => {
    // `(2024).match is not a function`. A year in a table is not exotic input.
    const scene = buildTableScene([["Year", 2024 as unknown as string]], 480);
    const texts = scene.nodes.filter((n): n is TextNode => n.kind === "text").map((n) => n.text);
    expect(texts).toContain("2024");
  });

  it("treats a row that is not an array as an empty row, not a crash", () => {
    // A blank line in a pasted block arrives as null; one row written without
    // its inner array arrives as a bare string. Both threw before any of the
    // cell-level care could run.
    for (const rows of [[["a"], null], [["a"], "b"], null] as unknown as string[][][]) {
      expect(() => buildTableScene(rows, 480)).not.toThrow();
    }
  });

  it("builds a KPI tile that has no value yet", () => {
    // `buildKpiTile({})` — the tile before its number is filled in — threw
    // `Cannot read properties of undefined (reading 'length')` from textWidth.
    for (const opts of [{}, { value: null }, { value: undefined }] as never[]) {
      expect(() => buildKpiTile(opts)).not.toThrow();
    }
  });

  it("renders a numeric delta rather than throwing on it", () => {
    // `RegExp.test` coerces and `.replace` does not, so the direction was
    // inferred correctly and only the render step fell over.
    const scene = buildKpiTile({ value: "4.2", delta: 12 as unknown as string });
    expect((node(scene, "kpi-delta") as TextNode).text).toBe("12");
  });

  it("keeps a process flow whose step is missing", () => {
    for (const steps of [[null], [undefined], ["ok", null]] as unknown as string[][]) {
      expect(() => buildProcessFlow(steps)).not.toThrow();
    }
  });

  it("measures a number as its digits, so fit-to-width still shrinks", () => {
    // The quiet half. `(2024).length` is undefined, so textWidth returned NaN,
    // and every fit test here is a comparison that is FALSE against NaN — the
    // shrink loop simply stopped shrinking. Same text, same measurement.
    expect(textWidth(2024 as unknown as string, 10)).toBe(textWidth("2024", 10));
    expect(textWidth(null as unknown as string, 10)).toBe(0);
    expect(textWidth(undefined as unknown as string, 10)).toBe(0);
  });
});

/**
 * The two element builders whose text was never fitted to the box it is in.
 *
 * Every chart label in the engine is fitted to the mark it sits on, and neither
 * PowerPoint renderer wraps or clips a text box — so anything wider than its
 * shape is drawn straight over whatever sits beside the element on the slide.
 * These two are public API (`src/index.ts`, and the skill's own element
 * helpers), and both drew past their own frame:
 *
 * - `buildKpiTile` shrinks its big number, but stops at 11pt because a KPI
 *   number below that is not the thing a tile exists to show. Past the floor it
 *   simply overflowed — 47.8pt beyond a 160pt tile. Its label and its delta
 *   were never fitted at all.
 * - `buildTableScene` sizes columns in proportion to their content and then
 *   SCALES them to the table's width, while the text in them stayed at 10pt. A
 *   table whose content is wider than it is therefore ran its cells into each
 *   other and its last cell 685pt off the end of a 480pt table.
 */
describe("element text stays inside the element", () => {
  const ink = (t: TextNode) => {
    const w = textWidth(t.text, t.fontSize, t.bold);
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    return { x0: x, x1: x + w };
  };
  const worstOverflow = (s: Scene) =>
    Math.max(
      0,
      ...s.nodes.filter((n): n is TextNode => n.kind === "text").map((t) => Math.max(-ink(t).x0, ink(t).x1 - s.width)),
    );

  it("keeps a KPI tile's value, label and delta on the tile", () => {
    for (const opts of [
      { value: "a really long value string here", label: "and a long label to go over it too" },
      {
        value: "1,234,567,890.12",
        label: "Revenue, trailing twelve months, all regions",
        delta: "+123.4% YoY vs plan",
      },
      { value: "€1.2bn", delta: "▲ +12.3 percentage points against a very long plan name" },
    ]) {
      const tile = buildKpiTile(opts as never);
      expect(worstOverflow(tile), `${opts.value} overflowed its tile`).toBeLessThanOrEqual(0.5);
    }
  });

  it("keeps every table cell inside the table", () => {
    for (const [cells, width] of [
      [[["A sentence long enough that it cannot possibly fit in one column", "b"]], 480],
      [
        [
          ["Region", "Commentary"],
          ["North America", "Growth held up through the quarter despite the pricing change"],
          ["Europe", "Soft, with the promotional calendar pulled forward into March"],
        ],
        480,
      ],
      [[["x".repeat(200), "y"]], 200],
    ] as [string[][], number][]) {
      const table = buildTableScene(cells, width);
      expect(worstOverflow(table), "a cell ran off the table").toBeLessThanOrEqual(0.5);
    }
  });

  it("does not let two cells' text run into each other", () => {
    // Narrow enough that the columns are squeezed — the case where the text
    // stayed at 10pt while its column did not.
    const table = buildTableScene(
      [
        ["Region", "Commentary", "Value"],
        ["North America", "Growth held up through the quarter despite the pricing change", "1,204"],
      ],
      200,
    );
    const row = table.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("cell-text-1-"));
    expect(row.length).toBe(3);
    const boxes = row.map(ink).sort((a, b) => a.x0 - b.x0);
    for (let i = 1; i < boxes.length; i++)
      expect(boxes[i].x0, "two cells' text overlap").toBeGreaterThanOrEqual(boxes[i - 1].x1 - 0.5);
  });

  it("leaves an element that already fits exactly where it was", () => {
    // The negative control. Neither fix may touch content with room to spare.
    const tile = buildKpiTile({ value: "42", label: "Revenue", delta: "+3%" } as never);
    expect((node(tile, "kpi-value") as TextNode).text).toBe("42");
    expect((node(tile, "kpi-label") as TextNode).text).toBe("Revenue");
    const table = buildTableScene(
      [
        ["Region", "Q1 revenue", "YoY"],
        ["North America", "1,204", "+12%"],
      ],
      480,
    );
    expect(table.nodes.filter((n): n is TextNode => n.kind === "text").map((t) => t.text)).toEqual([
      "Region",
      "Q1 revenue",
      "YoY",
      "North America",
      "1,204",
      "+12%",
    ]);
    expect(table.height).toBe(42);
  });

  it("clips a numeric value instead of adding an ellipsis to it", () => {
    // `clipToWidth` walked with `t.length`, which is `undefined` on a number —
    // so the walk never ran and it returned its input with an ellipsis stuck on
    // the end, WIDER than the box it was asked to fit.
    const clipped = clipToWidth(123456789012345 as unknown as string, 10, 20);
    expect(textWidth(clipped, 10)).toBeLessThanOrEqual(20);
    expect(clipToWidth(null as unknown as string, 10, 20)).toBe("");
  });
});
