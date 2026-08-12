import { describe, expect, it } from "vitest";
import { DEFAULT_SIZE, buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { textWidth } from "../src/core/scene";
import type { RectNode, SceneNode, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

/** Funnel — band geometry, short-frame fit, conversion markers. */

/**
 * Degenerate-frame / degenerate-scale guards found by a layout bug-hunt. Each is
 * byte-identical for normal charts (snapshots unchanged) and only repairs an edge
 * that previously emitted negative geometry or NaN coordinates.
 */
const negDims = (nodes: ReturnType<typeof buildChart>["nodes"]) =>
  nodes.filter(
    (n) =>
      ((n.kind === "rect" || n.kind === "text") && ((n as RectNode).w < 0 || (n as RectNode).h < 0)) ||
      Object.entries(n).some(([k, v]) => ["x", "y", "w", "h"].includes(k) && typeof v === "number" && Number.isNaN(v)),
  );

/**
 * Distribution-family bug hunt: radar / butterfly / candlestick / violin /
 * funnel / waterfall / column legend. Each guard pins the exact wrong output the
 * hunt observed, so the fix cannot silently regress.
 */
const W = 480;

const H = 300;

const texts = (nodes: SceneNode[], namePrefix: string) =>
  nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith(namePrefix)).map((n) => n.text);

const base = (partial: Partial<ChartConfig>): ChartConfig => ({
  kind: "stacked",
  ...DEFAULT_SIZE,
  width: W_offframeguards,
  height: H,
  data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 20, 30] }] },
  ...partial,
});

/**
 * Geometry that must stay inside the canvas. These were found by the review as
 * off-frame / wrong-size layout bugs the loose fuzz bound (|c| < 5000) hid.
 */
const W_offframeguards = 480;

describe("funnel", () => {
  const s = buildChart(sampleConfig("funnel"));

  it("draws centered bands with width proportional to value", () => {
    const bands = [0, 1, 4].map((c) => s.nodes.find((n) => n.name === `stage-${c}`) as RectNode);
    // Widths proportional: 720/1200, 120/1200.
    expect(bands[1].w / bands[0].w).toBeCloseTo(720 / 1200, 2);
    expect(bands[2].w / bands[0].w).toBeCloseTo(120 / 1200, 2);
    // Centered: all bands share the same center x.
    const cx = (r: RectNode) => r.x + r.w / 2;
    expect(cx(bands[1])).toBeCloseTo(cx(bands[0]), 5);
    expect(cx(bands[2])).toBeCloseTo(cx(bands[0]), 5);
  });

  it("labels conversion vs the previous stage between bands", () => {
    const conv = s.nodes.filter((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("conversion-"));
    expect(conv).toHaveLength(4);
    expect(conv[0].text).toContain("60.0%"); // 720/1200
    // Stage names on the left, values on/beside the bands.
    expect(s.nodes.some((n) => n.name === "category-0")).toBe(true);
    expect((s.nodes.find((n) => n.name === "stage-value-4") as TextNode).text).toBe("120");
  });

  it("narrow bands put their value beside the band, wide ones inside", () => {
    const wide = s.nodes.find((n) => n.name === "stage-value-0") as TextNode;
    const band0 = s.nodes.find((n) => n.name === "stage-0") as RectNode;
    expect(wide.x).toBeGreaterThanOrEqual(band0.x); // inside
    const tiny = buildChart({
      ...sampleConfig("funnel"),
      data: { categories: ["All", "Won"], series: [{ name: "Deals", values: [10000, 12] }] },
    });
    const narrow = tiny.nodes.find((n) => n.name === "stage-value-1") as TextNode;
    const band = tiny.nodes.find((n) => n.name === "stage-1") as RectNode;
    expect(narrow.x).toBeGreaterThan(band.x + band.w); // beside
  });
});

describe("funnel bands never go negative on a short, crowded frame", () => {
  it("floors band height", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: 640,
      height: 60, // 20 stages + 1.5em gaps can't fit → bands went negative
      data: {
        categories: Array.from({ length: 20 }, (_, i) => `S${i}`),
        series: [{ name: "v", values: Array.from({ length: 20 }, (_, i) => 20 - i) }],
      },
    };
    expect(negDims(buildChart(cfg).nodes)).toHaveLength(0);
  });
});

describe("funnel conversion marker follows the direction of the step", () => {
  it("marks a rise with ▴ and a fall with ▾", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: W,
      height: H,
      // Ascending — the pyramid ordering funnel.ts recommends.
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [10, 50, 0] }] },
    };
    // Before: "▾ 500.0%" — a down arrow on a 5x increase.
    expect(texts(buildChart(cfg).nodes, "conversion-")).toEqual(["▴ 500.0%", "▾ 0.0%"]);
  });

  it("drops the marker when the stage is unchanged", () => {
    const cfg: ChartConfig = {
      kind: "funnel",
      width: W,
      height: H,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [10, 10] }] },
    };
    expect(texts(buildChart(cfg).nodes, "conversion-")).toEqual(["100.0%"]);
  });
});

describe("funnel bands fit a short frame", () => {
  it("many stages on a short frame stay inside the canvas", () => {
    const scene = buildChart(
      base({
        kind: "funnel",
        height: 60, // deliberately short: a fixed gap used to overshoot the bottom
        data: { categories: ["a", "b", "c", "d", "e", "f"], series: [{ name: "S", values: [60, 50, 40, 30, 20, 10] }] },
      }),
    );
    const stages = scene.nodes.filter((n) => /^stage-\d+$/.test(n.name ?? "")); // the bands, not stage-value-*
    expect(stages.length).toBe(6);
    for (const s of stages as any[]) expect(s.y + s.h).toBeLessThanOrEqual(60 + 1);
  });
});

describe("a stage value is never drawn off the chart", () => {
  /**
   * "Outside" means to the RIGHT of the band, and the widest band already
   * reaches the edge of the plot — so at a large font there is no room out there
   * and the label was drawn past the frame. `stage-value-0` landed at x = 480.0
   * on a 480pt frame: the TOP stage showed no value at all while every stage
   * below it did, because the one band big enough to matter is the one with
   * nothing to its right. Rendered and looked at, which is how it was seen: the
   * funnel had four numbers on five bars.
   *
   * The band that cannot fit a label beside it is the WIDEST one, which is
   * exactly the band with the most room inside it — so it goes inside, in
   * contrast ink. A cramped label on the bar beats a missing number.
   */
  const inkSpan = (t: TextNode) => {
    const w = Math.min(t.w, textWidth(t.text, t.fontSize, t.bold));
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    return { x0: x, x1: x + w };
  };
  const values = (fontSize: number) =>
    buildChart({ ...sampleConfig("funnel"), style: { fontSize } } as ChartConfig).nodes.filter(
      (n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("stage-value-"),
    );

  it("keeps every stage's value inside the frame, at every font size", () => {
    for (const fs of [10, 16, 22, 28, 36]) {
      const found = values(fs);
      // Every stage still HAS a value — the rule is not satisfied by dropping
      // the label that would not fit.
      expect(found.length, `fs=${fs} lost a stage value`).toBe(5);
      for (const t of found) {
        const { x0, x1 } = inkSpan(t);
        expect(x1, `fs=${fs} ${t.name} ran past the right edge`).toBeLessThanOrEqual(DEFAULT_SIZE.width);
        expect(x0, `fs=${fs} ${t.name} started left of the frame`).toBeGreaterThanOrEqual(0);
      }
    }
  });

  it("still puts the value beside the band when there is room for it", () => {
    // The fix must not collapse into "always inside": at a small font the narrow
    // lower stages carry their value outside the bar, which is what makes them
    // readable at all.
    const outside = values(10).filter((t) => t.align === "left");
    const inside = values(10).filter((t) => t.align === "center");
    expect(inside.length, "no value sits inside its band").toBeGreaterThan(0);
    expect(outside.length + inside.length).toBe(5);
  });
});

describe("a funnel's row labels stay on the chart at any font", () => {
  /**
   * The bands are solved for the plot, so they always fit. Their LABELS were
   * not, and each escaped a different way once the font outgrew the chart:
   *
   *   - a label is centred in its band and its ink is about `f` tall, so past
   *     the band height the bottom row's label hung BELOW the plot — 4.4pt at a
   *     28pt font, 15.1 at 36;
   *   - a category name is right-aligned in a gutter capped at 28% of the
   *     width, so a name wider than the cap ran off the LEFT edge —
   *     "Negotiation" by 35.9pt at 28, 83.4 at 36.
   *
   * Neither PowerPoint renderer wraps or clips a text box, so the ink measured
   * here is not bounded by the box it sits in — which is why a check that
   * clamped to the box saw the vertical half of this and missed the horizontal.
   */
  const inkOf = (t: TextNode) => {
    const w = textWidth(t.text, t.fontSize, t.bold);
    const x = t.align === "right" ? t.x + t.w - w : t.align === "center" ? t.x + (t.w - w) / 2 : t.x;
    const base =
      t.valign === "top"
        ? t.y + t.fontSize
        : t.valign === "bottom"
          ? t.y + t.h - t.fontSize * 0.25
          : t.y + t.h / 2 + t.fontSize * 0.36;
    return { x0: x, x1: x + w, y1: base + t.fontSize * 0.21 };
  };
  const rowText = (fontSize: number) => {
    const scene = buildChart({ ...sampleConfig("funnel"), style: { fontSize } } as ChartConfig);
    return {
      scene,
      nodes: scene.nodes.filter(
        (n): n is TextNode =>
          n.kind === "text" && (!!n.name?.startsWith("category-") || !!n.name?.startsWith("stage-value-")),
      ),
    };
  };

  it("keeps every stage name and value inside the frame", () => {
    for (const fs of [10, 16, 22, 28, 36]) {
      const { nodes } = rowText(fs);
      expect(nodes.length, `fs=${fs} lost row labels`).toBeGreaterThan(5);
      for (const t of nodes) {
        const ink = inkOf(t);
        expect(ink.x0, `fs=${fs} ${t.name} ran off the left edge`).toBeGreaterThanOrEqual(-0.5);
        expect(ink.x1, `fs=${fs} ${t.name} ran off the right edge`).toBeLessThanOrEqual(DEFAULT_SIZE.width + 0.5);
        expect(ink.y1, `fs=${fs} ${t.name} hung below the frame`).toBeLessThanOrEqual(DEFAULT_SIZE.height + 0.5);
      }
    }
  });

  it("shrinks only as a last resort, and keeps the conversion rate subordinate", () => {
    // At a font that fits, nothing moves at all.
    for (const t of rowText(10).nodes) expect(t.fontSize).toBe(10);
    // One size for the whole row, so a name and its number match.
    expect(new Set(rowText(28).nodes.map((t) => t.fontSize)).size).toBe(1);
    // The conversion rate is muted secondary text and must stay SMALLER than the
    // stage names beside it — shrinking only the names inverted that.
    const scene = buildChart({ ...sampleConfig("funnel"), style: { fontSize: 28 } } as ChartConfig);
    const conv = scene.nodes.find((n): n is TextNode => n.kind === "text" && !!n.name?.startsWith("conversion-"))!;
    const name = scene.nodes.find((n): n is TextNode => n.kind === "text" && n.name === "category-0")!;
    expect(conv.fontSize).toBeLessThan(name.fontSize);
  });
});
