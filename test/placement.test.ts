import { describe, expect, it } from "vitest";
import { placeChart, SLIDE_HEIGHT_PT, SLIDE_WIDTH_PT, type Rect } from "../src/core/placement";

/**
 * Where a second chart goes when the first one is already there.
 *
 * The insert path used to answer this with a fixed 14pt cascade, which against
 * a 480x300 chart is not a cascade but a pile: the two overlapped by better
 * than 90% of their area, and the user saw one chart drawn on top of another.
 *
 * It then answered it downwards only, because sliding right needs the slide's
 * width and nothing could read it. `slideSize` reads it three ways now, so the
 * rule considers both directions and keeps the better one.
 */

const SIZE = { width: 480, height: 300 };
const ORIGIN = { left: 60, top: 90 };
/** The old cascade, which survives as the last resort. */
const CASCADE = { left: 74, top: 104 };

const WIDE = { width: SLIDE_WIDTH_PT, height: SLIDE_HEIGHT_PT };
/** A 4:3 deck — same height, 240pt narrower. */
const NARROW = { width: 720, height: SLIDE_HEIGHT_PT };

const rect = (left: number, top: number, width: number, height: number): Rect => ({ left, top, width, height });

/** Do these two boxes share any area at all? */
const overlaps = (a: Rect, b: Rect) =>
  a.left < b.left + b.width && b.left < a.left + a.width && a.top < b.top + b.height && b.top < a.top + a.height;

describe("placeChart", () => {
  it("uses the default origin on an empty slide", () => {
    expect(placeChart([], SIZE, ORIGIN, CASCADE)).toMatchObject({ ...ORIGIN, ...SIZE, shrunk: false, moved: "none" });
  });

  it("ignores shapes it does not actually collide with", () => {
    // A narrow shape off to the side is not a reason to move the chart at all.
    const aside = rect(700, 90, 100, 300);
    expect(placeChart([aside], SIZE, ORIGIN, CASCADE)).toMatchObject({ ...ORIGIN, shrunk: false, moved: "none" });
  });

  it("clears a SECOND shape it lands on after clearing the first", () => {
    // Clearing one shape can land you on the next, so the rule has to re-check
    // rather than move once. Two stacked bands, and a chart short enough that a
    // single move would have parked it inside the lower one.
    const a = rect(60, 90, 480, 100);
    const b = rect(60, 200, 480, 100);
    const size = { width: 480, height: 80 };
    const at = placeChart([a, b], size, ORIGIN, CASCADE, WIDE);
    expect(overlaps({ ...at, width: at.width, height: at.height }, a)).toBe(false);
    expect(overlaps({ ...at, width: at.width, height: at.height }, b)).toBe(false);
  });

  describe("choosing a direction", () => {
    it("puts the second chart BESIDE the first when both fit whole", () => {
      // Two charts on a slide are overwhelmingly meant to be read side by side,
      // so a tie goes to beside rather than below.
      const first = rect(60, 90, 300, 150);
      const size = { width: 300, height: 150 };
      const at = placeChart([first], size, ORIGIN, CASCADE, WIDE);
      expect(at).toMatchObject({ shrunk: false, moved: "beside" });
      expect(at.left).toBeGreaterThanOrEqual(first.left + first.width);
      expect(at.top).toBe(ORIGIN.top);
      expect(overlaps(at, first)).toBe(false);
    });

    it("prefers the direction that leaves the LARGER chart when both must shrink", () => {
      // The ordinary case the original bug report was about, and the one that
      // makes horizontal placement worth having. Two default 480x300 charts fit
      // neither way on a 16:9 slide, so both directions shrink — but the band
      // left beside the first is 390pt wide against the 120pt left below it.
      const first = rect(60, 90, 480, 300);
      const at = placeChart([first], SIZE, ORIGIN, CASCADE, WIDE);
      expect(at.moved).toBe("beside");
      expect(at.shrunk).toBe(true);
      expect(overlaps(at, first)).toBe(false);
      // Four times the area the downward answer would have given (192x120).
      expect(at.width * at.height).toBeGreaterThan(3 * (192 * 120));
      // Same shape, smaller — a squashed chart is worse than a small one.
      expect(at.width / at.height).toBeCloseTo(SIZE.width / SIZE.height, 1);
    });

    it("goes BELOW instead when the slide is too narrow to sit beside", () => {
      // The same two charts on a 4:3 deck. Beside leaves 150pt, which scales the
      // chart to 94pt tall — under the readability floor — so down it goes. This
      // is the case that made width worth reading rather than assuming: guessing
      // 16:9 here would have placed a chart 4pt tall or run it off the edge.
      const first = rect(60, 90, 480, 300);
      const at = placeChart([first], SIZE, ORIGIN, CASCADE, NARROW);
      expect(at.moved).toBe("below");
      expect(at.shrunk).toBe(true);
      expect(at.left).toBe(ORIGIN.left);
      expect(at.top).toBeGreaterThanOrEqual(first.top + first.height);
      expect(overlaps(at, first)).toBe(false);
    });

    it("reports which way it moved", () => {
      // The pane says what it did, and "it appeared beside the last one" is a
      // surprise worth one sentence.
      expect(placeChart([], SIZE, ORIGIN, CASCADE, WIDE).moved).toBe("none");
      expect(placeChart([rect(60, 90, 300, 150)], { width: 300, height: 150 }, ORIGIN, CASCADE, WIDE).moved).toBe(
        "beside",
      );
      expect(placeChart([rect(60, 90, 480, 300)], SIZE, ORIGIN, CASCADE, NARROW).moved).toBe("below");
    });
  });

  describe("staying on the slide", () => {
    it("never runs off the bottom", () => {
      for (const h of [100, 200, 300, 420]) {
        for (const occupied of [40, 120, 260, 380]) {
          const at = placeChart([rect(60, 90, 480, occupied)], { width: 480, height: h }, ORIGIN, CASCADE, WIDE);
          // Either it fits, or it is the cascade — and the cascade is the
          // caller's business, not a claim about fitting.
          if (at.moved !== "cascade") expect(at.top + at.height).toBeLessThanOrEqual(WIDE.height);
        }
      }
    });

    it("never runs off the right edge", () => {
      // The failure mode horizontal placement introduces, and the reason the
      // width has to be measured rather than assumed. A chart pushed right past
      // a wide shape has to be checked against the slide, not just the shape.
      for (const w of [150, 300, 480]) {
        for (const occupiedWidth of [100, 400, 700, 880]) {
          for (const slide of [WIDE, NARROW]) {
            const at = placeChart(
              [rect(60, 90, occupiedWidth, 200)],
              { width: w, height: 200 },
              ORIGIN,
              CASCADE,
              slide,
            );
            if (at.moved !== "cascade") expect(at.left + at.width).toBeLessThanOrEqual(slide.width);
          }
        }
      }
    });

    it("does not claim a fit for a tall chart squeezed in beside a narrow shape", () => {
      // Sliding right does not change the top, so a chart that clears the shape
      // horizontally can still hang off the bottom. Without that check this
      // reported a clean placement for a chart half of which was off-slide.
      const narrowShape = rect(60, 90, 80, 40);
      const tall = { width: 300, height: 500 };
      const at = placeChart([narrowShape], tall, ORIGIN, CASCADE, WIDE);
      if (at.moved !== "cascade") expect(at.top + at.height).toBeLessThanOrEqual(WIDE.height);
    });
  });

  it("falls back to the cascade only when NEITHER direction has room", () => {
    // A chart placed off the slide is worse than one the user can see and drag,
    // so the old cascade is the floor. It takes a shape that blocks both ways to
    // reach it now — a tall shape alone no longer does, because there is room
    // beside it.
    const blocking = rect(60, 90, 880, 460);
    const at = placeChart([blocking], SIZE, ORIGIN, CASCADE, WIDE);
    expect(at).toMatchObject({ ...CASCADE, ...SIZE, shrunk: false, moved: "cascade" });
  });

  it("terminates on shapes that cannot be cleared", () => {
    // A pile of overlapping rectangles is the shape of an accidental infinite
    // loop: each move can land on another. Bounded by the shape count, on both
    // axes.
    const many = Array.from({ length: 40 }, (_, i) => rect(60 + i, 90 + i, 480, 300));
    expect(() => placeChart(many, SIZE, ORIGIN, CASCADE, WIDE)).not.toThrow();
  });

  it("assumes 16:9 when the caller supplies no slide size", () => {
    // The default exists for callers that cannot measure. It matches the floor
    // `slideSize` itself falls back to, so the two agree about what "unknown"
    // means rather than each guessing separately.
    const first = rect(60, 90, 300, 150);
    const size = { width: 300, height: 150 };
    expect(placeChart([first], size, ORIGIN, CASCADE)).toEqual(placeChart([first], size, ORIGIN, CASCADE, WIDE));
  });
});
