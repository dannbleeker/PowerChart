import { describe, expect, it } from "vitest";
import { placeChart, SLIDE_HEIGHT_PT, type Rect } from "../src/core/placement";

/**
 * Where a second chart goes when the first one is already there.
 *
 * The insert path used to answer this with a fixed 14pt cascade, which against
 * a 480x300 chart is not a cascade but a pile: the two overlapped by better
 * than 90% of their area, and the user saw one chart drawn on top of another.
 */

const SIZE = { width: 480, height: 300 };
const ORIGIN = { left: 60, top: 90 };
/** The old cascade, which survives as the last resort. */
const CASCADE = { left: 74, top: 104 };

const rect = (left: number, top: number, width: number, height: number): Rect => ({ left, top, width, height });

/** Do these two boxes share any area at all? */
const overlaps = (a: Rect, b: Rect) =>
  a.left < b.left + b.width && b.left < a.left + a.width && a.top < b.top + b.height && b.top < a.top + a.height;

describe("placeChart", () => {
  it("uses the default origin on an empty slide", () => {
    expect(placeChart([], SIZE, ORIGIN, CASCADE)).toMatchObject({ ...ORIGIN, ...SIZE, shrunk: false });
  });

  it("drops the chart clear of one already on the slide, instead of onto it", () => {
    const first = rect(60, 90, 480, 300);
    const at = placeChart([first], { width: 480, height: 120 }, ORIGIN, CASCADE);
    expect(overlaps({ ...at, width: 480, height: 120 }, first)).toBe(false);
    expect(at.top).toBeGreaterThanOrEqual(first.top + first.height);
  });

  it("clears a SECOND shape it lands on after clearing the first", () => {
    // Dropping below one shape can land you on the next, so the rule has to
    // re-check rather than move once. Two stacked bands, and a chart short
    // enough that a single move would have parked it inside the lower one.
    const a = rect(60, 90, 480, 100);
    const b = rect(60, 200, 480, 100);
    const size = { width: 480, height: 80 };
    const at = placeChart([a, b], size, ORIGIN, CASCADE);
    expect(overlaps({ ...at, ...size }, a)).toBe(false);
    expect(overlaps({ ...at, ...size }, b)).toBe(false);
  });

  it("ignores shapes it does not actually collide with", () => {
    // A narrow shape off to the side is not a reason to push the chart down.
    const aside = rect(700, 90, 100, 300);
    expect(placeChart([aside], SIZE, ORIGIN, CASCADE)).toMatchObject({ ...ORIGIN, shrunk: false });
  });

  /**
   * The ordinary case the bug report was about: two default charts, one slide.
   * They need 612pt of height between them and a slide is 540, so "put it
   * below" alone cannot place the second one — it has to come down in size.
   */
  it("shrinks a second default-sized chart into the band that is left", () => {
    const first = rect(60, 90, 480, 300);
    const at = placeChart([first], SIZE, ORIGIN, CASCADE);
    expect(at.shrunk).toBe(true);
    expect(overlaps(at, first)).toBe(false);
    expect(at.top + at.height).toBeLessThanOrEqual(SLIDE_HEIGHT_PT);
    // Same shape, smaller — a squashed chart would be worse than an overlapping one.
    expect(at.width / at.height).toBeCloseTo(SIZE.width / SIZE.height, 1);
  });

  it("falls back to the cascade when not even a small chart fits", () => {
    // A chart placed past the bottom edge is worse than one the user can see
    // and drag, so the old cascade is the floor — never an off-slide position.
    const tall = rect(60, 90, 480, 460);
    const at = placeChart([tall], SIZE, ORIGIN, CASCADE);
    expect(at).toMatchObject({ ...CASCADE, ...SIZE, shrunk: false });
  });

  it("never returns a chart that runs off the bottom of the slide", () => {
    for (const h of [100, 200, 300, 420]) {
      for (const occupied of [40, 120, 260, 380]) {
        const at = placeChart([rect(60, 90, 480, occupied)], { width: 480, height: h }, ORIGIN, CASCADE);
        // Either it fits, or it is the cascade — and the cascade is the
        // caller's business, not a claim about fitting.
        if (at.top !== CASCADE.top) expect(at.top + at.height).toBeLessThanOrEqual(SLIDE_HEIGHT_PT);
      }
    }
  });

  it("terminates on shapes that cannot be cleared", () => {
    // A pile of overlapping rectangles is the shape of an accidental infinite
    // loop: each move can land on another. Bounded by the shape count.
    const many = Array.from({ length: 40 }, (_, i) => rect(60, 90 + i, 480, 300));
    expect(() => placeChart(many, SIZE, ORIGIN, CASCADE)).not.toThrow();
  });
});
