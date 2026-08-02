/**
 * Where the next chart goes on a slide that already has things on it.
 *
 * The insert path used to cascade: 60+n, 90+n, stepping n by 14pt and wrapping
 * at 84. Against a default chart of 480x300 that is not a cascade, it is a
 * pile — two charts inserted back to back overlapped by better than 90% of
 * their area, and the fourth landed back under the first. The step was chosen
 * to look like PowerPoint's own "paste offsets a little", but PowerPoint is
 * offsetting a copy the user is about to drag, and this is a chart the user
 * asked to have placed.
 *
 * Pure on purpose: no Office import here, so the rule is unit-testable and the
 * host layer only has to supply rectangles and the slide's own size. See
 * `getSlideShapeBounds` and `slideSize`.
 */

export interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * Default slide height in points, for callers that cannot supply the real one.
 *
 * This used to say PowerPointApi exposes no slide dimensions at any requirement
 * set. That was wrong: `PageSetup.slideWidth`/`slideHeight` land it directly in
 * 1.10, an exported slide's `<p:sldSz>` carries it at 1.8, and the Common API's
 * `getFileAsync` carries it on any host at all — see `slideSize` in
 * `src/render/powerpoint.ts`, which works down exactly that ladder.
 *
 * The number survives as a DEFAULT because it is a good one: every standard
 * PowerPoint slide is 7.5in tall, 4:3 (720x540) and 16:9 (960x540) alike.
 */
export const SLIDE_HEIGHT_PT = 540;

/**
 * Default slide width in points — 16:9, the same floor `slideSize` falls to.
 *
 * Unlike the height this one is a real guess: a 4:3 deck is 720. It is only
 * ever reached when the caller did not pass a measured size, and being wrong
 * here costs a chart placed beside another when it should have gone below.
 */
export const SLIDE_WIDTH_PT = 960;

/** Points of clear space left between a chart and whatever it sits under. */
const GAP = 12;

/** Bottom and right margin — a chart flush to the slide edge reads as a mistake. */
const MARGIN = 18;

/**
 * The shortest a chart may be shrunk to before shrinking stops being a favour.
 *
 * Two default 480x300 charts do not both fit on a 540pt slide at full size —
 * they need 612pt of height and there are about 500 — so "put it below" alone
 * cannot answer the ordinary case of a second chart. Scaling the new one into
 * whatever band is left does, and beats the alternative the user actually
 * reported: one chart drawn on top of another.
 *
 * Applied to the resulting HEIGHT in both directions. A chart is unreadable
 * below a certain height whichever axis the squeeze came from, and the aspect
 * ratio is preserved, so one floor covers both cases.
 */
const MIN_HEIGHT = 110;

export interface Placement {
  left: number;
  top: number;
  width: number;
  height: number;
  /** The chart was scaled down to fit the space left on the slide. */
  shrunk: boolean;
  /**
   * Which way the chart had to move to find room. `"none"` means it landed on
   * the requested origin. Surfaced so the pane can say what it did — a chart
   * that quietly appears beside the last one is a surprise worth one sentence.
   */
  moved: "none" | "beside" | "below" | "cascade";
}

const overlaps = (a: Rect, b: Rect): boolean =>
  a.left < b.left + b.width && b.left < a.left + a.width && a.top < b.top + b.height && b.top < a.top + a.height;

/** Slide the chart along one axis until it clears everything, or give up. */
function slidePast(
  occupied: Rect[],
  size: { width: number; height: number },
  origin: { left: number; top: number },
  axis: "x" | "y",
): number {
  let at = axis === "x" ? origin.left : origin.top;
  // Re-check after every move: clearing one shape can land you on the next.
  // Bounded by the shape count, since each pass either settles or clears at
  // least one more rectangle.
  for (let pass = 0; pass <= occupied.length; pass++) {
    const here: Rect = axis === "x" ? { left: at, top: origin.top, ...size } : { left: origin.left, top: at, ...size };
    const hits = occupied.filter((r) => overlaps(here, r));
    if (!hits.length) break;
    const past = Math.max(...hits.map((r) => (axis === "x" ? r.left + r.width : r.top + r.height)));
    if (past + GAP <= at) break; // not making progress — stop rather than spin
    at = past + GAP;
  }
  return at;
}

/**
 * Fit `size` into `room` along one axis, keeping the aspect ratio.
 *
 * Returns the scaled size, or null when the result would be too short to read.
 */
function fitInto(
  size: { width: number; height: number },
  room: number,
  axis: "x" | "y",
): { width: number; height: number; shrunk: boolean } | null {
  const along = axis === "x" ? size.width : size.height;
  if (along <= room) return { ...size, shrunk: false };
  const scale = room / along;
  const height = axis === "x" ? size.height * scale : room;
  if (height < MIN_HEIGHT) return null;
  return {
    width: axis === "x" ? Math.round(room) : Math.round(size.width * scale),
    height: Math.round(height),
    shrunk: true,
  };
}

/** The best placement to the RIGHT of what is already there, or null. */
function placeBeside(
  occupied: Rect[],
  size: { width: number; height: number },
  origin: { left: number; top: number },
  slide: { width: number; height: number },
): Placement | null {
  const left = slidePast(occupied, size, origin, "x");
  const fitted = fitInto(size, slide.width - MARGIN - left, "x");
  if (!fitted) return null;
  // Sliding right does not change the top, so the chart still has to fit
  // DOWNWARDS from the origin it was given. Without this a tall chart beside a
  // narrow shape reported a fit while hanging off the bottom of the slide.
  if (origin.top + fitted.height > slide.height - MARGIN) return null;
  return { left, top: origin.top, ...fitted, moved: left === origin.left ? "none" : "beside" };
}

/** The best placement BELOW what is already there, or null. */
function placeBelow(
  occupied: Rect[],
  size: { width: number; height: number },
  origin: { left: number; top: number },
  slide: { width: number; height: number },
): Placement | null {
  const top = slidePast(occupied, size, origin, "y");
  const fitted = fitInto(size, slide.height - MARGIN - top, "y");
  if (!fitted) return null;
  // Same guard on the other axis: a chart pushed down is no narrower, so a
  // slide too narrow for it at full width is still too narrow here.
  if (origin.left + fitted.width > slide.width - MARGIN) return null;
  return { left: origin.left, top, ...fitted, moved: top === origin.top ? "none" : "below" };
}

const area = (p: Placement) => p.width * p.height;

/**
 * The top-left for a chart of `size`, given what is already on the slide.
 *
 * Tries BESIDE and BELOW, and keeps the better of the two.
 *
 * It used to try only below, because sliding right needs the slide's width and
 * the comment here claimed no requirement set would tell us — which was wrong,
 * and `slideSize` now reads it three different ways. Being able to go right
 * matters more than it sounds: two default 480x300 charts fit neither way on a
 * 16:9 slide, so both are shrunk, but the band left BESIDE the first is 390pt
 * wide against the 120pt left BELOW it. Same slide, same two charts, and the
 * horizontal answer is a chart 390x244 where the vertical one is 192x120 —
 * four times the area, for the arrangement people actually want.
 *
 * "Better" is therefore area, with one thumb on the scale: a placement that
 * needs no shrinking always beats one that does, however large the shrunk one
 * would be, because a chart at the size the user asked for is what they asked
 * for. Between two that both fit whole, beside wins — two charts on a slide
 * are overwhelmingly meant to be read side by side.
 *
 * Falls back to `fallback` — the caller's cascade — when neither direction has
 * room. That case genuinely has nowhere good to go, and a visible overlap the
 * user can drag beats a chart parked off the slide where they cannot find it.
 */
export function placeChart(
  occupied: Rect[],
  size: { width: number; height: number },
  origin: { left: number; top: number },
  fallback: { left: number; top: number },
  slide: { width: number; height: number } = { width: SLIDE_WIDTH_PT, height: SLIDE_HEIGHT_PT },
): Placement {
  const beside = placeBeside(occupied, size, origin, slide);
  const below = placeBelow(occupied, size, origin, slide);
  const best = (() => {
    if (!beside) return below;
    if (!below) return beside;
    // Area alone decides it, and that already encodes "whole beats scaled":
    // both candidates start from the SAME requested size, so an unshrunk one
    // necessarily has more area than any shrunk one. An explicit `shrunk`
    // tiebreak ahead of this would be a branch no input could reach.
    //
    // A tie means both fit whole, and it goes to beside.
    return area(below) > area(beside) ? below : beside;
  })();
  if (best) return best;
  return { ...fallback, ...size, shrunk: false, moved: "cascade" };
}
