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
 * host layer only has to supply rectangles. See `getSlideShapeBounds`.
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
 * PowerPoint slide is 7.5in tall, 4:3 (720x540) and 16:9 (960x540) alike, so a
 * host that answers nothing still gets the height right. Width is the dimension
 * that actually differs, and it is now passed in rather than assumed away.
 */
export const SLIDE_HEIGHT_PT = 540;

/** Points of clear space left between a chart and whatever it sits under. */
const GAP = 12;

/** Bottom margin — a chart flush to the slide edge reads as a mistake. */
const MARGIN = 18;

/**
 * The shortest a chart may be shrunk to before shrinking stops being a favour.
 *
 * Two default 480x300 charts do not both fit on a 540pt slide at full size —
 * they need 612pt of height and there are about 500 — so "put it below" alone
 * cannot answer the ordinary case of a second chart. Scaling the new one into
 * whatever band is left does, and beats the alternative the user actually
 * reported: one chart drawn on top of another.
 */
const MIN_HEIGHT = 110;

export interface Placement {
  left: number;
  top: number;
  width: number;
  height: number;
  /** The chart was scaled down to fit the space left on the slide. */
  shrunk: boolean;
}

const overlaps = (a: Rect, b: Rect): boolean =>
  a.left < b.left + b.width && b.left < a.left + a.width && a.top < b.top + b.height && b.top < a.top + a.height;

/**
 * The top-left for a chart of `size`, given what is already on the slide.
 *
 * Downward only. Sliding right needs the slide's width, which — now that
 * `slideSize` can actually read it — is knowable, but placing horizontally is a
 * layout decision this function has never made and is deliberately not being
 * given here along with the number. What changes first is the thing that was
 * silently wrong: the generated-deck fast path declaring 16:9 on a 4:3 deck.
 *
 * Falls back to `fallback` — the caller's cascade — when the chart will not fit
 * below what is already there. That case genuinely has nowhere good to go, and
 * a visible overlap the user can drag beats a chart parked off the slide where
 * they cannot find it at all.
 */
export function placeChart(
  occupied: Rect[],
  size: { width: number; height: number },
  origin: { left: number; top: number },
  fallback: { left: number; top: number },
): Placement {
  let top = origin.top;
  // Re-check after every move: dropping below one shape can land on the next.
  // Bounded by the shape count, since each pass either settles or clears at
  // least one more rectangle.
  for (let pass = 0; pass <= occupied.length; pass++) {
    const here: Rect = { left: origin.left, top, ...size };
    const hits = occupied.filter((r) => overlaps(here, r));
    if (!hits.length) break;
    const lowest = Math.max(...hits.map((r) => r.top + r.height));
    if (lowest + GAP <= top) break; // not making progress — stop rather than spin
    top = lowest + GAP;
  }
  // Checked once, at the end, and never skipped: an early return from inside
  // the loop would hand back the first non-overlapping position without asking
  // whether it is still ON the slide, which for a tall chart under a tall shape
  // is a chart pushed off the bottom edge where the user cannot reach it.
  const room = SLIDE_HEIGHT_PT - MARGIN - top;
  if (size.height <= room) return { left: origin.left, top, ...size, shrunk: false };
  // Scale into what is left, keeping the aspect ratio the user configured.
  if (room >= MIN_HEIGHT) {
    const scale = room / size.height;
    return {
      left: origin.left,
      top,
      width: Math.round(size.width * scale),
      height: Math.round(room),
      shrunk: true,
    };
  }
  // Not even a small chart fits. Overlapping where the user can see and drag it
  // beats parking it off the slide, so this is where the old cascade earns its
  // keep — as the last resort it always should have been, not the only rule.
  return { ...fallback, ...size, shrunk: false };
}
