import { describe, expect, it } from "vitest";
import { thinOutline } from "../src/core/geometry";

/**
 * `tol` IS A GUARANTEE, and the first implementation did not give one.
 *
 * On the Office.js sink a polygon has no freeform fill, so the renderer draws
 * one line per EDGE and an outline's point count is a cost the host is charged.
 * The violin was 246 of its 259 office shapes on three KDE-sampled bodies, most
 * of the points in tails where consecutive samples land on top of each other.
 *
 * The first cut kept a point when it deviated from the segment joining its KEPT
 * predecessor to its IMMEDIATE successor — a local test that says nothing about
 * where the retained line ends up. It was dismissed in its own comment as
 * "marginally better point count for a lot more machinery" against
 * Douglas-Peucker, which is backwards: the recursive form's value is that it
 * BOUNDS THE ERROR. Asked for 0.25pt the local one moved the violin's outline
 * 0.845pt; asked for 0.5 it moved it 3.18 — six times over. The SVG snapshot
 * caught it, showing a 56-point horizontal jump where twenty samples had been.
 *
 * This is the test that should have caught it, and it is the only property
 * worth asserting: every ORIGINAL point stays within `tol` of the line that
 * comes back. A point count is not a claim about a shape.
 */

type P = { x: number; y: number };

/** Distance from p to the SEGMENT a-b — not to the infinite line through it. */
const segDist = (p: P, a: P, b: P) => {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const l2 = dx * dx + dy * dy;
  if (l2 < 1e-12) return Math.hypot(p.x - a.x, p.y - a.y);
  const t = Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / l2));
  return Math.hypot(p.x - (a.x + t * dx), p.y - (a.y + t * dy));
};

const worstDeviation = (orig: P[], kept: P[]) =>
  Math.max(...orig.map((p) => Math.min(...kept.slice(0, -1).map((_, i) => segDist(p, kept[i], kept[i + 1])))));

/** A violin-ish body: a smooth bell sampled fine, mirrored into a closed loop. */
const bell = (n: number): P[] => {
  const half: P[] = [];
  for (let i = 0; i <= n; i++) {
    const t = (i / n) * 6 - 3;
    half.push({ x: 100 + 40 * Math.exp(-0.5 * t * t), y: 200 - i * 4 });
  }
  return [...half, ...half.map((p) => ({ x: 200 - p.x, y: p.y })).reverse()];
};

describe("thinning an outline keeps it where it was", () => {
  it("never moves a point further than the tolerance asked for", () => {
    const outline = bell(40);
    for (const tol of [0.1, 0.25, 0.5, 1, 2]) {
      const kept = thinOutline(outline, tol);
      expect(kept.length, `tol ${tol} kept everything, so it thinned nothing`).toBeLessThan(outline.length);
      expect(
        worstDeviation(outline, kept),
        `tol ${tol} let the outline move further than it promised`,
      ).toBeLessThanOrEqual(tol + 1e-9);
    }
  });

  it("collapses a straight run to its two ends", () => {
    /**
     * Where the saving actually comes from, stated as the property rather than
     * as a ratio. A KDE tail is a long nearly-straight run of samples that each
     * cost a shape and none of which change the outline.
     *
     * NOT asserted as "thins the bell above by half": measured, that curve
     * keeps 46 of 82 at a quarter point, because it is steeper than a real
     * violin's and those points are genuinely needed. A test written to the
     * ratio I hoped for would have been a test I then loosened.
     */
    const run: P[] = [];
    for (let i = 0; i <= 20; i++) run.push({ x: i * 5, y: 100 });
    run.push({ x: 100, y: 40 });
    const kept = thinOutline(run, 0.25);
    expect(kept, "a straight run kept more than its ends and its corner").toHaveLength(3);
  });

  it("keeps both ends, so a closed outline still closes", () => {
    const outline = bell(40);
    const kept = thinOutline(outline, 1);
    expect(kept[0]).toEqual(outline[0]);
    expect(kept[kept.length - 1]).toEqual(outline[outline.length - 1]);
  });

  it("returns a short or untolerated outline untouched", () => {
    const two = [
      { x: 0, y: 0 },
      { x: 1, y: 1 },
    ];
    expect(thinOutline(two, 1)).toBe(two);
    const outline = bell(10);
    expect(thinOutline(outline, 0), "a tolerance of zero must not drop anything").toBe(outline);
  });

  it("keeps a corner however sharp the turn", () => {
    // A spike is exactly what a local test loses and a chord test keeps: the
    // two points either side of it are close together, and the point itself is
    // far from the line between them.
    const spike = [
      { x: 0, y: 0 },
      { x: 10, y: 0 },
      { x: 11, y: 40 },
      { x: 12, y: 0 },
      { x: 22, y: 0 },
    ];
    expect(thinOutline(spike, 0.25)).toContainEqual({ x: 11, y: 40 });
  });
});
