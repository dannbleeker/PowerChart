/**
 * Pure geometry shared by the three renderers (SVG, Office.js, PptxgenJS).
 *
 * No renderer, Office, or scene imports — just math on scene coordinates. The
 * tip-anchoring, pie-angle, and wedge-sampling formulas below used to be
 * copy-pasted per renderer, which is how they drifted (an arrowhead anchored
 * differently in each; an inner radius honoured in one renderer and dropped in
 * the others). Keeping them here, unit-tested, makes that class of bug
 * impossible. Angle convention matches the scene graph: 0° = 12 o'clock,
 * clockwise.
 */

/** Point on a circle for wedge geometry (0° = 12 o'clock, clockwise). */
export function polar(cx: number, cy: number, r: number, angleDeg: number): { x: number; y: number } {
  const a = ((angleDeg - 90) * Math.PI) / 180;
  return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
}

/**
 * Placement of an arrowhead drawn as a rotated geometric triangle whose tip
 * must land on the scene point (x, y). Office.js and PptxgenJS have no freeform
 * tip anchor, so each drops a `size*2`-square triangle (whose tip sits at the
 * box's top-centre) and rotates it about the box centre — the SVG renderer,
 * which draws the exact triangle, anchors the tip instead. Returns the
 * pre-offset box top-left corner {left, top}, the square side, and the rotation
 * in degrees; the caller adds its own frame offset and unit scale.
 *
 * `angle` is the ArrowheadNode angle (0 = east, clockwise); `size` is its size
 * (the triangle side is 2·size).
 */
export function arrowheadBox(
  x: number,
  y: number,
  size: number,
  angle: number,
): { left: number; top: number; size: number; rotation: number } {
  const s = size * 2;
  const rotation = (((angle + 90) % 360) + 360) % 360;
  const rad = (rotation * Math.PI) / 180;
  // Offset the box so the top-centre tip lands on (x, y) after rotating.
  const bx = x - (s / 2) * Math.sin(rad);
  const by = y + (s / 2) * Math.cos(rad);
  return { left: bx - s / 2, top: by - s / 2, size: s, rotation };
}

/**
 * Scene angle (0 = 12 o'clock, clockwise) → OOXML pie preset angle
 * (0 = 3 o'clock, clockwise), normalised to [0, 360). Feeds the PptxgenJS pie
 * shape's `angleRange`.
 */
export function sceneToOoxmlPieAngle(deg: number): number {
  return (((deg - 90) % 360) + 360) % 360;
}

/**
 * Step count for the Office.js wedge fan. Office.js has no freeform paths, so a
 * pie/annular slice is approximated by a fan of rotated triangles/rectangles.
 * Density keeps the chord sagitta under ~0.5pt (stepDeg ≈ 2·√(2·tol/r) rad),
 * clamped so the shape count per wedge stays bounded. Returns the step count
 * and the per-step angular width.
 */
export function wedgeFanSteps(r: number, span: number): { steps: number; step: number } {
  const stepDeg = Math.max(3, Math.min(12, (2 * Math.sqrt((2 * 0.5) / Math.max(r, 1)) * 180) / Math.PI));
  const steps = Math.max(1, Math.min(60, Math.ceil(span / stepDeg)));
  return { steps, step: span / steps };
}

/**
 * Outline of an annular sector as scene-coordinate points, for a filled
 * PptxgenJS `custGeom` (OOXML's pie preset can't express an inner radius): the
 * outer arc forward from startAngle→endAngle, then the inner arc back, each arc
 * approximated by `max(2, ceil(span/6))` chord segments. The caller marks the
 * first point moveTo, maps to its box origin/unit scale, and appends the close.
 */
export function annularSectorPoints(
  cx: number,
  cy: number,
  innerR: number,
  r: number,
  startAngle: number,
  endAngle: number,
): { x: number; y: number }[] {
  const span = endAngle - startAngle;
  const steps = Math.max(2, Math.ceil(span / 6));
  const pts: { x: number; y: number }[] = [];
  for (let i = 0; i <= steps; i++) {
    pts.push(polar(cx, cy, r, startAngle + (span * i) / steps));
  }
  for (let i = 0; i <= steps; i++) {
    pts.push(polar(cx, cy, innerR, endAngle - (span * i) / steps));
  }
  return pts;
}

/**
 * Marker symbol shapes — the ones the scene graph cannot already draw filled.
 *
 * Every shape here reproduces its OOXML preset EXACTLY, and that is the entry
 * requirement. The SVG renderer draws these points while the PowerPoint
 * renderers name the preset and let PowerPoint draw it, so a shape whose
 * points only approximate its preset makes the preview lie about the deck —
 * and, because `markerScale` measures area off these points, it also breaks
 * the bubble's "area ∝ size" claim in the deck while keeping it in the preview.
 *
 * `circle` and `square` are absent: those are EllipseNode and RectNode, and
 * re-expressing them here would fork two shapes that already render correctly
 * everywhere. A SymbolNode carries only the rest.
 *
 * Two shapes were tried and rejected on the rule above:
 *   - `star5` — its preset stretches itself by hf=1.05146 / vf=1.10557 so the
 *     star fills its box (a 5-point star spans 1.902R by 1.809R, and those
 *     factors are exactly 2/1.902 and 2/1.809). An inscribed SVG star is
 *     therefore 16.2% smaller in AREA than the one PowerPoint draws — enough
 *     to make an area-matched star bubble over-read by a sixth in the deck.
 *   - `mathMultiply` (an X) — angled arms, same class of problem. Redundant
 *     with `plus` at 3-4pt anyway.
 */
export type SymbolShape = "diamond" | "triangle" | "plus";

/**
 * A point shape. "circle" and "square" are the scene's existing ellipse and
 * rect; the rest are SymbolNode shapes drawn from PowerPoint preset geometry.
 */
export type MarkerSymbol = "circle" | "square" | SymbolShape;

/** OOXML `plus` arm half-width as a fraction of the box side (its default adj). */
const PLUS_ARM = 25000 / 100000;

/**
 * Area of each shape when inscribed with half-extent 1, i.e. in a 2x2 box.
 * Verified against the shoelace area of `symbolPoints` in the geometry tests —
 * these are the numbers `markerScale` divides by, so a wrong one here silently
 * mis-sizes every marker of that shape.
 */
const MARKER_AREA: Record<MarkerSymbol, number> = {
  circle: Math.PI,
  square: 4,
  diamond: 2,
  triangle: 2,
  // Two 2x1 bars crossing, minus their 1x1 overlap.
  plus: 3,
};

/**
 * Radius multiplier that gives `shape` the same AREA as a circle of the same
 * radius.
 *
 * Bubble radius carries the data — "area ∝ size" is the chart's central
 * quantitative claim — and a shape drawn at a bare radius does not honour it: a
 * star inscribed in the same box as a square holds barely a quarter of its ink,
 * so two bubbles with identical `size` values would differ threefold in area
 * for no reason but their group. Equalising area keeps shape a purely
 * categorical channel that cannot be misread as magnitude. It costs the
 * bounding box: an area-matched star reaches further than the circle it
 * replaces, which is the correct trade — ink is what the eye measures.
 */
export function markerScale(shape: MarkerSymbol): number {
  return Math.sqrt(Math.PI / MARKER_AREA[shape]);
}

/**
 * SymbolShape → OOXML preset geometry name. Both PowerPoint renderers read
 * this same table: Office.js as a `GeometricShapeType` key, PptxgenJS as an
 * `addShape` name — the two vocabularies are the same lowercase OOXML names,
 * so one table serves both and they cannot drift apart.
 */
export const SYMBOL_PRESET: Record<SymbolShape, string> = {
  diamond: "diamond",
  triangle: "triangle",
  plus: "plus",
};

/**
 * Outline of a marker symbol as scene-coordinate points, centred on (cx, cy)
 * and inscribed in the `2*size` square the PowerPoint renderers hand to the
 * preset. The SVG renderer draws these points directly; the other two name the
 * preset instead, so this is the one place the two descriptions of a symbol sit
 * side by side.
 *
 * Each shape reproduces its preset's geometry exactly — see SymbolShape for
 * why that is a hard requirement rather than a nicety.
 */
export function symbolPoints(shape: SymbolShape, cx: number, cy: number, size: number): { x: number; y: number }[] {
  const s = Math.max(0, size);
  switch (shape) {
    case "diamond":
      return [
        { x: cx, y: cy - s },
        { x: cx + s, y: cy },
        { x: cx, y: cy + s },
        { x: cx - s, y: cy },
      ];
    case "triangle":
      return [
        { x: cx, y: cy - s },
        { x: cx + s, y: cy + s },
        { x: cx - s, y: cy + s },
      ];
    case "plus": {
      // Arm half-width, from the preset's adj against the box side (2*size).
      const a = 2 * s * PLUS_ARM;
      return [
        { x: cx - s, y: cy - a },
        { x: cx - a, y: cy - a },
        { x: cx - a, y: cy - s },
        { x: cx + a, y: cy - s },
        { x: cx + a, y: cy - a },
        { x: cx + s, y: cy - a },
        { x: cx + s, y: cy + a },
        { x: cx + a, y: cy + a },
        { x: cx + a, y: cy + s },
        { x: cx - a, y: cy + s },
        { x: cx - a, y: cy + a },
        { x: cx - s, y: cy + a },
      ];
    }
  }
}
