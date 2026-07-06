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
