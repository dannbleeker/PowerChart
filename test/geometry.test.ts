import { describe, expect, it } from "vitest";
import {
  polar,
  arrowheadBox,
  sceneToOoxmlPieAngle,
  wedgeFanSteps,
  wedgeFanChord,
  annularSectorPoints,
  symbolPoints,
  markerScale,
  SYMBOL_PRESET,
  type MarkerSymbol,
  type SymbolShape,
} from "../src/core/geometry";

/**
 * These reference implementations are byte-for-byte copies of the formulas that
 * used to live inline in each renderer (svg.ts / powerpoint.ts /
 * render-pptx.mjs) before they were extracted here. They guard the extraction:
 * the shared helpers must keep producing identical output, since the three
 * renderers' snapshots / OOXML depend on it.
 */
const refPolar = (cx: number, cy: number, r: number, angleDeg: number) => {
  const a = ((angleDeg - 90) * Math.PI) / 180;
  return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
};
const refArrow = (x: number, y: number, size: number, angle: number) => {
  const s = size * 2;
  const rotation = (((angle + 90) % 360) + 360) % 360;
  const rad = (rotation * Math.PI) / 180;
  const bx = x - (s / 2) * Math.sin(rad);
  const by = y + (s / 2) * Math.cos(rad);
  return { left: bx - s / 2, top: by - s / 2, size: s, rotation };
};
const refPie = (deg: number) => (((deg - 90) % 360) + 360) % 360;
const refFan = (r: number, span: number) => {
  const stepDeg = Math.max(3, Math.min(12, (2 * Math.sqrt((2 * 0.5) / Math.max(r, 1)) * 180) / Math.PI));
  const steps = Math.max(1, Math.min(60, Math.ceil(span / stepDeg)));
  return { steps, step: span / steps };
};

describe("polar", () => {
  it("places 0/90/180/270° at 12/3/6/9 o'clock", () => {
    expect(polar(0, 0, 10, 0)).toEqual({ x: expect.closeTo(0, 9), y: expect.closeTo(-10, 9) });
    expect(polar(0, 0, 10, 90)).toEqual({ x: expect.closeTo(10, 9), y: expect.closeTo(0, 9) });
    expect(polar(0, 0, 10, 180)).toEqual({ x: expect.closeTo(0, 9), y: expect.closeTo(10, 9) });
    expect(polar(0, 0, 10, 270)).toEqual({ x: expect.closeTo(-10, 9), y: expect.closeTo(0, 9) });
  });

  it("honours the centre offset and matches the reference formula", () => {
    for (const [cx, cy, r, a] of [
      [5, 7, 12, 33],
      [-3, 4, 8, 201],
      [100, 100, 40, -47],
      [0, 0, 1, 359.5],
    ] as const) {
      expect(polar(cx, cy, r, a)).toEqual(refPolar(cx, cy, r, a));
    }
  });
});

describe("arrowheadBox", () => {
  it("matches the reference formula across angles (incl. wrap-around)", () => {
    for (const [x, y, size, angle] of [
      [10, 10, 4, 45],
      [0, 0, 3, 0],
      [50, 20, 6, -90],
      [50, 20, 6, 300],
      [12, 34, 5, 450],
      [12, 34, 5, -270],
    ] as const) {
      expect(arrowheadBox(x, y, size, angle)).toEqual(refArrow(x, y, size, angle));
    }
  });

  it("normalises rotation to [0,360) and doubles size", () => {
    expect(arrowheadBox(0, 0, 4, -90).rotation).toBe(0);
    expect(arrowheadBox(0, 0, 4, 300).rotation).toBe(30);
    expect(arrowheadBox(0, 0, 4, 45).rotation).toBe(135);
    expect(arrowheadBox(0, 0, 7, 0).size).toBe(14);
  });

  it("anchors the triangle's top-centre tip on (x,y) after rotation", () => {
    for (const [x, y, size, angle] of [
      [10, 10, 4, 45],
      [70, 30, 8, 210],
      [5, 90, 6, -30],
    ] as const) {
      const b = arrowheadBox(x, y, size, angle);
      const cx = b.left + b.size / 2;
      const cy = b.top + b.size / 2;
      const rad = (b.rotation * Math.PI) / 180;
      expect(cx + (b.size / 2) * Math.sin(rad)).toBeCloseTo(x, 9);
      expect(cy - (b.size / 2) * Math.cos(rad)).toBeCloseTo(y, 9);
    }
  });
});

describe("sceneToOoxmlPieAngle", () => {
  it("rotates scene angles -90° into OOXML's 3-o'clock origin (skill-deck values)", () => {
    expect(sceneToOoxmlPieAngle(0)).toBe(270); // a 75% slice starts here → adj 16200000
    expect(sceneToOoxmlPieAngle(270)).toBe(180); //                         → adj 10800000
  });

  it("normalises to [0,360) for negative, large, and fractional inputs", () => {
    for (const deg of [0, 45, 90, 180, 270, 360, -45, 450, 719.9, 12.5]) {
      const a = sceneToOoxmlPieAngle(deg);
      expect(a).toBe(refPie(deg));
      expect(a).toBeGreaterThanOrEqual(0);
      expect(a).toBeLessThan(360);
    }
  });
});

describe("wedgeFanSteps", () => {
  it("matches the reference formula and returns step = span/steps", () => {
    for (const [r, span] of [
      [30, 90],
      [1, 6],
      [200, 360],
      [10000, 360],
      [15, 12],
    ] as const) {
      const got = wedgeFanSteps(r, span);
      expect(got).toEqual(refFan(r, span));
      expect(got.step).toBeCloseTo(span / got.steps, 12);
    }
  });

  it("clamps the step count between 1 and 60", () => {
    expect(wedgeFanSteps(10000, 360).steps).toBe(60); // fine steps → capped at 60
    expect(wedgeFanSteps(1, 1).steps).toBe(1); // tiny span → floored at 1
  });

  it("wedgeFanChord is wide enough to close the outer arc (no gapped spokes)", () => {
    // The fan tiles ONLY if each shape's width covers the arc chord at the OUTER
    // rim, 2·r·sin(step/2). Sizing at the mid radius (r/2 on a solid slice) left
    // the shapes half-width and the ring rendered as spokes on PowerPoint web.
    for (const r of [20, 60, 117, 300]) {
      const { step } = wedgeFanSteps(r, 104); // a real doughnut slice (~29% of 360)
      const outerArcChord = 2 * r * Math.sin(((step / 2) * Math.PI) / 180);
      expect(wedgeFanChord(r, step)).toBeGreaterThanOrEqual(outerArcChord);
      // And the OLD mid-radius width would NOT cover it on a solid slice — the bug.
      const midRadiusWidth = 2 * (r / 2) * Math.tan(((step / 2) * Math.PI) / 180) + 1;
      expect(midRadiusWidth, `r=${r}`).toBeLessThan(outerArcChord);
    }
  });
});

describe("annularSectorPoints", () => {
  it("samples outer arc forward then inner arc back, matching the reference", () => {
    const [cx, cy, innerR, r, s, e] = [50, 50, 15, 30, 0, 90] as const;
    const span = e - s;
    const steps = Math.max(2, Math.ceil(span / 6));
    const ref: { x: number; y: number }[] = [];
    for (let i = 0; i <= steps; i++) ref.push(refPolar(cx, cy, r, s + (span * i) / steps));
    for (let i = 0; i <= steps; i++) ref.push(refPolar(cx, cy, innerR, e - (span * i) / steps));

    const pts = annularSectorPoints(cx, cy, innerR, r, s, e);
    expect(pts).toEqual(ref);
    expect(pts).toHaveLength(2 * (steps + 1));
    // outer arc opens at (r, startAngle); inner arc opens at (innerR, endAngle)
    expect(pts[0]).toEqual(polar(cx, cy, r, s));
    expect(pts[steps + 1]).toEqual(polar(cx, cy, innerR, e));
  });

  it("uses max(2, ceil(span/6)) segments per arc", () => {
    // tiny span still gets the floor of 2 segments → 3 points per arc → 6 total
    expect(annularSectorPoints(0, 0, 5, 10, 0, 3)).toHaveLength(2 * (2 + 1));
    // 90° → ceil(90/6)=15 segments → 16 points per arc → 32 total
    expect(annularSectorPoints(0, 0, 5, 10, 0, 90)).toHaveLength(2 * (15 + 1));
  });
});

/** Area of a simple polygon (shoelace), sign-independent. */
function shoelace(pts: { x: number; y: number }[]): number {
  let a = 0;
  for (let i = 0; i < pts.length; i++) {
    const p = pts[i];
    const q = pts[(i + 1) % pts.length];
    a += p.x * q.y - q.x * p.y;
  }
  return Math.abs(a) / 2;
}

const SHAPES: SymbolShape[] = ["diamond", "triangle", "plus"];

describe("marker symbols", () => {
  it("every shape stays inside the box the PowerPoint renderers hand the preset", () => {
    // The two PowerPoint renderers place a 2*size box and let the preset fill
    // it. If the SVG outline left that box, the preview would draw a marker
    // bigger than the deck does — the drift this module exists to prevent.
    for (const s of SHAPES) {
      const pts = symbolPoints(s, 100, 50, 10);
      expect(pts.length).toBeGreaterThan(2);
      for (const p of pts) {
        expect(p.x).toBeGreaterThanOrEqual(90);
        expect(p.x).toBeLessThanOrEqual(110);
        expect(p.y).toBeGreaterThanOrEqual(40);
        expect(p.y).toBeLessThanOrEqual(60);
      }
    }
  });

  it("scales with size and translates with the centre", () => {
    for (const s of SHAPES) {
      const a = symbolPoints(s, 0, 0, 1);
      const b = symbolPoints(s, 7, -3, 4);
      a.forEach((p, i) => {
        expect(b[i].x).toBeCloseTo(p.x * 4 + 7, 10);
        expect(b[i].y).toBeCloseTo(p.y * 4 - 3, 10);
      });
    }
  });

  it("degenerate sizes stay finite", () => {
    for (const s of SHAPES) {
      for (const size of [0, -5]) {
        for (const p of symbolPoints(s, 3, 4, size)) {
          expect(Number.isFinite(p.x)).toBe(true);
          expect(Number.isFinite(p.y)).toBe(true);
        }
      }
    }
  });

  it("markerScale equalises AREA against a circle — the bubble's size claim", () => {
    // "area ∝ size" is what a bubble chart asserts. Shape must not perturb it:
    // a diamond inscribed in the same box as a square holds half its ink, so
    // without this the same Size value would read half as large by group.
    // Measured off the real outline, so a wrong constant in MARKER_AREA fails
    // here rather than silently mis-sizing every marker of that shape.
    for (const s of SHAPES) {
      const k = markerScale(s);
      expect(shoelace(symbolPoints(s, 0, 0, k))).toBeCloseTo(Math.PI, 6);
    }
    // circle and square have no outline here; check them against their formulas.
    expect(Math.PI * markerScale("circle") ** 2).toBeCloseTo(Math.PI, 10);
    expect((2 * markerScale("square")) ** 2).toBeCloseTo(Math.PI, 10);
  });

  it("leaves a circle alone, so the default scatter cannot move", () => {
    expect(markerScale("circle")).toBe(1);
  });

  it("names a preset for every symbol shape, and only lowercase OOXML names", () => {
    // One table feeds Office.js (as a GeometricShapeType key) and PptxgenJS (as
    // an addShape name). A typo here is invisible until it reaches PowerPoint.
    for (const s of SHAPES) {
      expect(SYMBOL_PRESET[s]).toMatch(/^[a-z][A-Za-z0-9]*$/);
    }
    expect(Object.keys(SYMBOL_PRESET).sort()).toEqual([...SHAPES].sort());
  });

  it("markerScale covers every MarkerSymbol", () => {
    const all: MarkerSymbol[] = ["circle", "square", ...SHAPES];
    for (const m of all) expect(markerScale(m)).toBeGreaterThan(0);
  });
});
