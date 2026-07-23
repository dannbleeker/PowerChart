import { describe, expect, it } from "vitest";
// The pure paint/node helpers extracted from render-pptx.mjs. The CLI runs as a
// subprocess (unmeasurable by v8); this module is imported in-process, so the
// third renderer's colour normalisation and scene→pptx node mapping finally get
// direct assertions and coverage instead of only black-box XML checks.
import { hex, alphaOf, fillOf, visible, hslToHex, makeAddNode } from "../skill/scripts/pptx-paint.mjs";

/** Records the PptxgenJS calls a node mapping makes, so the mapping is assertable. */
function recorder() {
  const shapes: { type: string; opts: Record<string, unknown> }[] = [];
  const texts: { text: string; opts: Record<string, unknown> }[] = [];
  return {
    shapes,
    texts,
    addShape(type: string, opts: Record<string, unknown>) {
      shapes.push({ type, opts });
    },
    addText(text: string, opts: Record<string, unknown>) {
      texts.push({ text, opts });
    },
  };
}

/** A stand-in engine: the four helpers makeAddNode binds, with predictable output. */
const engine = {
  dashKind: (d: number[]): "dot" | "dash" => (d[0] < 2 ? "dot" : "dash"),
  // Outer points first (they carry moveTo), then inner — even length, half each.
  annularSectorPoints: (cx: number, cy: number, innerR: number, r: number) => [
    { x: cx + r, y: cy },
    { x: cx, y: cy + r },
    { x: cx + innerR, y: cy },
    { x: cx, y: cy + innerR },
  ],
  SYMBOL_PRESET: { circle: "ellipse", square: "rect" } as Record<string, string>,
  arrowheadBox: (x: number, y: number, size: number, angle: number) => ({
    left: x - size / 2,
    top: y - size / 2,
    size,
    rotation: angle,
  }),
};
const addNode = makeAddNode(engine);

describe("hex — normalises any allow-listed paint to 6 hex digits", () => {
  it("passes through 6-digit hex and drops an 8-digit alpha", () => {
    expect(hex("#4682b4")).toBe("4682b4");
    expect(hex("#4682b480")).toBe("4682b4");
  });

  it("expands 3- and 4-digit shorthand", () => {
    expect(hex("#abc")).toBe("aabbcc");
    expect(hex("#abcd")).toBe("aabbcc");
  });

  it("parses rgb() as 0–255 and rgb(%) as 0–100, ignoring a percentage alpha", () => {
    expect(hex("rgb(70, 130, 180)")).toBe("4682b4");
    expect(hex("rgb(100%, 0%, 0%)")).toBe("ff0000");
    // The 50% here is ALPHA — it must not rescale the RGB channels to white.
    expect(hex("rgba(100, 150, 200, 50%)")).toBe("6496c8");
  });

  it("parses hsl() through the colour wheel", () => {
    expect(hex("hsl(0, 100%, 50%)")).toBe("ff0000");
    expect(hslToHex(120, 100, 50)).toBe("00ff00");
  });

  it("resolves CSS names and falls back to black for the unknown", () => {
    expect(hex("steelblue")).toBe("4682b4");
    expect(hex("red")).toBe("ff0000");
    expect(hex("not-a-colour")).toBe("000000");
    expect(hex("")).toBe("000000");
  });

  it("never emits anything but six hex digits (OOXML injection guard)", () => {
    for (const c of ['#000"/><a:x', "rgb(999,999,999)", "hsl(720, 300%, 300%)", "#zzz"]) {
      expect(hex(c)).toMatch(/^[0-9a-f]{6}$/);
    }
  });
});

describe("alphaOf / fillOf / visible", () => {
  it("reads alpha from 8-/4-digit hex, rgba()/hsla(), and the transparent keyword", () => {
    expect(alphaOf("#00000080")).toBeCloseTo(128 / 255, 5);
    expect(alphaOf("#0008")).toBeCloseTo(0x88 / 255, 5);
    expect(alphaOf("rgba(0,0,0,0.25)")).toBeCloseTo(0.25, 5);
    expect(alphaOf("hsla(0,0%,0%,50%)")).toBeCloseTo(0.5, 5);
    expect(alphaOf("transparent")).toBe(0);
    expect(alphaOf("#123456")).toBe(1);
    expect(alphaOf("rgba(0,0,0,junk)")).toBe(1); // non-finite → opaque
  });

  it("folds alpha into OOXML transparency, and emits no fill when fully clear", () => {
    expect(fillOf("#4682b4")).toEqual({ color: "4682b4" }); // opaque → bare colour
    const semi = fillOf("#4682b480") as { color: string; transparency: number };
    expect(semi.color).toBe("4682b4");
    expect(semi.transparency).toBeGreaterThan(0);
    expect(fillOf("transparent")).toEqual({ type: "none" });
    // fillOpacity multiplies the paint's own alpha.
    expect((fillOf("#000000", 0) as { type: string }).type).toBe("none");
  });

  it("visible is true only for a present, non-transparent paint", () => {
    expect(visible("#000000")).toBe(true);
    expect(visible("transparent")).toBe(false);
    expect(visible(undefined)).toBe(false);
  });
});

describe("addNode — maps each scene node kind to PptxgenJS", () => {
  it("rect: solid vs hollow fill, and a stroke only when wide enough", () => {
    const r1 = recorder();
    addNode(r1, { kind: "rect", x: 10, y: 20, w: 30, h: 40, fill: "#ff0000", stroke: "#000000", strokeWidth: 1 }, 1, 2);
    expect(r1.shapes[0].type).toBe("rect");
    expect(r1.shapes[0].opts.fill).toEqual({ color: "ff0000" });
    expect(r1.shapes[0].opts.line).toEqual({ color: "000000", width: 1 });

    const r2 = recorder();
    addNode(r2, { kind: "rect", x: 0, y: 0, w: 5, h: 5, fill: "none", stroke: "#000000", strokeWidth: 0 }, 0, 0);
    expect(r2.shapes[0].opts.fill).toEqual({ type: "none" }); // hollow
    expect(r2.shapes[0].opts.line).toEqual({ type: "none" }); // zero width → no line
  });

  it("line: flips vertically for a rising segment and maps a dotted dash", () => {
    const rising = recorder();
    addNode(rising, { kind: "line", x1: 0, y1: 10, x2: 10, y2: 0, stroke: "#000000", dash: [1, 1] }, 0, 0);
    expect(rising.shapes[0].type).toBe("line");
    expect(rising.shapes[0].opts.flipV).toBe(true);
    expect((rising.shapes[0].opts.line as { dashType: string }).dashType).toBe("sysDot");

    const falling = recorder();
    addNode(falling, { kind: "line", x1: 0, y1: 0, x2: 10, y2: 10, stroke: "#000000" }, 0, 0);
    expect(falling.shapes[0].opts.flipV).toBe(false);
  });

  it("text: emits addText with the font and colour", () => {
    const r = recorder();
    addNode(
      r,
      { kind: "text", x: 0, y: 0, w: 100, h: 20, text: "Hi", fontSize: 12, color: "#333333", bold: true },
      0,
      0,
    );
    expect(r.texts[0].text).toBe("Hi");
    expect(r.texts[0].opts.color).toBe("333333");
    expect(r.texts[0].opts.bold).toBe(true);
  });

  it("ellipse: stroke-only ring when fill is none", () => {
    const r = recorder();
    addNode(
      r,
      { kind: "ellipse", cx: 50, cy: 50, rx: 10, ry: 10, fill: "none", stroke: "#000000", strokeWidth: 1 },
      0,
      0,
    );
    expect(r.shapes[0].type).toBe("ellipse");
    expect(r.shapes[0].opts.fill).toEqual({ type: "none" });
  });

  it("polygon: a closed custGeom path with a translucent fill", () => {
    const r = recorder();
    addNode(
      r,
      {
        kind: "polygon",
        points: [
          { x: 0, y: 0 },
          { x: 10, y: 0 },
          { x: 5, y: 10 },
        ],
        fill: "#00ff00",
        fillOpacity: 0.5,
      },
      0,
      0,
    );
    expect(r.shapes[0].type).toBe("custGeom");
    const pts = r.shapes[0].opts.points as { moveTo?: boolean; close?: boolean }[];
    expect(pts[0].moveTo).toBe(true);
    expect(pts[pts.length - 1].close).toBe(true);
    expect((r.shapes[0].opts.fill as { transparency: number }).transparency).toBeGreaterThan(0);
  });

  it("wedge: a full disc becomes an ellipse, a sector becomes a custGeom fan", () => {
    const disc = recorder();
    addNode(
      disc,
      { kind: "wedge", cx: 50, cy: 50, r: 20, innerR: 0, startAngle: 0, endAngle: 360, fill: "#123456" },
      0,
      0,
    );
    expect(disc.shapes[0].type).toBe("ellipse");

    const sector = recorder();
    addNode(
      sector,
      { kind: "wedge", cx: 50, cy: 50, r: 20, innerR: 8, startAngle: 0, endAngle: 90, fill: "#123456" },
      0,
      0,
    );
    expect(sector.shapes[0].type).toBe("custGeom");
    const pts = sector.shapes[0].opts.points as { moveTo?: boolean; close?: boolean }[];
    expect(pts[0].moveTo).toBe(true);
    expect(pts[pts.length - 1].close).toBe(true);
  });

  it("chevron: homePlate when flat-left, chevron otherwise", () => {
    const flat = recorder();
    addNode(flat, { kind: "chevron", x: 0, y: 0, w: 40, h: 20, flatLeft: true, fill: "#000000" }, 0, 0);
    expect(flat.shapes[0].type).toBe("homePlate");

    const arrow = recorder();
    addNode(arrow, { kind: "chevron", x: 0, y: 0, w: 40, h: 20, flatLeft: false, fill: "#000000" }, 0, 0);
    expect(arrow.shapes[0].type).toBe("chevron");
  });

  it("symbol: maps the marker shape through the engine's preset table", () => {
    const r = recorder();
    addNode(
      r,
      { kind: "symbol", shape: "circle", cx: 10, cy: 10, size: 4, fill: "#000000", stroke: "#000000", strokeWidth: 1 },
      0,
      0,
    );
    expect(r.shapes[0].type).toBe("ellipse"); // circle → ellipse preset
  });

  it("arrowhead: a rotated triangle at the box the engine computes", () => {
    const r = recorder();
    addNode(r, { kind: "arrowhead", x: 30, y: 40, size: 6, angle: 90, fill: "#000000" }, 1, 1);
    expect(r.shapes[0].type).toBe("triangle");
    expect(r.shapes[0].opts.rotate).toBe(90);
  });

  it("polygon: no fill and no stroke draws an empty custGeom outline", () => {
    const r = recorder();
    addNode(
      r,
      {
        kind: "polygon",
        points: [
          { x: 0, y: 0 },
          { x: 4, y: 0 },
          { x: 2, y: 4 },
        ],
      },
      0,
      0,
    );
    expect(r.shapes[0].opts.fill).toEqual({ type: "none" });
    expect(r.shapes[0].opts.line).toEqual({ type: "none" });
  });

  it("polygon and wedge honour a visible stroke", () => {
    const poly = recorder();
    addNode(
      poly,
      {
        kind: "polygon",
        points: [
          { x: 0, y: 0 },
          { x: 4, y: 0 },
        ],
        stroke: "#ff0000",
        strokeWidth: 2,
      },
      0,
      0,
    );
    expect(poly.shapes[0].opts.line).toEqual({ color: "ff0000", width: 2 });

    const wedge = recorder();
    addNode(
      wedge,
      { kind: "wedge", cx: 5, cy: 5, r: 4, innerR: 0, startAngle: 0, endAngle: 90, fill: "#000000", stroke: "#ff0000" },
      0,
      0,
    );
    expect((wedge.shapes[0].opts.line as { color: string }).color).toBe("ff0000");
  });

  it("symbol without a stroke width draws no outline", () => {
    const r = recorder();
    addNode(r, { kind: "symbol", shape: "square", cx: 5, cy: 5, size: 3, fill: "#000000", strokeWidth: 0 }, 0, 0);
    expect(r.shapes[0].type).toBe("rect");
    expect(r.shapes[0].opts.line).toEqual({ type: "none" });
  });

  it("ignores an unknown node kind without throwing", () => {
    const r = recorder();
    expect(() => addNode(r, { kind: "mystery" } as unknown as { kind: string }, 0, 0)).not.toThrow();
    expect(r.shapes).toHaveLength(0);
  });
});
