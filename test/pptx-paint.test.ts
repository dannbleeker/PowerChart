import { describe, expect, it } from "vitest";
// The pure paint/node helpers extracted from render-pptx.mjs. The CLI runs as a
// subprocess (unmeasurable by v8); this module is imported in-process, so the
// third renderer's colour normalisation and scene→pptx node mapping finally get
// direct assertions and coverage instead of only black-box XML checks.
import {
  hex,
  hexOr,
  alphaOf,
  fillOf,
  visible,
  hslToHex,
  makeAddNode,
  xmlText,
  lineOf,
  textInk,
} from "../skill/scripts/pptx-paint.mjs";
import type { SceneNode } from "../src/core/scene";

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
  // Predictable, but it must still answer `none` the way the real one does —
  // the mapping now asks it whether the line is dashed AT ALL, so a stub that
  // can only say dot/dash would hide a call site that stopped honouring it.
  dashKind: (d: unknown): "none" | "dot" | "dash" =>
    !Array.isArray(d) || !d.some((v) => typeof v === "number" && v > 0) ? "none" : d[0] < 2 ? "dot" : "dash",
  // Outer points first (they carry moveTo), then inner — even length, half each.
  annularSectorPoints: (cx: number, cy: number, innerR: number, r: number) => [
    { x: cx + r, y: cy },
    { x: cx, y: cy + r },
    { x: cx + innerR, y: cy },
    { x: cx, y: cy + innerR },
  ],
  symbolPreset: (shape: string) => ({ circle: "ellipse", square: "rect" })[shape as "circle" | "square"] ?? "ellipse",
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

  it("holds that guarantee for a number that is not a number", () => {
    // The cases above all reached the maths as numbers. These do not: the regex
    // that finds a colour's components matches a bare "." and a "..", and
    // `parseFloat(".")` is NaN. `rgb(., ., .)` returned the nine-character
    // string "NaNNaNNaN", and `hsl(., 50%, 50%)` THREW — the hue sector table
    // has no NaN entry, so destructuring `undefined` blew up.
    //
    // A throw is the worse of the two. It happens inside pptxgenjs at
    // writeFile, outside every per-chart guard, so one bad colour destroys the
    // whole batch — which is exactly what the prototype-named colours below did
    // once already. So the guarantee is enforced where it is stated, at the
    // exit, rather than trusted to each branch.
    for (const c of ["hsl(., 50%, 50%)", "hsl(-.., 50%, 50%)", "rgb(., ., .)", "rgba(-.., 1, 1)", "hsl(50, ., .)"]) {
      expect(() => hex(c), `hex() threw on ${c}`).not.toThrow();
      expect(hex(c), `hex() broke the six-digit guarantee on ${c}`).toMatch(/^[0-9a-f]{6}$/);
    }
    // The negative control: a sink that answered black to everything would pass
    // the line above and paint every deck one colour.
    expect(hex("hsl(-30,50%,50%)")).toBe("bf4080");
    expect(hex("steelblue")).toBe("4682b4");
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

  it("text: writes a font size OOXML can carry, whatever the node says", () => {
    // pptxgenjs emits `sz="${fontSize * 100}"` unchecked, and ST_TextFontSize is
    // 100..400000 — 1pt to 4000pt. A node carrying 0.0001 wrote sz="0" and one
    // carrying 1e6 wrote sz="120000000"; both are decks PowerPoint offers to
    // repair, produced by a CLI reporting success. Same shape as the
    // `x="Infinity"` that `finiteNodes` exists to prevent.
    //
    // Asserted HERE and not only on the engine, because `makeAddNode` is a pure
    // exported function taking a scene node — this route never passes through
    // `normalizeConfig` at all, and this file sits outside `tsconfig.include`
    // where nothing checks it.
    const size = (fontSize: unknown) => {
      const r = recorder();
      addNode(r, { kind: "text", x: 0, y: 0, w: 100, h: 20, text: "Hi", fontSize } as never, 0, 0);
      return r.texts[0].opts.fontSize as number;
    };
    for (const bad of [0.0001, 1e-30, 0, -12, 1e6, Infinity, NaN, undefined, "12pt"]) {
      const got = size(bad);
      expect(got, `fontSize ${String(bad)} wrote ${got}pt, outside OOXML's 1..4000`).toBeGreaterThanOrEqual(1);
      expect(got, `fontSize ${String(bad)} wrote ${got}pt, outside OOXML's 1..4000`).toBeLessThanOrEqual(4000);
    }
    // An ordinary size is passed through untouched, so the bound is a clamp and
    // not a replacement.
    expect(size(12)).toBe(12);
    expect(size(4000)).toBe(4000);
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

  /**
   * The seam, made loud — this is the third of the three CLAUDE.md lists as
   * silent when a `SceneNode` kind is added.
   *
   * Every OTHER renderer fails to build: `nodeToSvg` and `powerpoint.ts`'s
   * `addNode` are exhaustive switches TypeScript checks, and `translateNodes`
   * carries an explicit `never`. This mapping is neither — it lives outside
   * `tsconfig.include`, so nothing typechecks it, and its switch has no default
   * (deliberately: the test above pins that an unknown kind is ignored rather
   * than thrown on). A new kind therefore renders as NOTHING in the skill's
   * .pptx, in a file that opens cleanly and is reported as a success. Missing
   * shapes are the one failure the headless path cannot surface on its own.
   *
   * `Record<SceneNode["kind"], SceneNode>` closes it: the union cannot grow
   * without this map going red, and the assertion below then says whether the
   * mapping actually learned the kind or merely compiles.
   */
  it("draws something for every node kind in the scene contract", () => {
    const EACH: Record<SceneNode["kind"], SceneNode> = {
      rect: { kind: "rect", x: 1, y: 2, w: 3, h: 4, fill: "#123456" },
      line: { kind: "line", x1: 1, y1: 2, x2: 5, y2: 6, stroke: "#123456" },
      text: {
        kind: "text",
        x: 1,
        y: 2,
        w: 30,
        h: 12,
        text: "t",
        fontSize: 10,
        color: "#123456",
        align: "left",
        valign: "top",
      },
      ellipse: { kind: "ellipse", cx: 10, cy: 10, rx: 5, ry: 5, fill: "#123456" },
      wedge: { kind: "wedge", cx: 10, cy: 10, r: 5, innerR: 2, startAngle: 0, endAngle: 90, fill: "#123456" },
      chevron: { kind: "chevron", x: 1, y: 2, w: 30, h: 12, fill: "#123456" },
      arrowhead: { kind: "arrowhead", x: 5, y: 5, angle: 90, size: 3, fill: "#123456" },
      symbol: { kind: "symbol", shape: "diamond", cx: 10, cy: 10, size: 3, fill: "#123456" },
      polygon: {
        kind: "polygon",
        points: [
          { x: 0, y: 0 },
          { x: 5, y: 0 },
          { x: 5, y: 5 },
        ],
        fill: "#123456",
      },
    };

    for (const [kind, node] of Object.entries(EACH)) {
      const r = recorder();
      addNode(r, node, 0, 0);
      expect(
        r.shapes.length + r.texts.length,
        `a "${kind}" node drew nothing — the skill's .pptx would be missing it, silently`,
      ).toBeGreaterThan(0);
    }
  });
});

describe("hostile input never breaks the OOXML contract (regression)", () => {
  it("hex() stays six hex digits for prototype-named colours", () => {
    // CSS_NAMES is a plain object, so these reached Object.prototype and hex()
    // returned a FUNCTION / an object — which detonated inside pptxgenjs at
    // writeFile, OUTSIDE every per-chart guard, destroying the whole batch.
    for (const c of ["__proto__", "constructor", "toString", "valueOf", "hasOwnProperty"]) {
      const out = hex(c);
      expect(typeof out, `hex(${c}) type`).toBe("string");
      expect(out, `hex(${c})`).toMatch(/^[0-9a-f]{6}$/i);
    }
  });

  it("hex() stays six hex digits for a colour that is not a string at all", () => {
    // The same contract, one type further out. A prototype-named colour was
    // caught before; a NUMERIC one — `style.palette: [1, 2, 3]` — was not, and
    // `(c ?? "").trim()` threw rather than returning anything. A config reaches
    // this renderer straight from an agent, so a numeric colour is a plausible
    // mistake, and the throw lands outside every per-chart guard: one bad
    // colour, no deck at all.
    for (const c of [1, 0, null, undefined, {}, [], true, NaN, Symbol.iterator]) {
      const out = hex(c as unknown as string);
      expect(typeof out, `hex(${String(c)}) type`).toBe("string");
      expect(out, `hex(${String(c)})`).toMatch(/^[0-9a-f]{6}$/i);
      expect(alphaOf(c as unknown as string), `alphaOf(${String(c)})`).toBe(1);
    }
  });

  it("xmlText drops what XML forbids and keeps what it allows", () => {
    expect(xmlText(`a${String.fromCharCode(11)}b`)).toBe("ab");
    expect(xmlText("tab\tnewline\n")).toBe("tab\tnewline\n"); // legal whitespace survives
    expect(xmlText("emoji \u{1F4C8}")).toBe("emoji \u{1F4C8}"); // a valid surrogate PAIR survives
    expect(xmlText(`lone \uD83D`)).toBe("lone "); // an unpaired surrogate does not
  });
});

/**
 * A stroke's alpha, which only the fill ever had.
 */
describe("lineOf", () => {
  it("folds an alpha into transparency, like fillOf does", () => {
    // Every stroke site handed pptxgenjs a bare `hex(n.stroke)`, so one
    // translucent series colour rendered at two opacities inside a single chart
    // — fill honoured, outline opaque — and disagreed with both other
    // renderers. SVG emits `stroke="#2a78d659"`; Office.js splits it through
    // `strokeColor` into `transparency: 0.651`. Stroke alpha is not in the
    // parity contract's list of intentional divergences.
    expect(lineOf("#2a78d659")).toEqual({ color: "2a78d6", transparency: 65 });
    expect(lineOf("#2a78d6"), "an opaque stroke gained a transparency key").toEqual({ color: "2a78d6" });
  });
});

/**
 * The slide BACKGROUND is the one paint black is not a safe default for.
 *
 * `hex`'s black fallback is right for ink — a label in an unknown colour is
 * still a label — and catastrophic here: `background: "off-white"` (a typo) or
 * `"transparent"` (a paint SSF Charts documents elsewhere) wrote
 * `<a:srgbClr val="000000"/>` with the chart's own near-black ink on top. A
 * deck of black slides carrying invisible charts, which opens cleanly,
 * validates, and is reported as a success by the CLI. The SVG renderer has
 * always fallen back to white, so the two sinks disagreed about one config.
 */
describe("hexOr — an unrecognised paint falls back to the caller's default", () => {
  it("gives back the fallback, not black, for a paint it does not know", () => {
    for (const bad of ["off-white", "transparent", "not-a-colour", "", null, undefined, 5 as unknown as string]) {
      expect(hexOr(bad, "#ffffff"), `hexOr(${JSON.stringify(bad)})`).toBe("ffffff");
    }
  });

  it("still returns a recognised paint unchanged", () => {
    expect(hexOr("#1a1a1a", "#ffffff")).toBe("1a1a1a");
    expect(hexOr("red", "#ffffff")).toBe("ff0000");
    expect(hexOr("#abc", "#ffffff")).toBe("aabbcc");
    expect(hexOr("rgb(1,2,3)", "#ffffff")).toBe("010203");
  });

  it("returns a genuine black as black — the fallback must not swallow it", () => {
    // The distinction the old code could not make: `readHex` returned "000000"
    // for BOTH a real black and an unrecognised paint, so a fallback keyed on
    // the value could never tell them apart.
    expect(hexOr("#000000", "#ffffff")).toBe("000000");
    expect(hexOr("black", "#ffffff")).toBe("000000");
  });

  it("leaves hex()'s own black guarantee alone", () => {
    // Every other call site depends on six hex digits, whatever arrives.
    for (const bad of ["off-white", "", null, undefined]) {
      expect(hex(bad as unknown as string)).toBe("000000");
    }
  });
});

/**
 * Text was the one paint channel that dropped its alpha.
 *
 * Every fill and stroke on a slide goes through `fillOf`/`lineOf` and carries
 * its transparency into the OOXML; the text case took a bare `hex` and threw
 * the alpha away. So a deliberately muted label — `#0b0b0b80`, an `rgba()`, an
 * `hsla()` — was drawn faint by the SVG renderer and at FULL STRENGTH in the
 * deck, and the two pictures disagreed about which labels mattered.
 *
 * Office.js cannot follow: `font.color` is a hex string with nowhere to put an
 * alpha. That divergence is declared in the parity contract at the top of
 * `src/core/scene.ts` rather than fixed, because it is a host limit.
 */
describe("a text run keeps the alpha its colour carries", () => {
  it("maps alpha to the same 0-100 transparency a shape uses", () => {
    expect(textInk("#0b0b0b80")).toEqual({ color: "0b0b0b", transparency: 50 });
    expect(textInk("rgba(11, 11, 11, 0.5)")).toEqual({ color: "0b0b0b", transparency: 50 });
    expect(textInk("hsla(0, 0%, 0%, 0.25)")).toEqual({ color: "000000", transparency: 75 });
  });

  it("writes an opaque colour exactly as it always did", () => {
    // No `transparency` key at all, so an ordinary label's XML is unchanged.
    expect(textInk("#0b0b0b")).toEqual({ color: "0b0b0b" });
    expect(textInk("steelblue")).toEqual({ color: "4682b4" });
  });

  it("carries it through the node mapping, not just the helper", () => {
    const slide = recorder();
    addNode(
      slide,
      {
        kind: "text",
        x: 0,
        y: 0,
        w: 40,
        h: 12,
        text: "muted",
        fontSize: 10,
        color: "#0b0b0b80",
        align: "left",
        valign: "top",
      } as SceneNode,
      0,
      0,
    );
    expect(slide.texts).toHaveLength(1);
    expect(slide.texts[0].opts.color).toBe("0b0b0b");
    expect(slide.texts[0].opts.transparency).toBe(50);
  });
});
