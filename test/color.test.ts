import { describe, expect, it } from "vitest";
import { contrastInk } from "../src/core/scene";
import { toRgb, toHex6, alphaOf, lerpColor } from "../src/core/color";
import { hex as pptxHex, alphaOf as pptxAlphaOf } from "../skill/scripts/pptx-paint.mjs";

/** Colour parsing and normalisation — every paint form the config admits. */

describe("contrastInk", () => {
  it("expands 3-digit hex and picks readable ink", () => {
    expect(contrastInk("#fff")).toBe(contrastInk("#ffffff"));
    expect(contrastInk("#000")).not.toBe(contrastInk("#fff"));
  });

  it("reads rgb()/hsl() fills, not as black", () => {
    // A hex-only parser returned NaN->0 (pure black) for functional colours, so
    // a near-white rgb() fill wrongly got WHITE ink. These forms are valid config
    // (Series.color etc. are plain strings the renderer's PAINT_OK admits).
    expect(contrastInk("rgb(250,250,250)")).toBe(contrastInk("#fafafa"));
    expect(contrastInk("rgb(250,250,250)")).toBe("#0b0b0b"); // dark ink on near-white
    expect(contrastInk("rgb(20,20,20)")).toBe("#ffffff"); // white ink on near-black
    expect(contrastInk("hsl(0,0%,100%)")).toBe("#0b0b0b"); // hsl white
  });
});

describe("toHex6 / alphaOf normalize colours for the PowerPoint renderers", () => {
  it("normalizes every allow-listed form to a 6-digit hex", () => {
    expect(toHex6("#4e79a7")).toBe("#4e79a7"); // identity for the hex the engine emits
    expect(toHex6("#abc")).toBe("#aabbcc");
    expect(toHex6("#4e79a780")).toBe("#4e79a7"); // alpha byte dropped
    expect(toHex6("rgb(78,121,167)")).toBe("#4e79a7");
    expect(toHex6("hsl(0,0%,100%)")).toBe("#ffffff");
    expect(toHex6("not-a-color")).toBe("#808080"); // named/unknown → grey, never black
  });

  it("reads the opacity a paint carries, 1 when opaque", () => {
    expect(alphaOf("#4e79a7")).toBe(1);
    expect(alphaOf("#4e79a780")).toBeCloseTo(128 / 255, 5);
    expect(alphaOf("rgba(1,2,3,0.5)")).toBe(0.5);
    expect(alphaOf("hsla(0,0%,0%,0.25)")).toBe(0.25);
    expect(alphaOf("rgb(1,2,3)")).toBe(1);
  });
});

describe("toRgb parses every allow-listed colour form", () => {
  it("matches hex for the equivalent rgb()/hsl(), strips alpha, expands short hex", () => {
    expect(toRgb("rgb(78,121,167)")).toEqual(toRgb("#4e79a7"));
    expect(toRgb("#abc")).toEqual(toRgb("#aabbcc"));
    expect(toRgb("#4e79a780")).toEqual(toRgb("#4e79a7")); // alpha byte dropped
    expect(toRgb("rgb(50%,50%,50%)")).toEqual([127, 127, 127]);
    expect(toRgb("hsl(0,0%,0%)")).toEqual([0, 0, 0]);
    // A malformed paint falls back to mid grey, never NaN/black.
    expect(toRgb("not-a-color")).toEqual([128, 128, 128]);
  });
});

describe("a colour string that is not a number where a number was expected", () => {
  /**
   * `parseFloat` is looser than the regex that feeds it. `[\d.]+` matches a
   * bare "." and a "..", and `parseFloat(".")` is NaN — so `hsl(., 50%, 50%)`
   * reached the maths as NaN.
   *
   * In the hsl branch that was a CRASH rather than a wrong colour:
   * `Math.floor(NaN / 60) % 6` is NaN, the hue sector table has no NaN entry,
   * and destructuring `undefined` throws. In the rgb branch it produced
   * `#NaNNaNNaN`, which is not a colour anywhere.
   *
   * Ordinary input, too: a config arrives from the JSON box, a saved template,
   * a shape tag written in another deck, and the skill's caller — the same
   * boundary that made `categories: [2023, 2024]` crash.
   */
  const BROKEN = [
    "hsl(., 50%, 50%)",
    "hsl(-.., 50%, 50%)",
    "hsla(., ., .)",
    "rgb(., ., .)",
    "rgba(-.., 1, 1)",
    "hsl(50, ., .)",
  ];

  it("never throws, whatever it is handed", () => {
    for (const c of BROKEN) expect(() => toRgb(c), `toRgb threw on ${c}`).not.toThrow();
  });

  it("falls back to mid grey rather than to NaN", () => {
    for (const c of BROKEN) {
      expect(toRgb(c), `${c} produced a non-colour`).toEqual([128, 128, 128]);
      expect(toHex6(c), `${c} produced a non-colour`).toBe("#808080");
    }
  });

  it("always produces six hex digits, from any channel value at all", () => {
    // Guarded in `rgbToHex` rather than at its four call sites: every producer
    // of a channel can arrive with a NaN for its own reasons, and `lerpColor`
    // does — an interpolation weight of NaN clamps to NaN, not to 0 or 1.
    expect(toHex6("rgb(., ., .)")).toMatch(/^#[0-9a-f]{6}$/);
    expect(lerpColor("#000000", "#ffffff", NaN)).toMatch(/^#[0-9a-f]{6}$/);
    expect(lerpColor("#000000", "#ffffff", Infinity)).toMatch(/^#[0-9a-f]{6}$/);
  });

  it("still reads the colours that were always fine", () => {
    // The negative control. A parser that answered grey to everything would
    // pass all three cases above and silently grey out every chart.
    expect(toRgb("hsl(-30, 50%, 50%)")).toEqual([191, 64, 128]);
    expect(toRgb("rgb(10, 20, 30)")).toEqual([10, 20, 30]);
    expect(toHex6("#abc")).toBe("#aabbcc");
  });
});

describe("a percentage ALPHA is not a percentage channel", () => {
  /**
   * `rgba(200,100,50,50%)` is legal CSS and legal config. The scale test was
   * `/%/.test(<whole string>)`, so the alpha's percent sign was read as
   * evidence that the CHANNELS were percentages: every one got multiplied by
   * 2.55 and clipped, turning a mid orange into near-white.
   *
   * The pptx sink was fixed for this and says so in a comment; this sink was
   * not — and this is the one `officeHex` calls, so the live add-in drew the
   * washed-out colour while the skill's headless deck drew the right one.
   */
  it("reads the channels the same with a % alpha as with a decimal one", () => {
    expect(toRgb("rgba(200,100,50,50%)")).toEqual(toRgb("rgba(200,100,50,0.5)"));
    expect(toHex6("rgba(200, 100, 50, 50%)")).toBe("#c86432");
    expect(toRgb("hsla(200, 50%, 40%, 50%)")).toEqual(toRgb("hsla(200, 50%, 40%, 0.5)"));
  });

  it("still reads the alpha, and still scales percentage CHANNELS", () => {
    // The negative control: dropping the % test altogether would pass the case
    // above and break every percentage-channel colour.
    expect(alphaOf("rgba(200,100,50,50%)")).toBe(0.5);
    expect(toRgb("rgba(100%, 0%, 0%, 50%)")).toEqual([255, 0, 0]);
    expect(toRgb("rgb(50%,50%,50%)")).toEqual([127, 127, 127]);
  });
});

describe("the preview sink and the pptx sink answer the same colour", () => {
  /**
   * CLAUDE.md: there are three colour sinks and they are separate code on
   * purpose — `src/core/color.ts` (preview, and `officeHex` in the live
   * renderer via `toHex6`), `skill/scripts/pptx-paint.mjs` (headless pptx) and
   * the Office pass-through. The same defect has now been found in all three
   * independently, and the percentage-alpha bug above lived in ONE of them for
   * as long as it did because nothing ever compared them.
   *
   * Named colours are the one declared divergence: the pptx sink carries the
   * CSS table, this one answers mid grey (documented on `toRgb`), and the live
   * renderer hands the name to Office untouched. Everything else must agree.
   */
  const FORMS = [
    "#4e79a7",
    "#ABC",
    "#4e79a780",
    "#abcd",
    "  #4e79a7  ",
    "rgb(78,121,167)",
    "rgb(78, 121, 167)",
    "rgba(78,121,167,0.5)",
    "rgba(200,100,50,50%)",
    "rgb(50%,50%,50%)",
    "rgba(100%,0%,0%,50%)",
    "hsl(0,0%,100%)",
    "hsl(210,60%,40%)",
    "hsla(210,60%,40%,0.25)",
    "hsla(210,60%,40%,25%)",
    "hsl(-30, 50%, 50%)",
    "hsl(400, 50%, 50%)",
    "rgb(300,-20,50)",
    "rgb(", // a truncated function is not a function to either sink: black both sides
    // Named colours. This sink read every one of them as mid grey until the
    // table it shares with the pptx sink was carried here too — a gap that put
    // white ink on light fills and flipped `background: "white"` into
    // dark-canvas mode. Listed here so the two tables are pinned to each other.
    "white",
    "black",
    "rebeccapurple",
    "lightyellow",
    "  LightYellow  ",
    "STEELBLUE",
    "darkgrey",
  ];

  /** Paints neither sink can read — the one place they answer differently. */
  const UNREADABLE_FORMS = ["rgb(., ., .)", "hsl(., 50%, 50%)", "", "not-a-color"];

  it("normalises every functional and hex form identically", () => {
    // Case only: the pptx sink echoes a hex literal's own case, and OOXML does
    // not care. Everything else about the six digits must match.
    for (const c of FORMS) {
      expect(toHex6(c).slice(1), `toHex6 vs pptx hex for ${JSON.stringify(c)}`).toBe(pptxHex(c).toLowerCase());
    }
  });

  it("reads the same opacity from every form", () => {
    for (const c of [...FORMS, ...UNREADABLE_FORMS, "transparent", "TRANSPARENT"]) {
      expect(alphaOf(c), `alphaOf vs pptx alphaOf for ${JSON.stringify(c)}`).toBeCloseTo(pptxAlphaOf(c), 6);
    }
  });

  it("differs on an UNREADABLE paint, and only there — grey here, black in the deck", () => {
    // Declared, not accidental: this sink answers mid grey on purpose (so
    // `contrastInk` still picks a sane ink), the pptx sink answers black on
    // purpose (its security note requires six hex digits and black is the
    // documented fallback). Pinned so the difference stays a decision.
    for (const c of UNREADABLE_FORMS) {
      expect(toHex6(c), `${JSON.stringify(c)}`).toBe("#808080");
      expect(pptxHex(c), `${JSON.stringify(c)}`).toBe("000000");
    }
  });
});
