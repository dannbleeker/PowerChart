import { describe, expect, it } from "vitest";
import { contrastInk } from "../src/core/scene";
import { toRgb, toHex6, alphaOf } from "../src/core/color";

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
