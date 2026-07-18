/**
 * Color-scale math for heatmaps and tile maps: linear-light RGB
 * interpolation (fixes the muddy midpoint of naive sRGB lerp) and the
 * sequential / diverging value→color scales.
 */

const hexToRgb = (hex: string): [number, number, number] => {
  const h = hex.replace("#", "");
  const n = parseInt(h.length === 3 ? h.replace(/./g, "$&$&") : h, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};

const rgbToHex = (rgb: number[]): string =>
  "#" +
  rgb
    .map((c) =>
      Math.max(0, Math.min(255, Math.round(c)))
        .toString(16)
        .padStart(2, "0"),
    )
    .join("");

const toLin = (c: number) => (c / 255 <= 0.04045 ? c / 255 / 12.92 : Math.pow((c / 255 + 0.055) / 1.055, 2.4));
const toSrgb = (c: number) => 255 * (c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055);

/** Interpolate two hex colors in linear-light RGB (t ∈ [0, 1]). */
export function lerpColor(c0: string, c1: string, t: number): string {
  const a = hexToRgb(c0);
  const b = hexToRgb(c1);
  return rgbToHex(
    [0, 1, 2].map((i) => toSrgb(toLin(a[i]) + (toLin(b[i]) - toLin(a[i])) * Math.max(0, Math.min(1, t)))),
  );
}

/** Fill for cells/tiles with no data. */
export const NO_DATA = "#e6e6e6";

/** Perceptual (linear-light) relative luminance of a hex colour, 0..1. */
const relLuminance = (hex: string): number => {
  const [r, g, b] = hexToRgb(hex).map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};

/**
 * A background panel/zone tint (weekend shading, marginal-total strips, scatter
 * quadrants) that stays subtle on either theme. The charts hardcode tints tuned
 * for a light canvas; on a LIGHT background this returns that literal UNCHANGED
 * (so default charts are byte-identical), but on a DARK background a light-grey
 * box would glare, so mirror the tint's small step away from white into an equal
 * step away from black — a faint lift off the dark canvas instead.
 */
export function zoneFill(background: string, lightFill: string): string {
  if (relLuminance(background) >= 0.5) return lightFill;
  const drop = Math.min(1, Math.max(0, 1 - relLuminance(lightFill)));
  return lerpColor(background, "#ffffff", drop);
}

/**
 * Sequential scale: 12% tint of the color (kept off pure white so a low
 * value never reads as "no data") → the full color.
 */
export function sequentialScale(min: number, max: number, color: string): (v: number) => string {
  const lo = lerpColor("#ffffff", color, 0.12);
  const span = max - min || 1;
  return (v) => lerpColor(lo, color, (v - min) / span);
}

/**
 * Diverging scale through white, symmetric around zero so equal distances
 * from zero get equal intensity on both sides.
 */
export function divergingScale(min: number, max: number, positive: string, negative: string): (v: number) => string {
  const extent = Math.max(Math.abs(min), Math.abs(max)) || 1;
  return (v) => (v >= 0 ? lerpColor("#ffffff", positive, v / extent) : lerpColor("#ffffff", negative, -v / extent));
}
