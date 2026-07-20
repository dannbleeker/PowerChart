/**
 * Color-scale math for heatmaps and tile maps: linear-light RGB
 * interpolation (fixes the muddy midpoint of naive sRGB lerp) and the
 * sequential / diverging value→color scales.
 */

/**
 * Parse any paint the renderer's allow-list accepts into RGB.
 *
 * A hex-only `parseInt(h, 16)` yields NaN for `rgb()`/`hsl()`, and the bitwise
 * ops then coerce that NaN to 0 — so every functional colour silently read as
 * PURE BLACK. Those forms are legitimate config (Series.color / Series.colors /
 * style.palette are plain strings, and svg.ts's PAINT_OK admits them), so the
 * mis-read reached contrastInk (white ink chosen for a near-white fill) and
 * lerpColor (tints lerping toward black instead of the real hue).
 *
 * Named CSS colours stay a known gap: PAINT_OK admits any bare word, and a
 * 148-entry table is not worth carrying here. They fall back to mid grey, which
 * at least yields a sane ink instead of asserting black.
 */
export function toRgb(color: string): [number, number, number] {
  const c = (color ?? "").trim();
  if (c.startsWith("#")) {
    const h = c.slice(1);
    // 4/8-digit forms carry an alpha byte the colour math has no use for.
    const rgb = h.length === 3 || h.length === 4 ? h.slice(0, 3).replace(/./g, "$&$&") : h.slice(0, 6);
    const n = parseInt(rgb, 16);
    return Number.isNaN(n) ? [128, 128, 128] : [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  const nums = (s: string) => (s.match(/-?[\d.]+%?/g) ?? []).map((v) => parseFloat(v));
  if (/^rgba?\(/i.test(c)) {
    const [r = 0, g = 0, b = 0] = nums(c);
    const scale = /%/.test(c) ? 2.55 : 1; // percentages are 0–100, bare numbers 0–255
    return [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v * scale)))) as [number, number, number];
  }
  if (/^hsla?\(/i.test(c)) {
    const [h = 0, s = 0, l = 0] = nums(c);
    return hslToRgb(((h % 360) + 360) % 360, Math.max(0, Math.min(1, s / 100)), Math.max(0, Math.min(1, l / 100)));
  }
  return [128, 128, 128];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  const chroma = (1 - Math.abs(2 * l - 1)) * s;
  const x = chroma * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - chroma / 2;
  const [r, g, b] = [
    [chroma, x, 0],
    [x, chroma, 0],
    [0, chroma, x],
    [0, x, chroma],
    [x, 0, chroma],
    [chroma, 0, x],
  ][Math.floor(h / 60) % 6];
  return [r, g, b].map((v) => Math.round((v + m) * 255)) as [number, number, number];
}

const hexToRgb = toRgb;

/**
 * Any allow-listed paint → a plain 6-digit `#RRGGBB` (no alpha). The PowerPoint
 * renderers' colour sinks accept only `#RRGGBB` (or a named colour), so rgb()/
 * hsl()/3-digit/8-digit forms — all valid config — must be normalised here or
 * they mis-render (Office.js) or fall back to black (pptx). Named colours are a
 * known gap and resolve to mid grey via toRgb; the office renderer passes named
 * colours through natively instead of calling this.
 */
export function toHex6(color: string): string {
  return rgbToHex(toRgb(color));
}

/**
 * Opacity in [0, 1] carried by a paint (8-digit `#RRGGBBAA`, `rgba()`, `hsla()`);
 * 1 for every opaque form. The PowerPoint renderers split this into a shape's
 * `transparency` so an alpha authored in the config isn't silently dropped.
 */
export function alphaOf(color: string): number {
  const c = (color ?? "").trim();
  const hex = /^#([0-9a-fA-F]{4}|[0-9a-fA-F]{8})$/.exec(c);
  if (hex) {
    const h = hex[1];
    const aa = h.length === 4 ? h[3] + h[3] : h.slice(6, 8);
    return parseInt(aa, 16) / 255;
  }
  const fn = /^(?:rgba|hsla)\(([^)]*)\)$/i.exec(c);
  if (fn) {
    const parts = fn[1].split(/[,/]/).map((s) => s.trim());
    if (parts.length >= 4) {
      const a = parts[3].endsWith("%") ? parseFloat(parts[3]) / 100 : parseFloat(parts[3]);
      return Number.isFinite(a) ? Math.max(0, Math.min(1, a)) : 1;
    }
  }
  return 1;
}

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
