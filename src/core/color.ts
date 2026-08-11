/**
 * Color-scale math for heatmaps and tile maps: linear-light RGB
 * interpolation (fixes the muddy midpoint of naive sRGB lerp) and the
 * sequential / diverging value→color scales.
 */

/**
 * A paint as text, for anything at all that arrives claiming to be one.
 *
 * `(color ?? "").trim()` guarded null and undefined and nothing else, so a
 * palette of NUMBERS — `style.palette: [1, 2, 3]` — threw
 * `TypeError: (color ?? "").trim is not a function` straight out of the
 * renderer. Every route into this function is user JSON: the pane's style
 * import stores whatever parses, a chart config can be pasted in whole, and the
 * skill hands one over from an agent. The type says `string`; TypeScript checks
 * the code, not the file someone pastes.
 *
 * Anything that is not a string reads as the empty string, which the callers
 * already handle: mid grey from `toRgb`, opaque from `alphaOf`. That is the
 * same answer they give for a named CSS colour, so a bad paint degrades exactly
 * like an unrecognised one rather than taking the chart down.
 */
function paintText(color: unknown): string {
  return typeof color === "string" ? color.trim() : "";
}

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
 * Named CSS colours were a "known gap" here for as long as the table looked
 * like more weight than it was worth, and the gap cost more than the table:
 * every one of them read as MID GREY, which is not a sane fallback — it is a
 * confident wrong answer, and two helpers act on it.
 *
 *  - `contrastInk` chose WHITE ink for a `lightyellow` fill. Grey's luminance
 *    sits just under the threshold, so every light named colour got invisible
 *    labels.
 *  - `relLuminance(style.background)` decides whether the canvas is dark, so
 *    `background: "white"` put every canvas-aware helper into DARK mode: the
 *    same chart written `#ffffff` and `white` came out with different heatmap
 *    tints, zone fills and no-data greys.
 *
 * The table is the one the pptx sink has always carried. The two are separate
 * files on purpose (that one ships standalone inside the skill zip), and
 * `test/color.test.ts` now drives NAMED colours through both sinks, so they are
 * pinned to each other by a test rather than by hope.
 */
/**
 * What an unparseable colour becomes. Mid grey, which is what every branch
 * below already falls back to — named here so the guard and the fallbacks stay
 * the same answer.
 */
/**
 * The CSS Color 4 named colours as "name rrggbb" pairs — the same table the
 * headless pptx sink carries (`skill/scripts/pptx-paint.mjs`, which ships
 * standalone in the skill zip and so cannot import this one). `test/color.test.ts`
 * drives named colours through both, so a change to one fails against the other.
 */
const CSS_NAMES: Record<string, string> = Object.fromEntries(
  `aliceblue f0f8ff, antiquewhite faebd7, aqua 00ffff, aquamarine 7fffd4, azure f0ffff, beige f5f5dc,
  bisque ffe4c4, black 000000, blanchedalmond ffebcd, blue 0000ff, blueviolet 8a2be2, brown a52a2a,
  burlywood deb887, cadetblue 5f9ea0, chartreuse 7fff00, chocolate d2691e, coral ff7f50, cornflowerblue 6495ed,
  cornsilk fff8dc, crimson dc143c, cyan 00ffff, darkblue 00008b, darkcyan 008b8b, darkgoldenrod b8860b,
  darkgray a9a9a9, darkgreen 006400, darkgrey a9a9a9, darkkhaki bdb76b, darkmagenta 8b008b,
  darkolivegreen 556b2f, darkorange ff8c00, darkorchid 9932cc, darkred 8b0000, darksalmon e9967a,
  darkseagreen 8fbc8f, darkslateblue 483d8b, darkslategray 2f4f4f, darkslategrey 2f4f4f, darkturquoise 00ced1,
  darkviolet 9400d3, deeppink ff1493, deepskyblue 00bfff, dimgray 696969, dimgrey 696969, dodgerblue 1e90ff,
  firebrick b22222, floralwhite fffaf0, forestgreen 228b22, fuchsia ff00ff, gainsboro dcdcdc,
  ghostwhite f8f8ff, gold ffd700, goldenrod daa520, gray 808080, green 008000, greenyellow adff2f, grey 808080,
  honeydew f0fff0, hotpink ff69b4, indianred cd5c5c, indigo 4b0082, ivory fffff0, khaki f0e68c,
  lavender e6e6fa, lavenderblush fff0f5, lawngreen 7cfc00, lemonchiffon fffacd, lightblue add8e6,
  lightcoral f08080, lightcyan e0ffff, lightgoldenrodyellow fafad2, lightgray d3d3d3, lightgreen 90ee90,
  lightgrey d3d3d3, lightpink ffb6c1, lightsalmon ffa07a, lightseagreen 20b2aa, lightskyblue 87cefa,
  lightslategray 778899, lightslategrey 778899, lightsteelblue b0c4de, lightyellow ffffe0, lime 00ff00,
  limegreen 32cd32, linen faf0e6, magenta ff00ff, maroon 800000, mediumaquamarine 66cdaa, mediumblue 0000cd,
  mediumorchid ba55d3, mediumpurple 9370db, mediumseagreen 3cb371, mediumslateblue 7b68ee,
  mediumspringgreen 00fa9a, mediumturquoise 48d1cc, mediumvioletred c71585, midnightblue 191970,
  mintcream f5fffa, mistyrose ffe4e1, moccasin ffe4b5, navajowhite ffdead, navy 000080, oldlace fdf5e6,
  olive 808000, olivedrab 6b8e23, orange ffa500, orangered ff4500, orchid da70d6, palegoldenrod eee8aa,
  palegreen 98fb98, paleturquoise afeeee, palevioletred db7093, papayawhip ffefd5, peachpuff ffdab9,
  peru cd853f, pink ffc0cb, plum dda0dd, powderblue b0e0e6, purple 800080, rebeccapurple 663399, red ff0000,
  rosybrown bc8f8f, royalblue 4169e1, saddlebrown 8b4513, salmon fa8072, sandybrown f4a460, seagreen 2e8b57,
  seashell fff5ee, sienna a0522d, silver c0c0c0, skyblue 87ceeb, slateblue 6a5acd, slategray 708090,
  slategrey 708090, snow fffafa, springgreen 00ff7f, steelblue 4682b4, tan d2b48c, teal 008080, thistle d8bfd8,
  tomato ff6347, turquoise 40e0d0, violet ee82ee, wheat f5deb3, white ffffff, whitesmoke f5f5f5, yellow ffff00,
  yellowgreen 9acd32`
    .split(",")
    .map((pair) => pair.trim().split(" ") as [string, string]),
);

const UNREADABLE: [number, number, number] = [128, 128, 128];

/**
 * Is this a CSS colour NAME the table knows?
 *
 * Exported for the live Office renderer, which hands named colours to the host
 * verbatim — Office knows the same names — and needs to tell a name from any
 * other bare word. Own-property only, like every lookup on this table.
 */
export function isNamedColor(color: string): boolean {
  const c = paintText(color).toLowerCase();
  return Object.prototype.hasOwnProperty.call(CSS_NAMES, c);
}

export function toRgb(color: string): [number, number, number] {
  const c = paintText(color);
  if (c.startsWith("#")) {
    const h = c.slice(1);
    // 4/8-digit forms carry an alpha byte the colour math has no use for.
    const rgb = h.length === 3 || h.length === 4 ? h.slice(0, 3).replace(/./g, "$&$&") : h.slice(0, 6);
    const n = parseInt(rgb, 16);
    return Number.isNaN(n) ? UNREADABLE : [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  // `parseFloat` is looser than the regex that fed it: `[\d.]+` matches a bare
  // "." and a "..", and `parseFloat(".")` is NaN. So a colour of `hsl(., 50%,
  // 50%)` — the sort of thing a hand-edited config or a template written in
  // another deck arrives with — used to reach the maths as NaN. In the hsl
  // branch that was a CRASH, not a wrong colour: `Math.floor(NaN / 60) % 6` is
  // NaN, the sector table has no NaN entry, and destructuring `undefined`
  // throws. In the rgb branch it was `#NaNNaNNaN`, which is not a colour and,
  // in the pptx sink, not the six hex digits that path's own security note
  // requires. Both are the same root: a `string` in the types is not a number
  // in the file someone pasted.
  const nums = (s: string) => (s.match(/-?[\d.]+%?/g) ?? []).map((v) => parseFloat(v));
  const finite = (vs: number[]) => vs.every((v) => Number.isFinite(v));
  if (/^rgba?\(/i.test(c)) {
    // Only the r/g/b components decide the 0–255 vs 0–100% scale. Testing the
    // whole string let a percentage ALPHA — the perfectly legal
    // `rgba(200,100,50,50%)` — multiply the CHANNELS by 2.55 and clip the
    // colour to near-white. The pptx sink (`skill/scripts/pptx-paint.mjs`) was
    // fixed for exactly this and carries a note saying so; this sink was not,
    // and it is the one `officeHex` calls — so the same config drew the right
    // colour in the headless deck and a washed-out one in the live add-in.
    const comps = (c.slice(c.indexOf("(") + 1).match(/-?[\d.]+%?/g) ?? []).slice(0, 3);
    const [r = 0, g = 0, b = 0] = comps.map((v) => parseFloat(v));
    if (!finite([r, g, b])) return UNREADABLE;
    const scale = comps.some((v) => v.endsWith("%")) ? 2.55 : 1; // percentages are 0–100, bare numbers 0–255
    return [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v * scale)))) as [number, number, number];
  }
  if (/^hsla?\(/i.test(c)) {
    const [h = 0, s = 0, l = 0] = nums(c);
    if (!finite([h, s, l])) return UNREADABLE;
    return hslToRgb(((h % 360) + 360) % 360, Math.max(0, Math.min(1, s / 100)), Math.max(0, Math.min(1, l / 100)));
  }
  // A named CSS colour. Own-property only: a series coloured `__proto__` or
  // `constructor` would otherwise reach Object.prototype and come back an object
  // or a FUNCTION, which is how the same table broke the pptx sink and the SVG
  // pattern table before it.
  const named = Object.prototype.hasOwnProperty.call(CSS_NAMES, c.toLowerCase()) ? CSS_NAMES[c.toLowerCase()] : "";
  if (/^[0-9a-fA-F]{6}$/.test(named)) {
    const n = parseInt(named, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  return UNREADABLE;
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
  const c = paintText(color);
  // The `transparent` KEYWORD, which the sibling in `skill/scripts/pptx-paint.mjs`
  // has always known and this one did not. It is the documented floating-segment
  // idiom, and only two layouts guarded it before it reached a renderer — mekko's
  // guard even says why: "Office.js hands 'transparent' to setSolidColor, which
  // it rejects". The guard was never swept to the SINK, so on every other kind
  // the bare word went to the live host.
  if (/^transparent$/i.test(c)) return 0;
  // 4- and 8-digit hex carry the alpha in their last digit / last byte.
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

/**
 * Three channels to `#RRGGBB`, always — six hex digits, whatever it is handed.
 *
 * The clamp used to be `Math.max(0, Math.min(255, Math.round(c)))`, which is
 * total for numbers and not for NaN: every one of those returns NaN, and
 * `NaN.toString(16)` is the STRING "NaN", so a single unreadable channel
 * produced `#NaNNaNNaN`. That is not a colour in SVG, and in the pptx sink it
 * is not the six hex digits that path's security note requires.
 *
 * Guarded here rather than at the four call sites, because every producer of a
 * channel — `toRgb`, `lerpColor`'s interpolation, both colour scales — can
 * arrive with a NaN for its own reasons, and a fix at the root cannot be
 * forgotten by the next one.
 */
const rgbToHex = (rgb: number[]): string =>
  "#" +
  rgb
    .map((c) =>
      Math.max(0, Math.min(255, Number.isFinite(c) ? Math.round(c) : 0))
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

/** Fill for cells/tiles with no data, on a light canvas. */
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
 * "No data" fill for the canvas in play. A missing cell must read as ABSENT —
 * quieter than every value on the scale — and the light-canvas grey is instead
 * the brightest thing on a dark slide, which makes the gaps the loudest marks
 * in the grid. Light canvases keep the literal, so default charts are
 * byte-identical.
 *
 * Not zoneFill: that mirrors the tint's LINEAR-LIGHT step away from white into
 * an equal step away from black, which is accurate for the faint panel tints it
 * serves (#f5f5f5 and friends) but overshoots badly this far down — #e6e6e6
 * came back as #808080, a 4.3:1 block against a #1b1b1b slide where the light
 * pair is 1.25:1. Match the CONTRAST RATIO instead, so "absent" is equally
 * recessive on either canvas.
 */
export function noDataFill(background: string): string {
  const bg = relLuminance(background);
  if (bg >= 0.5) return NO_DATA;
  // The ratio NO_DATA keeps against a white canvas, re-solved against this one.
  const ratio = 1.05 / (relLuminance(NO_DATA) + 0.05);
  const target = Math.max(0, Math.min(1, ratio * (bg + 0.05) - 0.05));
  const c = Math.round(255 * (target <= 0.0031308 ? target * 12.92 : 1.055 * Math.pow(target, 1 / 2.4) - 0.055));
  return rgbToHex([c, c, c]);
}

/**
 * Sequential scale: 12% tint of the CANVAS (kept off the bare canvas so a low
 * value never reads as "no data") → the full color.
 *
 * The empty end is the background, not a hardcoded white: on a dark slide a
 * near-white low cell is the BRIGHTEST mark on the chart, so the scale reads
 * inverted — the smallest value shouts and the largest recedes. Defaulting to
 * white keeps every light-theme chart byte-identical.
 */
export function sequentialScale(
  min: number,
  max: number,
  color: string,
  background = "#ffffff",
): (v: number) => string {
  const lo = lerpColor(background, color, 0.12);
  const span = max - min || 1;
  return (v) => lerpColor(lo, color, (v - min) / span);
}

/**
 * Diverging scale through the canvas colour, symmetric around zero so equal
 * distances from zero get equal intensity on both sides. Zero must vanish into
 * the slide — a hardcoded white zero cell on a dark canvas is a glaring block
 * exactly where the data says "neutral".
 */
export function divergingScale(
  min: number,
  max: number,
  positive: string,
  negative: string,
  background = "#ffffff",
): (v: number) => string {
  const extent = Math.max(Math.abs(min), Math.abs(max)) || 1;
  return (v) => (v >= 0 ? lerpColor(background, positive, v / extent) : lerpColor(background, negative, -v / extent));
}
