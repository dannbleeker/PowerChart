/**
 * The style file, stored in the DECK rather than in the browser.
 *
 * WHY THIS EXISTS. The imported style file and the saved templates live in
 * `localStorage`, so they follow the BROWSER and not the presentation. Send a
 * branded deck to a colleague and their charts do not match yours — and every
 * chart they add drifts further, because the pane they are using has never seen
 * your palette. That is a sharing gap rather than a bug, which is why it sat in
 * the backlog: nothing is broken until the deck changes hands.
 *
 * A presentation-scoped custom XML part is the one thing this is genuinely the
 * right tool for. Chart CONFIG stays in shape tags and must: a config has to
 * travel with its shape, through copy/paste into another deck and through
 * PowerPoint's own Duplicate Slide, neither of which a presentation-scoped part
 * follows. A style is the opposite — it describes the deck, not any one shape.
 *
 * This module is the pure half: the XML the part carries, and which style wins
 * when the deck and the browser disagree. The host calls that read and write it
 * are `readDeckStyle` / `writeDeckStyle` in `src/render/powerpoint.ts`, so the
 * decisions here can be tested without a PowerPoint.
 */

/**
 * The namespace the part is filed under.
 *
 * A fully qualified schema URI, which is what `getByNamespace` requires — it is
 * an identifier, not a URL anyone fetches. Versioned in the path so a future
 * shape change can be a second namespace rather than a silent reinterpretation
 * of this one: a deck written by a newer build then reads as "no style I know"
 * on an older one, which is the safe direction.
 */
export const DECK_STYLE_NS = "https://powerchart.struktureretsundfornuft.dk/schemas/style/1.0";

/** The style defaults a deck carries. Every field optional — a partial style is normal. */
export interface DeckStyle {
  palette?: string[];
  fontFamily?: string;
  fontSize?: number;
  negative?: string;
  neutral?: string;
}

/** Which style wins when the deck carries one and this browser has one too. */
export type StylePreference = "deck" | "mine";

/**
 * Caps, so a deck cannot hand the pane something absurd.
 *
 * The XML comes out of a file somebody else made, which puts it in the same
 * class as the JSON box and a shape tag written in another deck: a `string` in
 * the types is not a string in the file someone pasted. These are the bounds a
 * style could plausibly want, not the bounds of what fits in memory.
 */
const MAX_PALETTE = 48;
const MAX_TEXT = 200;
const MIN_FONT = 1;
const MAX_FONT = 400;

/** A string, or undefined — never a number, an object, or something enormous. */
function textOr(v: unknown): string | undefined {
  if (typeof v !== "string") return undefined;
  const t = v.trim();
  return t && t.length <= MAX_TEXT ? t : undefined;
}

/**
 * Everything this module hands back, coerced and bounded.
 *
 * Built on a NULL-PROTOTYPE object and copied key by key from a fixed list, so
 * a deck carrying `{"__proto__": …}` or `{"constructor": …}` cannot reach
 * `Object.prototype` through it. `JSON.parse` creates `__proto__` as an own
 * property rather than by assignment, so the danger is not the parse — it is
 * everything downstream that reads the result with a plain lookup.
 */
export function normalizeDeckStyle(raw: unknown): DeckStyle {
  const out = Object.create(null) as DeckStyle;
  if (!raw || typeof raw !== "object") return out;
  const r = raw as Record<string, unknown>;
  const palette = Array.isArray(r.palette)
    ? r.palette
        .map(textOr)
        .filter((c): c is string => !!c)
        .slice(0, MAX_PALETTE)
    : [];
  if (palette.length) out.palette = palette;
  const family = textOr(r.fontFamily);
  if (family) out.fontFamily = family;
  const size = typeof r.fontSize === "number" && Number.isFinite(r.fontSize) ? r.fontSize : undefined;
  if (size != null && size >= MIN_FONT && size <= MAX_FONT) out.fontSize = size;
  const negative = textOr(r.negative);
  if (negative) out.negative = negative;
  const neutral = textOr(r.neutral);
  if (neutral) out.neutral = neutral;
  return out;
}

/** True when the style carries nothing worth storing. */
export function isEmptyStyle(style: DeckStyle | null | undefined): boolean {
  return !style || Object.keys(style).length === 0;
}

const ESCAPES: Record<string, string> = { "&": "&amp;", "<": "&lt;", ">": "&gt;" };

/**
 * The part's XML: one element, the style as JSON inside it.
 *
 * JSON rather than an element per field, because the style file the pane
 * exports IS this JSON — the same text the user pastes into the style box —
 * and a second spelling of it would be a second thing to keep in step.
 *
 * Escaped rather than wrapped in CDATA. A CDATA section cannot carry the
 * sequence `]]>`, and a palette entry or a font family is a string a user
 * chooses, so the one form that needs no exception is the one that escapes.
 */
export function styleToXml(style: DeckStyle): string {
  const json = JSON.stringify(normalizeDeckStyle(style));
  const escaped = json.replace(/[&<>]/g, (c) => ESCAPES[c]);
  return `<powerchartStyle xmlns="${DECK_STYLE_NS}" version="1">${escaped}</powerchartStyle>`;
}

/**
 * The style a part carries, or null when it carries none this build understands.
 *
 * Deliberately forgiving about the XML around the payload and strict about the
 * payload itself: a host may hand the part back with its own whitespace, an
 * attribute order of its choosing, or a self-closing element, and none of that
 * changes what the style says. What it may NOT do is arrive without the
 * namespace this module wrote — that would be somebody else's part.
 *
 * Parsed with a regex rather than `DOMParser`, which does not exist in the
 * headless renderer or in the test runner, and whose XML mode differs between
 * the two browsers this add-in runs in. The document is one element written by
 * this file; the general case is not the case.
 */
export function styleFromXml(xml: string | null | undefined): DeckStyle | null {
  if (typeof xml !== "string" || !xml.includes(DECK_STYLE_NS)) return null;
  const body = /<powerchartStyle\b[^>]*>([\s\S]*?)<\/powerchartStyle>/.exec(xml);
  if (!body) return null;
  const text = body[1]
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    // LAST, or an escaped `&lt;` written as `&amp;lt;` would come back as a
    // real `<` — unescaping is the reverse of escaping, and the ampersand is
    // the one that has to be undone at the end.
    .replace(/&amp;/g, "&");
  try {
    const style = normalizeDeckStyle(JSON.parse(text));
    return isEmptyStyle(style) ? null : style;
  } catch {
    // A part in our namespace whose payload will not parse. Null rather than an
    // empty style: "this deck has no style I can read" and "this deck's style
    // is empty" are the same for the caller, and neither should overwrite what
    // the user has in front of them.
    return null;
  }
}

/**
 * Which style the pane draws with, and why.
 *
 * THE DECK WINS BY DEFAULT, and this is the decision the backlog entry said had
 * to be taken before any code. The feature exists so a deck keeps its branding
 * when it changes hands; a colleague whose browser holds their own imported
 * style is exactly the case it was written for, and letting the browser win
 * would leave that colleague's charts mismatched — the original complaint,
 * unfixed. It is also the way think-cell's own style files work: the deck's
 * style is the deck's.
 *
 * The user is never LOCKED to it. `prefer: "mine"` is the pane's own override,
 * and the pane says which is in force rather than quietly applying one — a
 * palette changing under someone because they opened a different file is the
 * kind of surprise this project has a rule about.
 *
 * Not a merge, deliberately. Merging a deck's palette with a browser's font
 * would produce a style neither party chose and that no one could reproduce by
 * looking at either source.
 */
export function resolveStyleFile(
  deck: DeckStyle | null | undefined,
  mine: DeckStyle | null | undefined,
  prefer: StylePreference = "deck",
): { style: DeckStyle; from: "deck" | "mine" | "none" } {
  const deckHas = !isEmptyStyle(deck);
  const mineHas = !isEmptyStyle(mine);
  if (prefer === "deck" && deckHas) return { style: deck as DeckStyle, from: "deck" };
  if (mineHas) return { style: mine as DeckStyle, from: "mine" };
  if (deckHas) return { style: deck as DeckStyle, from: "deck" };
  return { style: Object.create(null) as DeckStyle, from: "none" };
}
