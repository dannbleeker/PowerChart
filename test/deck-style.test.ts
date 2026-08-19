import { describe, expect, it } from "vitest";
import {
  DECK_STYLE_NS,
  isEmptyStyle,
  normalizeDeckStyle,
  resolveStyleFile,
  styleFromXml,
  styleToXml,
  type DeckStyle,
} from "../src/core/deck-style";

/**
 * The style a DECK carries: the XML it travels as, and which style wins when
 * the deck and the browser disagree.
 *
 * Both halves are pure so they can be argued about without a PowerPoint. The
 * host calls that store and fetch the part are exercised against the fake in
 * `office-render.test.ts`, and the pane's use of them in
 * `pane-host-actions.test.ts`.
 */
describe("the XML a deck's style travels as", () => {
  const style: DeckStyle = {
    palette: ["#2a78d6", "#1baf7a"],
    fontFamily: "Segoe UI",
    fontSize: 11,
    negative: "#e34948",
    neutral: "#898781",
  };

  it("round-trips a style through the part", () => {
    expect(styleFromXml(styleToXml(style))).toEqual(style);
  });

  it("carries the namespace the host files it under", () => {
    // `getByNamespace` is how both the reader and the writer find it, so the
    // namespace is not decoration — a part written without it is a part this
    // add-in can never see again.
    expect(styleToXml(style)).toContain(`xmlns="${DECK_STYLE_NS}"`);
  });

  it("escapes a style that would otherwise break the XML", () => {
    // A font family is a string a user types, and `<`, `>` and `&` are all
    // legal in one. Unescaped, the first of them ends the element early and the
    // part becomes unparseable — by us and by PowerPoint.
    const hostile: DeckStyle = { fontFamily: 'A & B <chart> "quoted"' };
    const xml = styleToXml(hostile);
    expect(xml).not.toContain("<chart>");
    expect(styleFromXml(xml)).toEqual(hostile);
  });

  it("unescapes an ampersand LAST, so an escaped escape survives", () => {
    // `&amp;lt;` is the escaping of the literal text `&lt;`. Undoing `&amp;`
    // first would leave `&lt;` for the next rule to turn into `<` — the text
    // would come back as a character the user never wrote.
    const literal: DeckStyle = { fontFamily: "&lt; not a tag" };
    expect(styleFromXml(styleToXml(literal))).toEqual(literal);
  });

  it("reads nothing out of another add-in's part", () => {
    expect(styleFromXml('<other xmlns="https://example.com/other">{"palette":["#fff"]}</other>')).toBeNull();
  });

  it("reads nothing out of a part of ours whose payload is not JSON", () => {
    // "this deck has no style I can read" and "this deck's style is empty" are
    // the same answer for the caller, and neither may overwrite what the user
    // has in front of them.
    expect(styleFromXml(`<powerchartStyle xmlns="${DECK_STYLE_NS}">not json</powerchartStyle>`)).toBeNull();
  });

  it("reads nothing at all rather than throwing on junk", () => {
    for (const junk of [null, undefined, "", "<", 42 as unknown as string]) {
      expect(styleFromXml(junk)).toBeNull();
    }
  });
});

describe("what a deck is allowed to say", () => {
  it("keeps only the fields a style has, coerced", () => {
    const raw = {
      palette: ["#2a78d6", 7, null, "  #1baf7a  "],
      fontFamily: 42,
      fontSize: "11",
      negative: "#e34948",
      extra: "ignored",
    };
    expect(normalizeDeckStyle(raw)).toEqual({ palette: ["#2a78d6", "#1baf7a"], negative: "#e34948" });
  });

  it("refuses a font size no chart could use", () => {
    expect(normalizeDeckStyle({ fontSize: 0 }).fontSize).toBeUndefined();
    expect(normalizeDeckStyle({ fontSize: 1e6 }).fontSize).toBeUndefined();
    expect(normalizeDeckStyle({ fontSize: Number.NaN }).fontSize).toBeUndefined();
    expect(normalizeDeckStyle({ fontSize: 11 }).fontSize).toBe(11);
  });

  it("caps a palette and a string, so a deck cannot hand the pane something absurd", () => {
    const huge = normalizeDeckStyle({ palette: Array.from({ length: 500 }, () => "#000000") });
    expect(huge.palette!.length).toBeLessThanOrEqual(48);
    expect(normalizeDeckStyle({ fontFamily: "x".repeat(5000) }).fontFamily).toBeUndefined();
  });

  it("cannot be used to reach Object.prototype", () => {
    // The XML comes out of a file somebody else made — the same class as the
    // JSON box and a shape tag written in another deck. `JSON.parse` makes
    // `__proto__` an OWN property, so the danger is every plain lookup after it.
    const style = styleFromXml(
      `<powerchartStyle xmlns="${DECK_STYLE_NS}">{"__proto__":{"polluted":1},"palette":["#fff"]}</powerchartStyle>`,
    );
    expect(style).toEqual({ palette: ["#fff"] });
    expect(Object.getPrototypeOf(style!)).toBeNull();
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  });

  it("knows an empty style from one worth storing", () => {
    expect(isEmptyStyle(null)).toBe(true);
    expect(isEmptyStyle({})).toBe(true);
    expect(isEmptyStyle({ fontSize: 11 })).toBe(false);
  });
});

describe("which style wins", () => {
  const deck: DeckStyle = { palette: ["#deck00"] };
  const mine: DeckStyle = { palette: ["#mine00"] };

  it("gives the deck the last word by default", () => {
    // THE decision this feature turned on. The point of storing a style in the
    // deck is that a colleague opening it sees the brand — and a colleague is
    // precisely the person whose browser holds a different style.
    expect(resolveStyleFile(deck, mine)).toEqual({ style: deck, from: "deck" });
  });

  it("hands it back to the browser when the pane asks", () => {
    expect(resolveStyleFile(deck, mine, "mine")).toEqual({ style: mine, from: "mine" });
  });

  it("falls through in both directions rather than answering nothing", () => {
    expect(resolveStyleFile(null, mine)).toEqual({ style: mine, from: "mine" });
    expect(resolveStyleFile(deck, null, "mine")).toEqual({ style: deck, from: "deck" });
  });

  it("says so when neither has anything", () => {
    const { style, from } = resolveStyleFile(null, {}, "mine");
    expect(from).toBe("none");
    expect(style).toEqual({});
  });

  it("never merges the two", () => {
    // A deck's palette under a browser's font is a style neither party chose
    // and that nobody can reproduce by looking at either source.
    const merged = resolveStyleFile({ palette: ["#deck00"] }, { fontFamily: "Comic Sans" });
    expect(merged.style.fontFamily).toBeUndefined();
  });
});
