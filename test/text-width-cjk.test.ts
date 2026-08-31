import { describe, expect, it } from "vitest";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { clipTextToFrame, ellipsize, textWidth } from "../src/core/scene";
import { clipToWidth } from "../src/core/elements";
import type { ChartConfig } from "../src/core/types";
import type { TextNode } from "../src/core/scene";

/**
 * A FULL-WIDTH GLYPH TAKES A FULL EM, and the engine charged it 0.54.
 *
 * `textWidth` is the one question sixty-odd call sites ask about a title, a
 * category, a series name or a table cell — it decides what gets clipped, how
 * far a font shrinks, and whether two labels are judged to collide. It counted
 * code units and multiplied by a Latin average, so at 12pt:
 *
 *     "売上高"        measured 19.4pt   glyphs take 36.0   -46%
 *     "매출액"        measured 19.4pt   glyphs take 36.0   -46%
 *     "収益 Revenue"  measured 64.8pt   glyphs take 75.8   -15%
 *
 * Every Japanese, Chinese and Korean chart was therefore judged to fit where it
 * does not: labels not clipped that should have been, fonts not shrunk, overlaps
 * no fit could see. Nothing was wrong with the ARITHMETIC — it was measuring
 * something adjacent to what it meant to measure, which is this repo's most
 * repeated defect.
 */
describe("text is measured at the width its glyphs take", () => {
  const FS = 12;

  it("charges a full em for a full-width glyph", () => {
    // 3 ideographs at 12pt = 36pt, not 19.4.
    expect(textWidth("売上高", FS)).toBeCloseTo(36, 5);
    expect(textWidth("매출액", FS)).toBeCloseTo(36, 5);
    // Full-width brackets are full-width too — they are in the same block.
    expect(textWidth("（）", FS)).toBeCloseTo(24, 5);
  });

  it("leaves Latin EXACTLY where it was, so no shipped chart moves", () => {
    // The reason this could be changed at all: sixty-odd call sites and every
    // fit in the deck were measured against the old number. A string with no
    // full-width code point must come back bit-identical.
    for (const s of ["Revenue", "Q1", "", "Ünïcödé àccents", "1,234.56"]) {
      expect(textWidth(s, FS), s).toBeCloseTo(s.length * FS * 0.54, 10);
      expect(textWidth(s, FS, true), `${s} bold`).toBeCloseTo(s.length * FS * 0.58, 10);
    }
  });

  it("keeps half-width katakana narrow", () => {
    // U+FF61-FF9F is the half-width block — narrow is the entire point of it,
    // and a range that swallowed it would over-measure Japanese instead.
    expect(textWidth("ｱｲｳ", FS)).toBeCloseTo(3 * FS * 0.54, 5);
  });

  it("does not halve an emoji on the way past", () => {
    /**
     * The trap in this change. A surrogate PAIR used to be counted as two Latin
     * characters — 1.08em, about right for a pictograph by accident. Switching
     * to code points naively would have charged one narrow character, halving
     * every emoji and silently regressing something while fixing something else.
     */
    expect(textWidth("📊", FS)).toBeCloseTo(FS, 5);
    expect(textWidth("📊📈", FS)).toBeCloseTo(2 * FS, 5);
  });

  it("mixes the two without rounding either into the other", () => {
    // "収益" is 2 wide, " Revenue" is 8 narrow.
    expect(textWidth("収益 Revenue", FS)).toBeCloseTo((2 + 8 * 0.54) * FS, 5);
  });

  it("never clips a label into a lone surrogate", () => {
    /**
     * `clipToWidth` walked back one UTF-16 UNIT at a time, and an emoji is two.
     * A clip landing mid-pair left a lone high surrogate on the end of the
     * label — "Sales 📊\ud83d…" — which is not a rendering nuisance but INVALID
     * XML, written straight into a slide by the .pptx path. That is the
     * "PowerPoint found a problem with content" repair dialog, and a repair
     * takes the deck's tags and every chart's re-editability with it.
     *
     * Swept across widths because the defect is positional: it only appears
     * when the cut lands between the two halves of a pair.
     */
    const hasLoneSurrogate = (s: string) => {
      for (let i = 0; i < s.length; i++) {
        const c = s.charCodeAt(i);
        if (c >= 0xdc00 && c <= 0xdfff) return true; // a low half with no high
        if (c >= 0xd800 && c <= 0xdbff) {
          const n = s.charCodeAt(i + 1);
          if (!(n >= 0xdc00 && n <= 0xdfff)) return true;
          i++;
        }
      }
      return false;
    };
    const s = "Sales 📊📈📉 by region";
    for (let w = 4; w <= 120; w += 2) {
      const out = clipToWidth(s, 12, w, false);
      expect(hasLoneSurrogate(out), `maxW ${w} produced ${JSON.stringify(out)}`).toBe(false);
      // BOTH walks, because there were two. `clipToWidth` and the de-collision
      // pass in scene.ts had identical copies of this loop thirty lines apart,
      // and fixing one left the other cutting emoji in half — the same
      // "one call site was missed" shape this repo keeps finding, committed
      // once more while fixing an instance of it. They are one function now.
      const direct = ellipsize(s, 12, false, w);
      expect(hasLoneSurrogate(direct), `ellipsize at ${w} produced ${JSON.stringify(direct)}`).toBe(false);
      expect(direct, `the two walks disagree at ${w}`).toBe(out);
    }

    /**
     * AND THROUGH THE OTHER CALLER, which is the one that made this a lesson.
     * `clipTextToFrame` had its own copy of the walk thirty lines from the
     * shared one, so a test that only drove `clipToWidth` passed while the
     * frame clip went on cutting pairs in half. Asserting the helper twice is
     * not coverage; this drives the second CALL SITE.
     */
    for (let width = 20; width <= 140; width += 10) {
      const clipped = clipTextToFrame(
        [{ kind: "text", x: 2, y: 0, w: 400, h: 14, text: s, fontSize: 12, bold: false } as TextNode],
        width,
      );
      for (const n of clipped) {
        expect(
          hasLoneSurrogate(String((n as TextNode).text)),
          `clipTextToFrame at ${width} produced ${JSON.stringify((n as TextNode).text)}`,
        ).toBe(false);
      }
    }
  });

  it("clips a CJK title that does not fit, where it used to let it run", () => {
    /**
     * End to end: the measurement only matters because a fit depends on it. A
     * long Japanese title on a narrow chart must now be cut, and at the old
     * width it was judged to fit in roughly half the room it needs.
     */
    const long = "四半期別の売上高と粗利益の推移および前年同期比の分析";
    const scene = buildChart({
      kind: "clustered",
      ...DEFAULT_SIZE,
      width: 240,
      height: 160,
      title: long,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
    } as unknown as ChartConfig);
    const title = scene.nodes.find((n): n is TextNode => n.kind === "text" && n.name === "title");
    expect(title, "the chart drew no title at all").toBeTruthy();
    expect(
      textWidth(title!.text, title!.fontSize, title!.bold),
      "the title is drawn wider than the chart",
    ).toBeLessThanOrEqual(240);
  });
});
