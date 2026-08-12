import { describe, expect, it } from "vitest";
import { sceneToSvg } from "../src/render/svg";
import { buildChart, describeChart } from "../src/core/chart";
import type { RectNode, Scene, TextNode } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";
import { sampleConfig } from "../src/core/samples";
import { symbolPreset, SYMBOL_PRESET } from "../src/core/geometry";
import { buildCheckbox, type CheckState } from "../src/core/elements";

/**
 * SVG paint values come verbatim from ChartConfig (series color, custom palette,
 * per-point colors) and are interpolated into paint attributes. A crafted colour
 * must not be able to break out of the attribute and inject executable nodes —
 * the SVG is assigned via innerHTML in the pane preview and can be saved as a
 * standalone document, both of which run injected `<image onerror>` / scripts.
 */

const XSS = '#000"><image href=x onerror=alert(document.domain)><rect fill="#000';

describe("svg renderer neutralizes injected colours", () => {
  it("does not emit an attribute breakout from a malicious fill", () => {
    const scene: Scene = {
      width: 100,
      height: 100,
      nodes: [{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: XSS, name: "cell" }],
    };
    const svg = sceneToSvg(scene);
    expect(svg).not.toContain("<image");
    expect(svg).not.toContain("onerror");
    // The breakout quote-plus-bracket sequence must not survive into the markup.
    expect(svg).not.toContain('"><');
  });

  it("sanitizes fill, stroke, and text colour alike", () => {
    const scene: Scene = {
      width: 100,
      height: 100,
      nodes: [
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#111", stroke: XSS, strokeWidth: 1, name: "r" },
        {
          kind: "text",
          x: 0,
          y: 0,
          w: 10,
          h: 10,
          text: "hi",
          fontSize: 10,
          color: XSS,
          align: "center",
          valign: "middle",
          name: "t",
        },
      ],
    };
    const svg = sceneToSvg(scene);
    expect(svg).not.toContain("onerror");
    expect(svg).not.toContain("<image");
  });

  it("blocks injection through a per-point series colour end to end", () => {
    const cfg: ChartConfig = {
      kind: "stacked",
      width: 200,
      height: 150,
      data: { categories: ["A"], series: [{ name: "X", values: [1], color: XSS }] },
    };
    const svg = sceneToSvg(buildChart(cfg));
    expect(svg).not.toContain("onerror");
    expect(svg).not.toContain("<image");
  });

  it("draws a CSS Color 4 slash-alpha colour rather than falling back to black", () => {
    // `rgb(70 130 180 / 0.5)` is the form MDN documents first, and the allow-list
    // had no `/` — so this paint failed the test and the preview drew BLACK,
    // while both PowerPoint renderers parsed it and drew steel blue. One colour,
    // two pictures, and only the wrong one on screen.
    const fill = (c: string) =>
      /fill="([^"]*)"/.exec(
        sceneToSvg({ width: 10, height: 10, nodes: [{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: c }] }),
      )?.[1];
    for (const c of ["rgb(70 130 180 / 0.5)", "rgba(70 130 180 / 50%)", "hsl(207 44% 49% / 0.5)"]) {
      expect(fill(c), `${JSON.stringify(c)} fell back instead of being drawn`).toBe(c);
    }
  });

  it("still refuses every breakout, with the slash allowed", () => {
    // The widened class is one character, and the check that it cannot widen
    // what ESCAPES: everything inside the parentheses is still digits, `.`,
    // `,`, whitespace, `%` and `/` — no quote, no `<`/`>`, no `&`, and no `*`,
    // so not even a CSS comment can be opened.
    const fill = (c: string) =>
      /fill="([^"]*)"/.exec(
        sceneToSvg({ width: 10, height: 10, nodes: [{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: c }] }),
      )?.[1];
    for (const attack of [
      'rgb(1,2,3)"><script>alert(1)</script><rect fill="rgb(1,2,3',
      'rgb(1,2,3)" onload="alert(1)',
      "rgba(1/*)*/,2,3)",
      "rgb(1,2,3)/**/",
      "url(javascript:alert(1))",
      "url(//evil.example/x)",
      "rgb(1 2 3 / 0.5)</rect><script>alert(1)</script>",
      "rgb(1,2,3);behavior:url(#x)",
      "/",
      "//",
    ]) {
      expect(fill(attack), `${JSON.stringify(attack)} survived the allow-list`).toBe("#000000");
    }
  });

  it("passes legitimate colour forms through unchanged (no valid-chart regression)", () => {
    const scene: Scene = {
      width: 100,
      height: 100,
      nodes: [
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#3b82f6", name: "a" },
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#abc", name: "b" },
        { kind: "ellipse", cx: 5, cy: 5, rx: 4, ry: 4, fill: "rgb(10, 20, 30)", name: "c" },
        { kind: "polygon", points: [{ x: 0, y: 0 }], fill: "steelblue", stroke: "#fff", strokeWidth: 1, name: "d" },
      ],
    };
    const svg = sceneToSvg(scene);
    expect(svg).toContain('fill="#3b82f6"');
    expect(svg).toContain('fill="#abc"');
    expect(svg).toContain('fill="rgb(10, 20, 30)"');
    expect(svg).toContain('fill="steelblue"');
    expect(svg).toContain('stroke="#fff"');
  });

  it("keeps a pattern-tile fill and its generated id in sync after sanitizing", () => {
    const scene: Scene = {
      width: 100,
      height: 100,
      nodes: [{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: XSS, pattern: "diagonal", name: "p" }],
    };
    const svg = sceneToSvg(scene);
    expect(svg).not.toContain("onerror");
    // The rect references a pattern url whose id matches a defined <pattern>.
    const ref = svg.match(/fill="url\(#([\w.-]+)\)"/);
    expect(ref).toBeTruthy();
    expect(svg).toContain(`<pattern id="${ref![1]}"`);
  });

  it("builds a resolvable pattern id from a legitimate rgb() fill", () => {
    // rgb()/hsl() colours are valid (PAINT_OK accepts them) and reach the scene
    // from Series.color/colors, but their "(", ")" and "," used to leak into the
    // <pattern id> and the url(#…) reference — where the URL parser stops at the
    // first ")", leaving the fill pointing at a non-existent, unbalanced id.
    const scene: Scene = {
      width: 100,
      height: 100,
      nodes: [{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "rgb(78,121,167)", pattern: "diagonal", name: "p" }],
    };
    const svg = sceneToSvg(scene);
    const ref = svg.match(/fill="url\(#([^"]+)\)"/);
    expect(ref).toBeTruthy();
    // The captured id must be a single safe token — no stray parens/commas that
    // would truncate the reference — and must match a real <pattern> def.
    expect(ref![1]).toMatch(/^[\w.-]+$/);
    expect(svg).toContain(`<pattern id="${ref![1]}"`);
    // The tile keeps the actual colour so the hatch renders over the right hue.
    expect(svg).toContain('fill="rgb(78,121,167)"');
  });
});

/**
 * The colour allow-list above guards PAINT attributes. NUMERIC attributes are the
 * other half of the same surface: font-size / fill-opacity / stroke-width /
 * stroke-dasharray are interpolated straight into the markup, and ChartConfig's
 * numeric fields are only `number` in TypeScript — erased at runtime. A config
 * from an untrusted source (a `#c=` share link, an imported JSON, a
 * POWERCHART_CONFIG shape tag authored in another deck) can put a STRING there.
 */
describe("svg renderer neutralizes injected numerics", () => {
  const BREAKOUT = '10"><image href=x onerror=alert(1) /><text x="';

  it("does not let style.fontSize break out of the font-size attribute", () => {
    const cfg = {
      kind: "clustered",
      width: 480,
      height: 300,
      style: { fontSize: BREAKOUT },
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
    } as unknown as ChartConfig;
    const svg = sceneToSvg(buildChart(cfg));
    expect(svg).not.toContain("<image");
    expect(svg).not.toContain("onerror");
    // Falls back to a usable size rather than emitting the hostile string.
    expect(svg).toMatch(/font-size="\d+(\.\d+)?"/);
  });

  it("does not let decorations.fillOpacity break out of the fill-opacity attribute", () => {
    const cfg = {
      kind: "radar",
      width: 300,
      height: 300,
      decorations: { fillOpacity: '0.5" onmouseover="alert(1)' },
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [1, 2, 3] }] },
    } as unknown as ChartConfig;
    const svg = sceneToSvg(buildChart(cfg));
    expect(svg).not.toContain("onmouseover");
    expect(svg).not.toContain('"><');
  });

  it("coerces a hostile strokeWidth and dash array on a raw scene", () => {
    const scene = {
      width: 100,
      height: 100,
      nodes: [
        {
          kind: "line",
          x1: 0,
          y1: 0,
          x2: 10,
          y2: 10,
          stroke: "#111",
          strokeWidth: '1" onload="alert(1)',
          dash: ['2" onload="alert(1)'],
        },
        {
          kind: "rect",
          x: 0,
          y: 0,
          w: 10,
          h: 10,
          fill: "#111",
          stroke: "#222",
          strokeWidth: '1"><script>alert(1)</script><rect x="',
        },
      ],
    } as unknown as Scene;
    const svg = sceneToSvg(scene);
    expect(svg).not.toContain("onload");
    expect(svg).not.toContain("<script");
    expect(svg).not.toContain('"><');
  });
});

describe("XML-illegal characters (regression)", () => {
  // XML 1.0 forbids the C0 controls outright — they cannot be escaped, only
  // removed. One in a chart label produced an .svg no conforming parser would
  // open (and a .pptx PowerPoint calls corrupt) while the renderer reported
  // success. Reachable from any imported config: JSON preserves them happily.
  const VT = String.fromCharCode(11); // a Word line break
  const render = (title: string) =>
    sceneToSvg(
      buildChart({
        kind: "clustered",
        width: 480,
        height: 320,
        title,
        data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
      } as unknown as ChartConfig),
    );

  const illegal = (s: string) =>
    [...s].filter((ch) => {
      const c = ch.codePointAt(0)!;
      return c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d;
    });

  it("strips control characters from the document", () => {
    const svg = render(`Q1${VT}Report`);
    expect(illegal(svg), "illegal XML characters reached the SVG").toEqual([]);
  });

  it("keeps every LEGAL character — emoji, CJK, accents", () => {
    const svg = render("Growth \u{1F4C8} 売上 Ærø");
    // A naive surrogate-range strip would eat the emoji's two code units.
    expect(svg).toContain("\u{1F4C8}");
    expect(svg).toContain("売上");
    expect(svg).toContain("Ærø");
  });
});

/**
 * `PATTERN_TILE` and the marker-shape tables are plain object literals indexed
 * with a raw string out of ChartConfig. Inherited Object.prototype members made
 * the truthiness guard pass: `__proto__` yielded a non-callable object (the
 * whole render threw), `constructor`/`toString` yielded functions that were
 * CALLED as tile builders. The sibling pptx renderer already guards its colour
 * table this way.
 */
describe("prototype keys in paint and shape tables", () => {
  const patterned = (pattern: string): ChartConfig =>
    ({
      kind: "stacked",
      width: 480,
      height: 300,
      data: { categories: ["A"], series: [{ name: "S", values: [10], pattern }] },
    }) as unknown as ChartConfig;

  it.each(["__proto__", "constructor", "toString", "valueOf", "hasOwnProperty", "nonsense"])(
    "renders a chart whose pattern is %s without throwing",
    (pattern) => {
      const svg = sceneToSvg(buildChart(patterned(pattern)));
      // Falls back to the solid fill: no <pattern> def, and no dangling url(#…).
      expect(svg).not.toContain("<pattern");
      expect(svg).not.toContain("url(#");
    },
  );

  it("still emits a real pattern for a known name", () => {
    expect(sceneToSvg(buildChart(patterned("diagonal")))).toContain("<pattern");
  });

  it.each(["__proto__", "constructor", "star", ""])("renders a scatter whose marker is %s", (marker) => {
    const cfg = {
      kind: "scatter",
      width: 480,
      height: 300,
      scatter: { markers: [marker] },
      data: {
        categories: ["a", "b"],
        series: [
          { name: "x", values: [1, 2] },
          { name: "y", values: [3, 4] },
        ],
      },
    } as unknown as ChartConfig;
    const svg = sceneToSvg(buildChart(cfg));
    // Unknown shapes fall back to the circle, which is an <ellipse>.
    expect(svg).toContain("<ellipse");
    expect(svg).not.toContain('points=""');
  });
});

describe("prototype keys in the tables a config can index", () => {
  it("does not put Object.prototype into a chart's accessible description", () => {
    // `KIND_LABEL[cfg.kind]` was a bare lookup with a `?? "chart"` fallback,
    // and `??` does not fire for a function. `kind: "constructor"` therefore
    // opened the description with "function Object() { [native code] }" —
    // which is what a screen reader announces, and what the .pptx carries as
    // the shape's alt text. A config reaches this from the JSON import, from a
    // template, and from the skill.
    for (const kind of ["constructor", "toString", "valueOf", "hasOwnProperty", "__proto__"]) {
      const cfg = { ...sampleConfig("clustered"), kind } as unknown as ChartConfig;
      const text = describeChart(cfg);
      expect(text, `kind "${kind}" leaked a prototype member`).not.toMatch(/native code|\[object /);
      expect(text.startsWith("chart"), `kind "${kind}" did not fall back to "chart"`).toBe(true);
    }
  });

  it("draws a real preset for a symbol name off Object.prototype", () => {
    // Both PowerPoint renderers hand this straight to the host as a geometry
    // name. A function there resolves to undefined and draws a shape with no
    // geometry at all — invisible in the deck, and impossible to explain from
    // the file.
    for (const shape of ["constructor", "toString", "__proto__", "nope"]) {
      const preset = symbolPreset(shape);
      expect(typeof preset, `symbol "${shape}" did not yield a name`).toBe("string");
      expect(preset.length).toBeGreaterThan(0);
    }
    // …and a real one still resolves to itself.
    expect(symbolPreset("diamond")).toBe(SYMBOL_PRESET.diamond);
  });
});

/**
 * `buildCheckbox` keys two literal tables — glyph and colour — off its `state`
 * argument. `CheckState` is a union in the types and the function is exported
 * from `src/index.ts`, so the value is whatever a library caller passed. This is
 * the same class as the tables above and was the seventh instance of it in the
 * repo; the six CLAUDE.md listed did not include this file.
 */
describe("prototype keys in the checkbox element's tables", () => {
  const glyphOf = (state: string) =>
    (buildCheckbox(state as CheckState).nodes.find((n) => n.name === "check-glyph") as TextNode).text;
  const strokeOf = (state: string) =>
    (buildCheckbox(state as CheckState).nodes.find((n) => n.name === "check-box") as RectNode).stroke;

  it.each(["__proto__", "constructor", "toString", "valueOf", "hasOwnProperty", "nonsense"])(
    "falls back to the neutral mark for state %s",
    (state) => {
      // `__proto__` put Object.prototype into both the glyph and the stroke and
      // `constructor` put a FUNCTION there, and both reached the markup: the
      // glyph node's text became "[object Object]" / "function Object() { …" and
      // the stroke became a non-colour the renderers hand straight to the host.
      expect(glyphOf(state), `state "${state}" did not fall back`).toBe(glyphOf("partial"));
      expect(strokeOf(state), `state "${state}" did not fall back`).toBe(strokeOf("partial"));
      const svg = sceneToSvg(buildCheckbox(state as CheckState));
      expect(svg).not.toMatch(/native code|\[object /);
    },
  );

  it("still draws each real state's own glyph and colour", () => {
    // The rule must not be satisfiable by rejecting everything: the three real
    // states keep three distinct glyphs, and two of them three distinct strokes.
    const glyphs = (["yes", "no", "partial"] as const).map(glyphOf);
    expect(new Set(glyphs).size).toBe(3);
    expect(new Set((["yes", "no", "partial"] as const).map(strokeOf)).size).toBe(3);
  });
});
