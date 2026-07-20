import { describe, expect, it } from "vitest";
import { sceneToSvg } from "../src/render/svg";
import { buildChart } from "../src/core/chart";
import type { Scene } from "../src/core/scene";
import type { ChartConfig } from "../src/core/types";

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
