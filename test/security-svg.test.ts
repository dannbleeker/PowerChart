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
});
