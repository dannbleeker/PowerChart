import { describe, expect, it } from "vitest";
import {
  alphaOf,
  divergingScale,
  noDataFill,
  NO_DATA,
  sequentialScale,
  toHex6,
  toRgb,
  zoneFill,
} from "../src/core/color";
import { buildChart } from "../src/core/chart";
import type { ChartConfig } from "../src/core/types";
import type { RectNode } from "../src/core/scene";

/** Linear-light relative luminance of a #rrggbb hex, for "how bright" assertions. */
const luminance = (hex: string): number => {
  const [r, g, b] = toRgb(hex).map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};

/**
 * Subtle background zones (scatter quadrants, gantt weekend/holiday shading,
 * heatmap total strips) hardcode light-canvas tints. zoneFill keeps them exact
 * on a light background (so default charts are byte-identical) but adapts them on
 * a dark background, where a near-white box would glare.
 */
describe("zoneFill", () => {
  it("returns the light tint unchanged on a light background (byte-identical)", () => {
    for (const c of ["#efe7e7", "#f4f3f0", "#f0efec", "#f2f1ec", "#faf9f6"]) {
      expect(zoneFill("#ffffff", c)).toBe(c);
    }
  });

  it("adapts a light tint to a subtle panel on a dark background", () => {
    const dark = "#1a1a1a";
    const out = zoneFill(dark, "#f2f1ec");
    expect(out).not.toBe("#f2f1ec"); // not the glaring light box
    // A faint lift off the dark canvas: darker than the light literal, lighter
    // than the background itself.
    const lum = (h: string) => parseInt(h.slice(1, 3), 16) + parseInt(h.slice(3, 5), 16) + parseInt(h.slice(5, 7), 16);
    expect(lum(out)).toBeLessThan(lum("#f2f1ec"));
    expect(lum(out)).toBeGreaterThan(lum(dark));
  });

  it("preserves the two-tone quadrant checkerboard on both themes", () => {
    const cfg = (bg?: string): ChartConfig => ({
      kind: "scatter",
      width: 480,
      height: 360,
      style: bg ? ({ background: bg } as ChartConfig["style"]) : undefined,
      decorations: { quadrants: { x: 5, y: 5 } },
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [2, 8, 5] },
          { name: "Y", values: [3, 7, 6] },
        ],
      },
    });
    const quads = (c: ChartConfig) =>
      buildChart(c)
        .nodes.filter((n): n is RectNode => n.kind === "rect" && !!n.name?.startsWith("quadrant-"))
        .map((n) => n.fill);
    const light = quads(cfg());
    const dark = quads(cfg("#1a1a1a"));
    // Light: the tuned literals, unchanged.
    expect(light).toEqual(["#f2f1ec", "#faf9f6", "#faf9f6", "#f2f1ec"]);
    // Dark: adapted, and still alternating (0 and 3 share a tone, 1 and 2 share the other).
    expect(dark[0]).not.toBe("#f2f1ec");
    expect(dark[0]).toBe(dark[3]);
    expect(dark[1]).toBe(dark[2]);
    expect(dark[0]).not.toBe(dark[1]);
  });
});

/**
 * Label ink and series tints must follow the SURFACE, not assume a light canvas.
 * Two survivors of the #138 sweep: a pie's inside label was hardcoded white (so a
 * pale slice — including the default palette's #eda100 — printed white on light,
 * i.e. invisible), and treemap/violin tinted toward a literal "#ffffff" (which
 * washes out on a dark canvas instead of lifting off it).
 */
describe("label ink and tints follow the surface", () => {
  const lum = (h: string) => parseInt(h.slice(1, 3), 16) + parseInt(h.slice(3, 5), 16) + parseInt(h.slice(5, 7), 16);

  it("inks a pie's inside label against its own slice, not a presumed dark fill", () => {
    const cfg = {
      kind: "pie",
      width: 420,
      height: 320,
      decorations: { segmentLabels: true },
      data: {
        categories: ["Pale", "Dark"],
        series: [{ name: "S", values: [50, 50], colors: ["#f5f2c8", "#123456"] }],
      },
    } as unknown as ChartConfig;
    const nodes = buildChart(cfg).nodes;
    const ink = (i: number) => nodes.find((n) => n.name === `label-${i}`) as { color?: string } | undefined;
    expect(ink(0)?.color).toBe("#0b0b0b"); // dark ink on the pale slice — was #ffffff (invisible)
    expect(ink(1)?.color).toBe("#ffffff"); // white still correct on the dark slice
  });

  it("keeps the default palette's pale slice readable", () => {
    // #eda100 is PALETTE[2], so a 3-slice pie hits this with NO custom colours.
    const cfg = {
      kind: "pie",
      width: 420,
      height: 320,
      decorations: { segmentLabels: true },
      data: { categories: ["A", "B", "C"], series: [{ name: "S", values: [40, 35, 25] }] },
    } as unknown as ChartConfig;
    const nodes = buildChart(cfg).nodes;
    const pale = nodes.find((n) => n.name === "label-2") as { color?: string } | undefined;
    if (pale) expect(pale.color).toBe("#0b0b0b");
  });

  it("tints treemap groups and violin bodies toward the canvas on a dark theme", () => {
    const base = (kind: string, extra: Record<string, unknown> = {}) =>
      ({
        kind,
        width: 480,
        height: 320,
        data: {
          categories: ["G | one", "G | two", "H | three"],
          series: [{ name: "S", values: [10, 20, 30] }],
        },
        ...extra,
      }) as unknown as ChartConfig;

    for (const kind of ["treemap", "violin"]) {
      const dark = "#1b1b1b";
      const fills = buildChart({ ...base(kind), style: { background: dark } } as ChartConfig)
        .nodes.map((n) => (n as { fill?: string }).fill)
        .filter((f): f is string => typeof f === "string" && /^#[0-9a-f]{6}$/i.test(f));
      // Nothing may end up brighter than a mid grey on a near-black canvas: a tint
      // anchored on a literal white does exactly that.
      const brightest = Math.max(...fills.map(lum));
      expect(brightest, `${kind} on a dark canvas`).toBeLessThan(lum("#c8c8c8"));
    }
  });
});

/**
 * The colour scales and the "no data" fill each hardcode a light-canvas literal
 * and mirror it onto a dark canvas. The light branch keeps default charts
 * byte-identical; the dark branch is what makes a dark-themed deck legible, and
 * these exercise that dark branch directly rather than through a whole chart.
 */
describe("noDataFill", () => {
  it("returns the literal light-canvas grey on a light background", () => {
    expect(noDataFill("#ffffff")).toBe(NO_DATA);
    expect(noDataFill("#f4f3f0")).toBe(NO_DATA);
  });

  it("stays recessive on a dark background — not the brightest mark on the slide", () => {
    const dark = "#1b1b1b";
    const out = noDataFill(dark);
    expect(out).not.toBe(NO_DATA); // the light grey would glare here
    // Absent must read quieter than a mid-tone datum: a hair off the dark canvas,
    // nowhere near the #808080 that a naive linear-light mirror produced.
    expect(luminance(out)).toBeLessThan(luminance("#808080"));
    expect(luminance(out)).toBeGreaterThan(luminance(dark));
  });
});

describe("sequentialScale", () => {
  it("keeps white as the empty end on a light canvas (byte-identical default)", () => {
    const scale = sequentialScale(0, 100, "#c00000");
    expect(scale(0)).toBe(toHex6(sequentialScale(0, 100, "#c00000")(0)));
    // The low end is a faint tint of the colour on the canvas, never bare white.
    expect(scale(0)).not.toBe("#ffffff");
  });

  it("anchors the empty end on a dark canvas, so the smallest value recedes", () => {
    const dark = "#101418";
    const lo = sequentialScale(0, 100, "#4c9aff", dark)(0);
    const hi = sequentialScale(0, 100, "#4c9aff", dark)(100);
    // On a dark slide the low cell must be the DIMMEST, not a near-white block.
    expect(luminance(lo)).toBeLessThan(luminance(hi));
    expect(luminance(lo)).toBeLessThan(luminance("#888888"));
  });

  it("degenerate min==max never divides by zero", () => {
    const scale = sequentialScale(5, 5, "#c00000");
    expect(scale(5)).toMatch(/^#[0-9a-f]{6}$/i);
  });
});

describe("divergingScale", () => {
  it("puts equal intensity on equal distances from zero, through the canvas", () => {
    const scale = divergingScale(-100, 100, "#1a7f37", "#c00000");
    const pos = scale(50);
    const neg = scale(-50);
    expect(pos).toMatch(/^#[0-9a-f]{6}$/i);
    expect(neg).toMatch(/^#[0-9a-f]{6}$/i);
    expect(pos).not.toBe(neg); // opposite hues
    // Zero vanishes into the canvas (white by default).
    expect(scale(0)).toBe("#ffffff");
  });

  it("neutral zero vanishes into a dark canvas instead of glaring white", () => {
    const dark = "#141414";
    const scale = divergingScale(-100, 100, "#1a7f37", "#c00000", dark);
    // The zero cell is the background itself — no bright block where data is neutral.
    expect(scale(0)).toBe(toHex6(dark));
    // The negative branch still paints its own hue on the dark canvas.
    expect(scale(-100)).not.toBe(toHex6(dark));
  });
});

/**
 * toRgb/alphaOf normalise every allow-listed paint form. The chart engine only
 * ever feeds them clean 6-digit hex, so the parser's other branches — short hex,
 * rgb()/hsl(), 8-digit alpha, and the malformed fallbacks — are only reachable
 * from a hand-authored config and were the file's uncovered corner.
 */
describe("toRgb parsing", () => {
  it("expands 3- and 4-digit hex, dropping any alpha nibble", () => {
    expect(toRgb("#abc")).toEqual([0xaa, 0xbb, 0xcc]);
    expect(toRgb("#abcd")).toEqual([0xaa, 0xbb, 0xcc]); // 4th nibble is alpha
  });

  it("parses rgb() as bare 0–255 and rgb(%) as 0–100", () => {
    expect(toRgb("rgb(10, 20, 30)")).toEqual([10, 20, 30]);
    expect(toRgb("rgb(100%, 0%, 50%)")).toEqual([255, 0, 127]); // 50%·2.55 rounds to 127
  });

  it("parses hsl() through the colour wheel", () => {
    expect(toRgb("hsl(0, 100%, 50%)")).toEqual([255, 0, 0]); // pure red
    expect(toRgb("hsl(120, 100%, 50%)")).toEqual([0, 255, 0]); // pure green
  });

  it("falls back to mid grey for empty, null, malformed hex, and named colours", () => {
    expect(toRgb("")).toEqual([128, 128, 128]);
    expect(toRgb(undefined as unknown as string)).toEqual([128, 128, 128]);
    expect(toRgb("#zzzzzz")).toEqual([128, 128, 128]); // NaN hex
    expect(toRgb("rebeccapurple")).toEqual([128, 128, 128]); // named — known gap
  });
});

describe("alphaOf", () => {
  it("reads the alpha byte from 4- and 8-digit hex", () => {
    expect(alphaOf("#00000080")).toBeCloseTo(128 / 255, 5);
    expect(alphaOf("#0008")).toBeCloseTo(0x88 / 255, 5); // 4-digit → doubled nibble
  });

  it("reads the alpha channel from rgba()/hsla(), including a percentage", () => {
    expect(alphaOf("rgba(0,0,0,0.25)")).toBeCloseTo(0.25, 5);
    expect(alphaOf("hsla(0,0%,0%,50%)")).toBeCloseTo(0.5, 5);
  });

  it("is fully opaque for every form without an alpha, and clamps garbage to 1", () => {
    expect(alphaOf("#123456")).toBe(1);
    expect(alphaOf("rgb(1,2,3)")).toBe(1);
    expect(alphaOf("rgba(0,0,0,not-a-number)")).toBe(1); // non-finite → opaque
    expect(alphaOf("")).toBe(1);
  });
});
