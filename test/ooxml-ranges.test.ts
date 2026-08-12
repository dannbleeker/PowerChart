import { describe, expect, it } from "vitest";
import { buildChart } from "../src/core/chart";
import { sampleConfig, CHART_KINDS } from "../src/core/samples";
import { makeAddNode } from "../skill/scripts/pptx-paint.mjs";
import { dashKind, annularSectorPoints, symbolPreset, arrowheadBox } from "../src/core/geometry";
import type { ChartConfig } from "../src/core/types";

/**
 * Every number the pptx sink hands pptxgenjs lands in an OOXML attribute with a
 * SCHEMA RANGE, and a value outside it produces a deck PowerPoint offers to
 * repair — from a CLI that reports success. `finiteNodes` covers one corner of
 * this (`x="Infinity"` is not an Int64); the ranges are the rest of it, and
 * nothing checked them.
 *
 * The gap this was written for: `style.fontSize` was guarded for sign and
 * finiteness and not for range, so `fontSize: 1e6` wrote `sz="120000000"`
 * against a `ST_TextFontSize` maximum of 400000, and `fontSize: 0.0001` wrote
 * `sz="0"` against its minimum of 100. Both were fixed at the boundary and at
 * the sink; this is the check that would have caught them, generalised to every
 * numeric attribute rather than to the one that happened to be found.
 *
 * It goes through `makeAddNode` rather than through a written .pptx on purpose:
 * the mapping is where the units are chosen, it is a pure function, and this
 * runs the whole sample matrix in well under a second. A written-file check
 * would be a better oracle and could not be run 350 times.
 */

const addNode = makeAddNode({ dashKind, annularSectorPoints, symbolPreset, arrowheadBox });

/** OOXML limits, expressed in the units pptxgenjs takes (inches and points). */
const EMU_PER_INCH = 914400;
/** ST_PositiveCoordinate's ceiling — the largest offset or extent OOXML holds. */
const COORD_MAX_IN = 27273042316900 / EMU_PER_INCH;
/** ST_LineWidth is 0..20116800 EMU, and pptxgenjs takes a line width in points. */
const LINE_MAX_PT = 20116800 / 12700;
/** ST_TextFontSize is 100..400000 in hundredths of a point. */
const FONT_MIN_PT = 1;
const FONT_MAX_PT = 4000;

/** Every way one shape's options can fall outside what the format can carry. */
function faultsIn(opts: Record<string, unknown>, where: string): string[] {
  const out: string[] = [];
  const bad = (k: string, v: unknown, why: string) => out.push(`${where}.${k}=${String(v)} ${why}`);

  for (const k of ["x", "y", "w", "h"] as const) {
    const v = opts[k];
    if (v === undefined) continue;
    if (typeof v !== "number" || !Number.isFinite(v)) bad(k, v, "is not a finite number");
    else if (Math.abs(v) > COORD_MAX_IN) bad(k, v, "is past ST_Coordinate");
    else if ((k === "w" || k === "h") && v < 0) bad(k, v, "is a negative extent");
  }

  const fs = opts.fontSize;
  if (fs !== undefined && !(typeof fs === "number" && fs >= FONT_MIN_PT && fs <= FONT_MAX_PT))
    bad("fontSize", fs, `is outside ST_TextFontSize (${FONT_MIN_PT}..${FONT_MAX_PT}pt)`);

  const rot = opts.rotate;
  if (rot !== undefined && !(typeof rot === "number" && Number.isFinite(rot) && Math.abs(rot) <= 360))
    bad("rotate", rot, "is outside -360..360");

  for (const key of ["line", "fill"] as const) {
    const o = opts[key] as Record<string, unknown> | undefined;
    if (!o || typeof o !== "object") continue;
    const w = o.width;
    if (w !== undefined && !(typeof w === "number" && Number.isFinite(w) && w >= 0 && w <= LINE_MAX_PT))
      bad(`${key}.width`, w, "is outside ST_LineWidth");
    const t = o.transparency;
    if (t !== undefined && !(typeof t === "number" && Number.isFinite(t) && t >= 0 && t <= 100))
      bad(`${key}.transparency`, t, "is outside 0..100");
    // The six-hex-digit rule `hex` states about itself, checked from outside it.
    const c = o.color;
    if (c !== undefined && !/^[0-9a-fA-F]{6}$/.test(String(c))) bad(`${key}.color`, c, "is not six hex digits");
  }

  const pts = opts.points as { x?: number; y?: number }[] | undefined;
  if (Array.isArray(pts))
    for (const p of pts)
      for (const k of ["x", "y"] as const) {
        const v = p[k];
        if (v !== undefined && !(Number.isFinite(v) && Math.abs(v) <= COORD_MAX_IN))
          bad(`points.${k}`, v, "is not a usable coordinate");
      }

  return out;
}

/** Build a config, map every node through the sink, and collect what it wrote. */
function sweep(cfg: ChartConfig, where: string): string[] {
  const out: string[] = [];
  const record = (_kindOrText: string, o: Record<string, unknown>) => out.push(...faultsIn(o, where));
  const slide = { addShape: record, addText: (s: string, o: Record<string, unknown>) => record(s, o) };
  for (const n of buildChart(cfg).nodes) addNode(slide, n, 0, 0);
  return out;
}

describe("every number the pptx sink writes is one OOXML can carry", () => {
  it("across every chart kind, as sampled", () => {
    const faults: string[] = [];
    for (const { kind } of CHART_KINDS) faults.push(...sweep(sampleConfig(kind), kind));
    expect(faults, `${faults.length} attribute(s) outside the schema`).toEqual([]);
  });

  it("across every chart kind, with a hostile config", () => {
    // Each entry is a config a user or an agent can actually write. The two
    // font sizes are the pair that motivated this file; the rest cover the
    // other attributes the same way, so a range bug in any of them is caught
    // by the check rather than by a deck someone could not open.
    const HOSTILE: [string, Record<string, unknown>][] = [
      ["biggest allowed chart", { width: 7199, height: 7199 }],
      ["smallest chart", { width: 1, height: 1 }],
      ["huge font", { style: { fontSize: 1e6 } }],
      ["sub-point font", { style: { fontSize: 0.0001 } }],
      ["font at the ceiling", { style: { fontSize: 2000 } }],
      ["transparent palette", { style: { palette: ["transparent", "#00000000"] } }],
      ["transparent background", { style: { background: "transparent" } }],
      ["no data at all", { data: { categories: [], series: [] } }],
      ["a single zero", { data: { categories: ["a"], series: [{ name: "s", values: [0] }] } }],
      [
        "values at the float ceiling",
        { data: { categories: ["a", "b"], series: [{ name: "s", values: [1e308, -1e308] }] } },
      ],
      ["subnormal values", { data: { categories: ["a", "b"], series: [{ name: "s", values: [5e-324, 0] }] } }],
      [
        "every arrow decoration",
        { decorations: { cagrArrow: true, deltaArrow: true, valueLabels: true, totals: true } },
      ],
    ];
    const faults: string[] = [];
    for (const { kind } of CHART_KINDS)
      for (const [label, patch] of HOSTILE)
        faults.push(...sweep({ ...sampleConfig(kind), ...patch } as ChartConfig, `${kind} / ${label}`));
    // Reported in full rather than by count: the point of failing here is to
    // name the attribute and the config, which a bare number would not.
    expect(faults.slice(0, 20), `${faults.length} attribute(s) outside the schema`).toEqual([]);
  });
});
