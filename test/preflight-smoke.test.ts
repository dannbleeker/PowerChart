import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { CHART_KINDS, sampleConfig } from "../src/core/samples";
import { buildChart } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";
import { arrowheadBox, dashKind, annularSectorPoints, SYMBOL_PRESET } from "../src/core/geometry";
import { makeAddNode } from "../skill/scripts/pptx-paint.mjs";
import type { ChartConfig } from "../src/core/types";
import type { Scene } from "../src/core/scene";

/**
 * Pre-flight smoke tripwire. Before validating in a real PowerPoint host — where
 * every issue costs a manual round-trip — sweep EVERY chart kind and every
 * showcase config through all three rendering paths that share the scene graph:
 * the pure engine, the SVG preview, and the headless-pptx node mapping (the same
 * native-shape philosophy the live add-in uses). It asserts the cheap invariants
 * a broken layout violates first: no throw, non-empty output, and finite geometry
 * (a single NaN coordinate is what puts a shape off-slide in PowerPoint). It does
 * not replace host testing — it catches the config/layout/geometry classes of bug
 * before you hit them by hand.
 */

/** Walk a scene and assert every coordinate is a finite number — no NaN/Infinity. */
function expectFiniteGeometry(scene: Scene, label: string) {
  const bad: string[] = [];
  const checkNum = (v: unknown, where: string) => {
    if (typeof v === "number" && !Number.isFinite(v)) bad.push(where);
  };
  for (const n of scene.nodes) {
    for (const [k, v] of Object.entries(n)) {
      checkNum(v, `${n.kind}.${k}`);
      if (Array.isArray(v)) {
        for (const p of v as unknown[]) {
          if (p && typeof p === "object")
            for (const [pk, pv] of Object.entries(p)) checkNum(pv, `${n.kind}.${k}.${pk}`);
        }
      }
    }
  }
  expect(bad, `${label}: non-finite coordinates on ${bad.join(", ")}`).toEqual([]);
}

/** A slide sink that records the pptx calls a node mapping makes. */
function recorder() {
  const shapes: { type: string; opts: Record<string, unknown> }[] = [];
  return {
    shapes,
    addShape(type: string, opts: Record<string, unknown>) {
      shapes.push({ type, opts });
    },
    addText(_t: string, opts: Record<string, unknown>) {
      shapes.push({ type: "text", opts });
    },
  };
}
const addNode = makeAddNode({ dashKind, annularSectorPoints, SYMBOL_PRESET, arrowheadBox });

/** Render one config through all three paths and assert the smoke invariants. */
function smoke(cfg: ChartConfig, label: string) {
  // 1. Engine — builds, produces marks, all coordinates finite.
  let scene!: Scene;
  expect(() => (scene = buildChart(cfg)), `${label}: buildChart threw`).not.toThrow();
  expect(scene.nodes.length, `${label}: rendered an empty scene`).toBeGreaterThan(0);
  expect(scene.width, `${label}: non-positive width`).toBeGreaterThan(0);
  expect(scene.height, `${label}: non-positive height`).toBeGreaterThan(0);
  expectFiniteGeometry(scene, label);

  // 2. SVG preview — non-empty document, and no NaN leaked into an attribute.
  let svg!: string;
  expect(() => (svg = sceneToSvg(scene)), `${label}: sceneToSvg threw`).not.toThrow();
  expect(svg.startsWith("<svg"), `${label}: not an SVG document`).toBe(true);
  expect(svg.includes("NaN"), `${label}: NaN in SVG output`).toBe(false);

  // 3. Headless pptx — every scene node maps to native shapes without throwing.
  const slide = recorder();
  expect(() => {
    for (const n of scene.nodes) addNode(slide, n, 0, 0);
  }, `${label}: pptx addNode threw`).not.toThrow();
  for (const s of slide.shapes) {
    for (const v of Object.values(s.opts)) {
      if (typeof v === "number") expect(Number.isFinite(v), `${label}: NaN in pptx ${s.type}`).toBe(true);
    }
  }
}

describe("preflight — every chart kind renders through every path", () => {
  for (const { kind } of CHART_KINDS) {
    it(`${kind}: engine + SVG + pptx render cleanly from its sample`, () => {
      smoke(sampleConfig(kind), `sample:${kind}`);
    });
  }

  it("covers all 25 documented kinds (no kind is silently unrendered)", () => {
    expect(CHART_KINDS.length).toBeGreaterThanOrEqual(25);
  });
});

// The showcase deck is large (121 real configs); drive it data-driven so a single
// failing config names itself rather than failing the whole file.
describe("preflight — every showcase config renders through every path", () => {
  const showcase = JSON.parse(readFileSync("examples/showcase.json", "utf8")) as ChartConfig[];

  it("loads the showcase deck", () => {
    expect(showcase.length).toBeGreaterThan(100);
  });

  it.each(showcase.map((cfg, i) => [i, cfg.kind, cfg] as const))(
    "showcase[%i] (%s) renders cleanly",
    (i, _kind, cfg) => {
      smoke(cfg, `showcase[${i}]:${cfg.kind}`);
    },
  );
});
