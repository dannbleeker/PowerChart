#!/usr/bin/env node
/**
 * Render PowerChart configs to a .pptx with NATIVE, editable shapes —
 * the same output philosophy as the live add-in, headless.
 *
 *   node scripts/render-pptx.mjs charts.json out.pptx
 *
 * charts.json holds one ChartConfig or an array (one chart per slide).
 * Requires: npm install pptxgenjs
 */
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import pptxgen from "pptxgenjs";
// Pure paint + node mapping (split out so it is unit-testable and measurable —
// this CLI runs as a subprocess and can't be). Shipped alongside by build-skill.
import { IN, hex, makeAddNode } from "./pptx-paint.mjs";

// Engine location: packaged skill layout first, then repo layout.
let engine;
const failures = [];
for (const candidate of ["../lib/powerchart.js", "../../dist-lib/powerchart.js"]) {
  const href = new URL(candidate, import.meta.url).href;
  try {
    engine = await import(href);
    break;
  } catch (err) {
    // THIS candidate being absent just means "not this layout" — keep looking.
    // Anything else is the engine itself failing, and swallowing it reported a
    // corrupt-but-present lib as missing, telling the user to rebuild a file
    // that is already there (and the packaged skill has no build:lib to run).
    // Node names the missing file by path, not by href — a missing import
    // *inside* the engine names a different one, and that is a real failure.
    const missingSelf =
      err?.code === "ERR_MODULE_NOT_FOUND" &&
      String(err?.message ?? "").includes(fileURLToPath(href));
    if (!missingSelf) failures.push(`${candidate}: ${err?.message ?? err}`);
  }
}
if (!engine) {
  console.error(
    failures.length
      ? `powerchart engine failed to load:\n  ${failures.join("\n  ")}`
      : "powerchart engine not found — run `npm run build:lib` first",
  );
  process.exit(1);
}
const { buildChart, buildAgendaScene, DEFAULT_SIZE, arrowheadBox, annularSectorPoints, SYMBOL_PRESET, dashKind } =
  engine;

// A stale packaged lib (the skill ships no build step) can be missing an export,
// which otherwise blows up mid-render on the first chart that needs it. Fail
// fast with an actionable message instead.
for (const [name, fn] of Object.entries({ buildChart, buildAgendaScene, arrowheadBox, annularSectorPoints, dashKind })) {
  if (typeof fn !== "function") {
    console.error(`powerchart engine is missing export "${name}" — rebuild the lib (npm run build:lib)`);
    process.exit(1);
  }
}

const [, , input, output = "powerchart.pptx"] = process.argv;
if (!input) {
  console.error("usage: node scripts/render-pptx.mjs <charts.json> [out.pptx]");
  process.exit(1);
}

let raw;
try {
  raw = JSON.parse(readFileSync(input, "utf8"));
} catch (err) {
  console.error(`couldn't read/parse ${input}: ${err?.message ?? err}`);
  process.exit(1);
}
// Accept one config or an array of them. A wrapper object like {"charts":[…]} or
// a primitive is the common LLM mistake — name it rather than render one blank slide.
if (raw == null || typeof raw !== "object") {
  console.error("expected a ChartConfig object or an array of them (got a " + typeof raw + ")");
  process.exit(1);
}
if (!Array.isArray(raw) && !("kind" in raw) && !("data" in raw) && !("chapters" in raw)) {
  const wrapped = Object.keys(raw).find((k) => Array.isArray(raw[k]));
  console.error(
    wrapped
      ? `expected a ChartConfig or an array — did you mean the "${wrapped}" array inside this object?`
      : "expected a ChartConfig (with a kind/data) or an array of them",
  );
  process.exit(1);
}
const configs = Array.isArray(raw) ? raw : [raw];

const SLIDE = { w: 13.333, h: 7.5 };

const pres = new pptxgen();
pres.defineLayout({ name: "WIDE", width: SLIDE.w, height: SLIDE.h });
pres.layout = "WIDE";

// The engine helpers the node mapping needs, bound once into a pure addNode.
const addNode = makeAddNode({ dashKind, annularSectorPoints, SYMBOL_PRESET, arrowheadBox });

/** Build the scene for one config: an agenda slide, or a chart. */
function sceneFor(cfg) {
  if (cfg && cfg.kind === "agenda") {
    // { kind:"agenda", chapters:[…], highlight?, title? } → a chapter-list slide
    // (highlight the current chapter, or -1 for an overview). Full-slide size.
    return buildAgendaScene(Array.isArray(cfg.chapters) ? cfg.chapters : [], {
      highlight: cfg.highlight,
      title: cfg.title,
      width: cfg.width,
      height: cfg.height,
    });
  }
  return buildChart({ ...DEFAULT_SIZE, ...cfg });
}

/** A visible error slide so a bad config in a batch surfaces instead of vanishing. */
function errorSlide(i, err) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addText(`Chart ${i + 1} failed: ${err?.message ?? err}`, {
    x: 0.4,
    y: 0.4,
    w: SLIDE.w - 0.8,
    h: 0.6,
    fontSize: 18,
    bold: true,
    color: "C0392B",
    fill: { color: "FBEAE8" },
    align: "left",
    valign: "middle",
  });
}

let failed = 0;
configs.forEach((cfg, i) => {
  // Isolate each config: one bad chart in a 50-slide batch must not throw away
  // the other 49 (the Office.js path isolates per item too). Stamp an error
  // slide and carry on, so partial output always survives.
  try {
    const scene = sceneFor(cfg);
    const slide = pres.addSlide();
    // The chart's own canvas colour, not a fixed white: a dark-styled config
    // paints its ink for `style.background`, so a white slide under it put white
    // labels on white. Default (no style.background) stays FFFFFF.
    slide.background = { color: hex(cfg?.style?.background ?? "#ffffff") };
    const dx = (SLIDE.w - scene.width * IN) / 2;
    const dy = (SLIDE.h - scene.height * IN) / 2;
    for (const node of scene.nodes) addNode(slide, node, dx, dy);
  } catch (err) {
    failed++;
    errorSlide(i, err);
    console.error(`chart ${i + 1}: ${err?.message ?? err}`);
  }
});

await pres.writeFile({ fileName: output });
console.log(
  `${output}: ${configs.length} slide(s), native shapes` + (failed ? ` (${failed} failed — see error slides)` : ""),
);
if (failed) process.exit(1);
