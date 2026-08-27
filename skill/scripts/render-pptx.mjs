#!/usr/bin/env node
/**
 * Render SSF Charts configs to a .pptx with NATIVE, editable shapes —
 * the same output philosophy as the live add-in, headless.
 *
 *   node scripts/render-pptx.mjs charts.json out.pptx
 *
 * charts.json holds one ChartConfig or an array (one chart per slide).
 * Requires: npm install pptxgenjs jszip
 */
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import pptxgen from "pptxgenjs";
// Pure paint + node mapping (split out so it is unit-testable and measurable —
// this CLI runs as a subprocess and can't be). Shipped alongside by build-skill.
import { IN, hexOr, makeAddNode } from "./pptx-paint.mjs";

// Engine location: packaged skill layout first, then repo layout.
let engine;
const failures = [];
for (const candidate of ["../lib/ssf-charts.js", "../../dist-lib/ssf-charts.js"]) {
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
      ? `SSF Charts engine failed to load:\n  ${failures.join("\n  ")}`
      : "SSF Charts engine not found — run `npm run build:lib` first",
  );
  process.exit(1);
}
const {
  buildChart,
  buildAgendaScene,
  DEFAULT_SIZE,
  arrowheadBox,
  annularSectorPoints,
  symbolPreset,
  dashKind,
  sceneToSvg,
  injectGroupsAndTags,
} = engine;

// A stale packaged lib (the skill ships no build step) can be missing an export,
// which otherwise blows up mid-render on the first chart that needs it. Fail
// fast with an actionable message instead.
for (const [name, fn] of Object.entries({
  buildChart,
  buildAgendaScene,
  arrowheadBox,
  annularSectorPoints,
  dashKind,
  sceneToSvg,
  injectGroupsAndTags,
})) {
  if (typeof fn !== "function") {
    console.error(`SSF Charts engine is missing export "${name}" — rebuild the lib (npm run build:lib)`);
    process.exit(1);
  }
}

/**
 * Rasterise a scene to PNG bytes. Lazy-loaded so a `render:"shapes"`-only run
 * never pays the @resvg/resvg-js require cost (native binding, ~50ms cold).
 * Returns the PNG buffer at 2× the scene's point size so it stays crisp on a
 * widescreen slide without ballooning the file.
 */
let RESVG;
async function rasterScene(scene, background) {
  if (!RESVG) {
    try {
      ({ Resvg: RESVG } = await import("@resvg/resvg-js"));
    } catch (err) {
      throw new Error(
        `image render mode needs @resvg/resvg-js installed (npm install @resvg/resvg-js): ${err?.message ?? err}`,
        { cause: err },
      );
    }
  }
  const svg = sceneToSvg(scene, background ? { background } : undefined);
  // 2× oversample: the scene is authored in points (72 dpi native), so a PNG at
  // point-size looks fuzzy on any zoomed slide. 2× = 144 dpi, matching what the
  // pane's PNG download uses via canvas.toDataURL.
  const png = new RESVG(svg, { fitTo: { mode: "width", value: Math.ceil(scene.width * 2) } }).render().asPng();
  return png;
}

const [, , input, output = "ssf-charts.pptx"] = process.argv;
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
// An empty array passes the shape guard above and then writes a deck PowerPoint
// REFUSES to open: with no slides, pptxgenjs emits a package whose
// [Content_Types].xml has no Override for the slide master, so the part falls
// back to the `xml` Default and the whole file is invalid — the repo's own
// `validate-ooxml` says "the document cannot be opened". The CLI printed
// "0 slide(s), native shapes" and exited 0. An agent whose data query came back
// empty handed the user an unopenable file having been told it worked, which is
// the same shape as Explode reporting a lost chart as a success.
if (configs.length === 0) {
  console.error("no chart configs in that file — nothing to render (an empty array writes an unopenable .pptx)");
  process.exit(1);
}

const SLIDE = { w: 13.333, h: 7.5 };

const pres = new pptxgen();
pres.defineLayout({ name: "WIDE", width: SLIDE.w, height: SLIDE.h });
pres.layout = "WIDE";

// The engine helpers the node mapping needs, bound once into a pure addNode.
const addNode = makeAddNode({ dashKind, annularSectorPoints, symbolPreset, arrowheadBox });

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

/**
 * A visible error slide so a bad config in a batch surfaces instead of vanishing.
 *
 * Takes the slide rather than adding one. The failing config may have already
 * had a slide added for it before it threw — see the catch below — and adding
 * a second here put the deck one slide ahead of `dressing`, which is indexed
 * strictly by position. From that point on every chart's POWERCHART_CONFIG
 * landed on the NEXT chart's slide: the error slide carried the following
 * chart's config, and the last chart in the deck carried none at all. The file
 * opens cleanly and looks right, so nothing surfaces it until someone edits a
 * chart and overwrites it with a different one's data.
 */
function errorSlide(slide, i, err) {
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
let imageSlides = 0;
// What to hang on each slide after pptxgenjs is done with it — see the
// injectGroupsAndTags call below. One entry per slide, in slide order.
const dressing = [];
for (let i = 0; i < configs.length; i++) {
  const cfg = configs[i];
  // Isolate each config: one bad chart in a 50-slide batch must not throw away
  // the other 49 (the Office.js path isolates per item too). Stamp an error
  // slide and carry on, so partial output always survives.
  // Held outside the try so the catch can tell "threw before a slide existed"
  // from "threw with a slide already added", and reuse it rather than add a
  // second one.
  let slide = null;
  try {
    const scene = sceneFor(cfg);
    slide = pres.addSlide();
    // The chart's own canvas colour, not a fixed white: a dark-styled config
    // paints its ink for `style.background`, so a white slide under it put white
    // labels on white. Default (no style.background) stays FFFFFF.
    // `hexOr`, not `hex`: an unrecognised paint must fall back to the caller's
    // WHITE, not to `hex`'s black. A typo like `background: "off-white"` — or
    // `"transparent"`, which SSF Charts documents as a paint elsewhere — turned
    // the slide black under the chart's own near-black ink, producing a deck of
    // invisible charts that opens cleanly and validates. The SVG renderer has
    // always fallen back to white here, so the two sinks disagreed.
    const bgHex = hexOr(cfg?.style?.background, "#ffffff");
    slide.background = { color: bgHex };
    if (cfg?.render === "image") {
      // Rasterise the whole scene into one picture object. Skips the per-node
      // shape flood — the escape hatch for dense charts (violin, area, sunburst)
      // whose native-shape counts trip the PowerPoint-web dense-shape wall.
      const png = await rasterScene(scene, `#${bgHex}`);
      const dataUri = `image/png;base64,${png.toString("base64")}`;
      // Match the shapes-path centering: same dx/dy as below, same width/height.
      const dx = (SLIDE.w - scene.width * IN) / 2;
      const dy = (SLIDE.h - scene.height * IN) / 2;
      // Named like every other chart object, which is not cosmetic. The repair
      // paths in `powerpoint.ts` find a chart by `name === "PowerChart"` —
      // `retagSlideChart`, `rescueGroupAndTag`, `slideHoldsOnlyChart` — so a
      // picture left with pptxgenjs's default `Image 0` carries its config tag
      // and is still unreachable by every one of them. The Office.js picture
      // path already names its shape this way and says why; the generated deck
      // did not, so the two renderers disagreed on the one string all of that
      // machinery keys on. Found by pointing `verify-deck` at the showcase.
      slide.addImage({
        data: dataUri,
        x: dx,
        y: dy,
        w: scene.width * IN,
        h: scene.height * IN,
        objectName: "PowerChart",
        // A picture is the one object pptxgenjs WILL carry alt text on, and an
        // image-mode chart is a single picture with no group to hang it from —
        // so this is the only route to a text alternative on those slides.
        altText: scene.desc,
      });
      imageSlides++;
    } else {
      const dx = (SLIDE.w - scene.width * IN) / 2;
      const dy = (SLIDE.h - scene.height * IN) / 2;
      for (const node of scene.nodes) addNode(slide, node, dx, dy);
    }
    // The chart's own config, so the add-in can re-open what this produced.
    // `scene.desc` is the chart's text alternative — the same string the SVG
    // renderer puts in `<desc>` and the add-in writes as the shape's alt text.
    // `injectGroupsAndTags` hangs it on the group it creates, which is the one
    // place in a generated deck that can carry one.
    dressing.push({
      configJson: JSON.stringify(cfg),
      title: cfg?.title ?? `Chart ${i + 1}`,
      desc: scene.desc,
    });
  } catch (err) {
    failed++;
    // Exactly one slide per config, whichever half of the try failed.
    errorSlide(slide ?? pres.addSlide(), i, err);
    console.error(`chart ${i + 1}: ${err?.message ?? err}`);
    // An error slide is not a chart: no group, nothing to re-edit.
    dressing.push({ group: false });
  }
}

// Group each chart and give it its POWERCHART_CONFIG tag.
//
// Without this the file is a pile of loose shapes with no identity: it looks
// right and the add-in cannot do a thing with it — no "Edit selected chart",
// no dragging a chart as one object. pptxgenjs can express neither (its only
// `grpSp` output is the mandatory empty boilerplate, and it has no concept of
// a tag part), so the bytes are post-processed by the SAME injector the add-in
// runs on its own generated decks. One implementation, so a chart Claude
// generates and a chart the pane drew are the same kind of object.
const { base64 } = await injectGroupsAndTags(await pres.write({ outputType: "base64" }), dressing);
writeFileSync(output, Buffer.from(base64, "base64"));
const modeSummary = imageSlides
  ? `${configs.length - imageSlides} native shapes + ${imageSlides} image`
  : "native shapes";
console.log(
  `${output}: ${configs.length} slide(s), ${modeSummary}` + (failed ? ` (${failed} failed — see error slides)` : ""),
);
if (failed) process.exit(1);
