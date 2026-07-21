#!/usr/bin/env node
/**
 * Batch automation (the .ppttc idea, open): render chart configs from a JSON
 * file to SVGs without PowerPoint.
 *
 *   npm run build:lib
 *   node scripts/render-batch.mjs examples/charts.json out/
 *
 * The JSON file holds one ChartConfig or an array of them (see README for
 * the schema; `sampleConfig` in the library produces valid examples).
 *
 * build-skill.mjs copies this file into the skill package as
 * scripts/render-svg.mjs, so every user-facing mention of the script name is
 * taken from argv — a usage line naming render-batch.mjs pointed the skill's
 * user at a file that does not exist in the zip.
 */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { basename, join } from "node:path";
import { buildChart, buildAgendaScene, sceneToSvg, DEFAULT_SIZE } from "../dist-lib/powerchart.js";

const [, self, input, outDir = "out"] = process.argv;
if (!input) {
  console.error(`usage: node scripts/${basename(self ?? "render-batch.mjs")} <charts.json> [outDir]`);
  process.exit(1);
}

let raw;
try {
  raw = JSON.parse(readFileSync(input, "utf8"));
} catch (err) {
  console.error(`Could not read ${input} as JSON: ${err.message}`);
  process.exit(1);
}
const configs = Array.isArray(raw) ? raw : [raw];
mkdirSync(outDir, { recursive: true });

/** Build the scene for one config: an agenda slide, or a chart. Mirrors render-pptx.mjs's sceneFor. */
function sceneFor(cfg) {
  if (cfg && cfg.kind === "agenda") {
    // { kind:"agenda", chapters:[…], highlight?, title? } — buildChart has no
    // such kind, so routing it there threw and took the whole batch with it.
    return buildAgendaScene(Array.isArray(cfg.chapters) ? cfg.chapters : [], {
      highlight: cfg.highlight,
      title: cfg.title,
      width: cfg.width,
      height: cfg.height,
    });
  }
  return buildChart({ ...DEFAULT_SIZE, ...cfg });
}

let failed = 0;
configs.forEach((cfg, i) => {
  // Isolate each config, as the pptx renderer does: one bad chart in a batch
  // must not throw away the previews of all the others.
  try {
    const svg = sceneToSvg(sceneFor(cfg), { background: "#ffffff" });
    const name = (cfg?.title ?? cfg?.kind ?? `chart-${i}`).replace(/[^\w-]+/g, "-").toLowerCase();
    const file = join(outDir, `${String(i + 1).padStart(2, "0")}-${name}.svg`);
    writeFileSync(file, svg);
    console.log(file);
  } catch (err) {
    failed++;
    console.error(`chart ${i + 1}: ${err?.message ?? err}`);
  }
});
if (failed) process.exit(1);
