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
 */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { buildChart, sceneToSvg, DEFAULT_SIZE } from "../dist-lib/powerchart.js";

const [, , input, outDir = "out"] = process.argv;
if (!input) {
  console.error("usage: node scripts/render-batch.mjs <charts.json> [outDir]");
  process.exit(1);
}

const raw = JSON.parse(readFileSync(input, "utf8"));
const configs = Array.isArray(raw) ? raw : [raw];
mkdirSync(outDir, { recursive: true });

configs.forEach((cfg, i) => {
  const full = { ...DEFAULT_SIZE, ...cfg };
  const svg = sceneToSvg(buildChart(full), { background: "#ffffff" });
  const name = (full.title ?? full.kind ?? `chart-${i}`).replace(/[^\w-]+/g, "-").toLowerCase();
  const file = join(outDir, `${String(i + 1).padStart(2, "0")}-${name}.svg`);
  writeFileSync(file, svg);
  console.log(file);
});
