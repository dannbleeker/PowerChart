#!/usr/bin/env node
/**
 * Render every sample chart kind to SVG files (docs/gallery-svg by default).
 * Requires `npm run build:lib` first.
 */
import { writeFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { buildChart, sceneToSvg, sampleConfig, CHART_KINDS } from "../dist-lib/powerchart.js";

const outDir = process.argv[2] ?? "docs/gallery-svg";
mkdirSync(outDir, { recursive: true });
for (const { kind } of CHART_KINDS) {
  const svg = sceneToSvg(buildChart(sampleConfig(kind)), { background: "#ffffff" });
  const file = join(outDir, `${kind}.svg`);
  writeFileSync(file, svg);
  console.log(file);
}
