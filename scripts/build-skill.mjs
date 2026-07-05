#!/usr/bin/env node
/**
 * Assemble the uploadable PowerChart Agent Skill:
 *   npm run build:lib && node scripts/build-skill.mjs
 * → skill-dist/powerchart-charts.zip  (upload at claude.ai → Customize → Skills;
 *   it then also appears inside Claude for PowerPoint)
 */
import { cpSync, mkdirSync, rmSync, writeFileSync, existsSync } from "node:fs";
import { execSync } from "node:child_process";

if (!existsSync("dist-lib/powerchart.js")) {
  console.error("run `npm run build:lib` first");
  process.exit(1);
}

const root = "skill-dist/powerchart-charts";
rmSync("skill-dist", { recursive: true, force: true });
mkdirSync(`${root}/scripts`, { recursive: true });
mkdirSync(`${root}/lib`, { recursive: true });

cpSync("skill/SKILL.md", `${root}/SKILL.md`);
cpSync("skill/reference.md", `${root}/reference.md`);
cpSync("skill/scripts/render-pptx.mjs", `${root}/scripts/render-pptx.mjs`);
cpSync("scripts/render-batch.mjs", `${root}/scripts/render-svg.mjs`);
cpSync("dist-lib/powerchart.js", `${root}/lib/powerchart.js`);
cpSync("examples/charts.json", `${root}/examples/charts.json`);
writeFileSync(
  `${root}/package.json`,
  JSON.stringify({ name: "powerchart-charts-skill", private: true, type: "module", dependencies: { pptxgenjs: "^4.0.1" } }, null, 2),
);

// render-svg.mjs was written for the repo layout — point it at the bundled lib.
execSync(`node -e "
const fs = require('node:fs');
const p = '${root}/scripts/render-svg.mjs';
fs.writeFileSync(p, fs.readFileSync(p, 'utf8').replace('../dist-lib/powerchart.js', '../lib/powerchart.js'));
"`);

execSync(`cd skill-dist && python3 -m zipfile -c powerchart-charts.zip powerchart-charts`);
console.log("skill-dist/powerchart-charts.zip");
