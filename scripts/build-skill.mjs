#!/usr/bin/env node
/**
 * Assemble the uploadable PowerChart Agent Skill:
 *   npm run build:lib && node scripts/build-skill.mjs
 * → skill-dist/powerchart-charts.zip  (upload at claude.ai → Customize → Skills;
 *   it then also appears inside Claude for PowerPoint)
 */
import { cpSync, mkdirSync, readFileSync, rmSync, writeFileSync, existsSync } from "node:fs";
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
// render-pptx.mjs imports its paint/node helpers from this sibling — ship both,
// so the relative "./pptx-paint.mjs" import resolves in the packaged layout too.
cpSync("skill/scripts/pptx-paint.mjs", `${root}/scripts/pptx-paint.mjs`);
cpSync("scripts/render-batch.mjs", `${root}/scripts/render-svg.mjs`);
cpSync("dist-lib/powerchart.js", `${root}/lib/powerchart.js`);
cpSync("examples/charts.json", `${root}/examples/charts.json`);
writeFileSync(
  `${root}/package.json`,
  JSON.stringify(
    { name: "powerchart-charts-skill", private: true, type: "module", dependencies: { pptxgenjs: "^4.0.1" } },
    null,
    2,
  ),
);

// render-svg.mjs was written for the repo layout — point it at the bundled lib.
// Done in-process: shelling out to `node -e` to rewrite a file from a Node
// script bought nothing and broke on Windows, silently shipping a skill whose
// renderer imported ../dist-lib/powerchart.js, a path that doesn't exist here.
const renderSvg = `${root}/scripts/render-svg.mjs`;
const patched = readFileSync(renderSvg, "utf8").replace("../dist-lib/powerchart.js", "../lib/powerchart.js");
if (!patched.includes("../lib/powerchart.js")) {
  console.error(`${renderSvg}: import of ../dist-lib/powerchart.js not found — skill would ship broken`);
  process.exit(1);
}
writeFileSync(renderSvg, patched);

execSync(`cd skill-dist && python3 -m zipfile -c powerchart-charts.zip powerchart-charts`);
console.log("skill-dist/powerchart-charts.zip");
