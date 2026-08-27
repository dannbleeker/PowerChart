#!/usr/bin/env node
/**
 * Assemble the uploadable PowerChart Agent Skill:
 *   npm run build:lib && node scripts/build-skill.mjs
 * → skill-dist/ssf-charts.zip  (upload at claude.ai → Customize → Skills;
 *   it then also appears inside Claude for PowerPoint)
 */
import { cpSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync, existsSync } from "node:fs";
import { join, relative } from "node:path";
import JSZip from "jszip";

if (!existsSync("dist-lib/ssf-charts.js")) {
  console.error("run `npm run build:lib` first");
  process.exit(1);
}

const root = "skill-dist/ssf-charts";
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
cpSync("dist-lib/ssf-charts.js", `${root}/lib/ssf-charts.js`);
cpSync("examples/charts.json", `${root}/examples/charts.json`);
writeFileSync(
  `${root}/package.json`,
  JSON.stringify(
    {
      name: "ssf-charts-skill",
      private: true,
      type: "module",
      dependencies: {
        pptxgenjs: "^4.0.1",
      },
      // Rasteriser for render:"image" mode — and OPTIONAL now means optional.
      //
      // It used to sit in `dependencies` under a comment saying a shapes-only
      // install still works if npm cannot pull the native binary. The runtime
      // half of that is true, since the CLI lazy-imports it. The INSTALL half
      // was not: npm resolves a manifest atomically, so a registry that will not
      // serve a scoped native package — a mirrored corporate one, which is
      // exactly what this add-in's users have — aborts the whole install and
      // rolls back. No node_modules at all, pptxgenjs included, so the
      // shapes-only path SKILL.md says needs nothing extra was dead as well.
      //
      // `optionalDependencies` is the field that means what the comment claimed.
      optionalDependencies: {
        "@resvg/resvg-js": "^2.6.2",
      },
    },
    null,
    2,
  ),
);

// render-svg.mjs was written for the repo layout — point it at the bundled lib.
// Done in-process: shelling out to `node -e` to rewrite a file from a Node
// script bought nothing and broke on Windows, silently shipping a skill whose
// renderer imported ../dist-lib/ssf-charts.js, a path that doesn't exist here.
const renderSvg = `${root}/scripts/render-svg.mjs`;
const patched = readFileSync(renderSvg, "utf8").replace("../dist-lib/ssf-charts.js", "../lib/ssf-charts.js");
if (!patched.includes("../lib/ssf-charts.js")) {
  console.error(`${renderSvg}: import of ../dist-lib/ssf-charts.js not found — skill would ship broken`);
  process.exit(1);
}
writeFileSync(renderSvg, patched);

/**
 * Zip the package in-process, because the interpreter this used to borrow is
 * not always there.
 *
 * `python3 -m zipfile` is a fine zipper on a machine that has Python. On
 * Windows `python3` is the Microsoft Store's alias STUB — it prints "Python was
 * not found; run without arguments to install from the Microsoft Store" and
 * fails — so `npm run skill` could not complete on the owner's own box, and
 * `test/skill.test.ts`'s packaging check went with it. `jszip` is already a
 * dependency of this repo and reads the result back in that same test, so the
 * zipper and the reader are now the same library.
 *
 * Entries are added in sorted order and given a fixed date. A zip is an
 * artifact this project attaches to releases, and one that differs run to run
 * for no reason is one nobody can diff. `DeflateLevel` 9 keeps the upload small.
 */
function filesUnder(dir) {
  return readdirSync(dir, { withFileTypes: true, recursive: true })
    .filter((e) => e.isFile())
    .map((e) => join(e.parentPath ?? e.path, e.name))
    .sort();
}

const zip = new JSZip();
for (const file of filesUnder(root)) {
  // Forward slashes and a path relative to skill-dist, so the archive unpacks
  // to `ssf-charts/…` exactly as the python version's did.
  const name = relative("skill-dist", file).split("\\").join("/");
  zip.file(name, readFileSync(file), { date: new Date(0) });
}
const bytes = await zip.generateAsync({
  type: "nodebuffer",
  compression: "DEFLATE",
  compressionOptions: { level: 9 },
});
writeFileSync("skill-dist/ssf-charts.zip", bytes);
console.log("skill-dist/ssf-charts.zip");
