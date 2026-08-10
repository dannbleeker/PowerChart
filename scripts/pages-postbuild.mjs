#!/usr/bin/env node
/**
 * Post-process the Vite build output for GitHub Pages hosting. Vite only emits
 * imported assets, but the add-in manifests reference the ribbon icons by URL
 * (/assets/icon-*.png) and those live in the repo's assets/ folder — so copy
 * them into dist/assets/ or the hosted icon URLs 404.
 *
 * (The CNAME and the static legal pages ride along automatically: they live in
 * public/, which Vite copies verbatim into dist/.)
 */
import { copyFileSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";

const dist = "dist";

// Ribbon icons referenced by the manifests (not bundled by Vite).
mkdirSync(`${dist}/assets`, { recursive: true });
const icons = readdirSync("assets").filter((f) => /^icon-\d+\.png$/.test(f));
for (const f of icons) copyFileSync(`assets/${f}`, `${dist}/assets/${f}`);

/**
 * What the SITE is serving, in a file small enough to re-fetch on every boot.
 *
 * GitHub Pages sends the pane's HTML with `Cache-Control: max-age=600` and does
 * not let us set headers, so for ten minutes after a deploy PowerPoint keeps
 * handing back the cached page — which names the PREVIOUS hashed bundle, still
 * in the browser's own cache even after that file 404s on the server. Old page,
 * old script, old stamp, and nothing on screen saying so.
 *
 * That has cost two rounds: one ran a fix that was not in the build under test
 * and was read as evidence about it, and one was spent hard-reloading and
 * wondering why the stamp would not move.
 *
 * The stamp is read back OUT OF THE BUILT BUNDLE rather than recomputed here.
 * Calling `buildStamp()` again would mint a fresh timestamp and produce a file
 * that disagreed with the pane by a minute or two on every deploy — a staleness
 * check that always cries stale is worse than none. This is the exact string
 * the pane will display.
 */
// Matched WITHOUT its surrounding quotes: the minifier picks those, and it
// picked backticks — a pattern that insisted on `"` found nothing and failed
// the build on its first real run, which is at least the right direction to
// fail in.
const bundle = readdirSync(`${dist}/assets`).find((f) => /^taskpane-.*\.js$/.test(f));
const stamp =
  bundle &&
  /[0-9a-f]{7,40} · \d{4}-\d{2}-\d{2} \d{2}:\d{2}Z/.exec(readFileSync(`${dist}/assets/${bundle}`, "utf8"))?.[0];
if (!stamp) {
  // Loudly. A missing build.json leaves the pane's check permanently unable to
  // tell, which on screen is indistinguishable from "you are up to date".
  console.error(`pages-postbuild: no build stamp in ${bundle ? `${dist}/assets/${bundle}` : "any taskpane bundle"}`);
  process.exit(1);
}
writeFileSync(`${dist}/build.json`, `${JSON.stringify({ build: stamp })}\n`);

console.log(`pages-postbuild: ${icons.length} ribbon icon(s) into ${dist}/assets/, build.json = ${stamp}`);
