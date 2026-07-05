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
import { copyFileSync, mkdirSync, readdirSync } from "node:fs";

const dist = "dist";

// Ribbon icons referenced by the manifests (not bundled by Vite).
mkdirSync(`${dist}/assets`, { recursive: true });
const icons = readdirSync("assets").filter((f) => /^icon-\d+\.png$/.test(f));
for (const f of icons) copyFileSync(`assets/${f}`, `${dist}/assets/${f}`);

console.log(`pages-postbuild: copied ${icons.length} ribbon icon(s) into ${dist}/assets/`);
