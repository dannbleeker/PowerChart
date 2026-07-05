#!/usr/bin/env node
/**
 * Post-process the Vite build output for GitHub Pages hosting. Vite only emits
 * imported assets, but the add-in manifests reference the ribbon icons by URL
 * (/assets/icon-*.png) and those live in the repo's assets/ folder — so copy
 * them into dist/assets/ or the hosted icon URLs 404. Also drops a CNAME so an
 * Actions-based deploy keeps the custom domain.
 */
import { copyFileSync, mkdirSync, readdirSync, writeFileSync } from "node:fs";

const DOMAIN = process.env.PAGES_DOMAIN ?? "powerchart.struktureretsundfornuft.dk";
const dist = "dist";

// Ribbon icons referenced by the manifests (not bundled by Vite).
mkdirSync(`${dist}/assets`, { recursive: true });
const icons = readdirSync("assets").filter((f) => /^icon-\d+\.png$/.test(f));
for (const f of icons) copyFileSync(`assets/${f}`, `${dist}/assets/${f}`);

// Custom-domain marker so the Pages deploy doesn't reset the domain.
writeFileSync(`${dist}/CNAME`, `${DOMAIN}\n`);

console.log(`pages-postbuild: copied ${icons.length} icon(s), wrote CNAME (${DOMAIN})`);
