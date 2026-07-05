#!/usr/bin/env node
/**
 * Generate the production add-in manifests from the dev manifests by swapping
 * the localhost dev-server origin for the hosted (GitHub Pages) origin. Keeps
 * the two in lockstep: the ONLY difference between dev and prod is the URL —
 * GUIDs, requirement sets and ribbon layout stay identical, so PowerPoint
 * treats them as the same add-in.
 *
 *   node scripts/build-manifest.mjs            # write manifest-*-prod.xml
 *   node scripts/build-manifest.mjs --check    # fail if the prod files are stale
 *
 * The site is served from a custom domain at its root, so the origin has no
 * path segment. Override with PAGES_ORIGIN if the host ever changes.
 */
import { readFileSync, writeFileSync } from "node:fs";

const DEV_ORIGIN = "https://localhost:3000";
const PAGES_ORIGIN = (process.env.PAGES_ORIGIN ?? "https://powerchart.struktureretsundfornuft.dk").replace(/\/$/, "");

const PAIRS = [
  ["manifest.xml", "manifest-prod.xml"],
  ["manifest-excel.xml", "manifest-excel-prod.xml"],
];

const check = process.argv.includes("--check");
let stale = false;

for (const [dev, prod] of PAIRS) {
  const src = readFileSync(dev, "utf8");
  if (!src.includes(DEV_ORIGIN)) {
    throw new Error(`${dev} has no ${DEV_ORIGIN} origin to rewrite — did the dev manifest change shape?`);
  }
  const out = src.replaceAll(DEV_ORIGIN, PAGES_ORIGIN);
  if (out.includes(DEV_ORIGIN)) throw new Error(`${dev}: dev origin survived the swap`);
  if (check) {
    let current = null;
    try {
      current = readFileSync(prod, "utf8");
    } catch {
      /* missing → stale */
    }
    if (current !== out) {
      stale = true;
      console.error(`${prod} is stale — run \`npm run build:manifest\` and commit it`);
    } else {
      console.log(`${prod} is current`);
    }
  } else {
    writeFileSync(prod, out);
    console.log(`wrote ${prod} → ${PAGES_ORIGIN}`);
  }
}

if (check && stale) process.exit(1);
