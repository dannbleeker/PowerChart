#!/usr/bin/env node
/**
 * Audit a real PowerPoint deck for what PowerChart actually left in it.
 *
 * Every hard diagnosis in this project was settled by reading the .pptx, not
 * by reading the add-in's own report — including the one where the report was
 * the thing that was wrong. A 39-slide web run said it had tagged 20 charts;
 * the file carried 31. The log agreed with itself and was false, and only the
 * bytes said so.
 *
 * So this reads the bytes. It knows nothing about what the run believed and
 * takes no input from it: slot tags, config tags, groups and frames as they
 * exist in the file. Point it at a deck saved out of a real host and it says
 * what is there.
 *
 *   node scripts/verify-deck.mjs <deck.pptx> [--json]
 *
 * Exit code is 0 when the deck is structurally sound (whatever the run's
 * outcome), 1 when a fault is found, 2 when the file cannot be read — so it
 * can gate a check as well as inform one.
 *
 * "Structurally sound" deliberately excludes run outcomes. A deck missing
 * eight charts because the host dropped them is a bad RUN and a fine FILE; a
 * deck whose tag part is unreferenced is a broken file. Only the second is a
 * fault here, because only the second is something this repo wrote wrong.
 */
import { readFileSync } from "fs";
import JSZip from "jszip";

const CONFIG_TAG = "POWERCHART_CONFIG";
const ORIGIN_TAG = "POWERCHART_ORIGIN";
const SLOT_TAG = "POWERCHART_DEMO_SLOT";
const GROUP_NAME = "PowerChart";
const NOT_COMPLETE = "PowerChart:not-complete";

/** Undo the five XML attribute-value escapes, in the one order that is safe. */
const unescapeAttr = (v) =>
  v
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");

/** Slide parts in DECK order — numeric, never lexicographic (slide10 < slide9). */
function slidePartsInOrder(zip) {
  return Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => Number(/(\d+)/.exec(a)[1]) - Number(/(\d+)/.exec(b)[1]));
}

/** The shape tree's direct children, with nesting handled — see ooxml.ts. */
function topLevel(xml) {
  const tree = /<p:spTree>([\s\S]*)<\/p:spTree>/.exec(xml);
  if (!tree) return [];
  const out = [];
  let depth = 0;
  let start = -1;
  const tag = /<(\/?)([A-Za-z0-9:_-]+)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  let m;
  while ((m = tag.exec(tree[1]))) {
    if (m[4] === "/") {
      if (depth === 0) out.push(m[0]);
      continue;
    }
    if (m[1] === "/") {
      depth--;
      if (depth === 0 && start >= 0) {
        out.push(tree[1].slice(start, m.index + m[0].length));
        start = -1;
      }
      continue;
    }
    if (depth === 0) start = m.index;
    depth++;
  }
  // The first two children are the tree's own mandatory nvGrpSpPr/grpSpPr
  // boilerplate, not shapes.
  return out.filter((e) => !/^<p:(nvGrpSpPr|grpSpPr)[\s>]/.test(e));
}

export async function readDeck(path) {
  return readDeckBytes(readFileSync(path));
}

/** The same audit against bytes already in hand — what the tests drive. */
export async function readDeckBytes(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const types = await zip.file("[Content_Types].xml")?.async("string");
  const slides = slidePartsInOrder(zip);
  const referencedTagParts = new Set();
  const rows = [];

  for (const [i, part] of slides.entries()) {
    const xml = await zip.file(part).async("string");
    const relsPath = part.replace("slides/", "slides/_rels/") + ".rels";
    const rels = (await zip.file(relsPath)?.async("string")) ?? "";
    const tagRefs = [...rels.matchAll(/Id="([^"]+)"[^>]*Target="\.\.\/tags\/(tag\d+\.xml)"/g)].map((m) => m[2]);

    let slot = null;
    let config = null;
    let origin = false;
    const missingParts = [];
    for (const t of tagRefs) {
      const file = zip.file(`ppt/tags/${t}`);
      if (!file) {
        missingParts.push(t);
        continue;
      }
      referencedTagParts.add(`ppt/tags/${t}`);
      const tx = await file.async("string");
      const s = new RegExp(`name="${SLOT_TAG}" val="([^"]*)"`).exec(tx);
      if (s) {
        try {
          slot = JSON.parse(unescapeAttr(s[1]));
        } catch {
          slot = { malformed: s[1] };
        }
      }
      const c = new RegExp(`name="${CONFIG_TAG}" val="([^"]*)"`).exec(tx);
      if (c) {
        try {
          config = JSON.parse(unescapeAttr(c[1]));
        } catch {
          config = { malformed: true };
        }
      }
      if (tx.includes(ORIGIN_TAG)) origin = true;
    }

    const shapes = topLevel(xml);
    const groups = shapes.filter((e) => e.startsWith("<p:grpSp"));
    // The chart object: a real group, or the single named shape a degraded
    // picture leaves behind. Both are "one object named PowerChart".
    const named = shapes.filter((e) => new RegExp(`name="${GROUP_NAME}"`).test(e));
    rows.push({
      index: i,
      slot: slot?.i ?? null,
      title: slot?.title ?? null,
      run: slot?.run ?? null,
      shapes: shapes.length,
      groups: groups.length,
      chartObject: named.length > 0,
      picture: named.length > 0 && groups.length === 0,
      config: !!config,
      configMalformed: config?.malformed === true,
      origin,
      stamped: xml.includes(NOT_COMPLETE),
      missingTagParts: missingParts,
    });
  }

  const allTagParts = Object.keys(zip.files).filter((f) => /^ppt\/tags\/tag\d+\.xml$/.test(f));
  return { rows, allTagParts, referencedTagParts, types: types ?? "" };
}

/** Faults are things THIS REPO wrote wrong — never things the host did badly. */
export function faultsIn({ rows, allTagParts, referencedTagParts, types }) {
  const faults = [];
  for (const r of rows) {
    if (r.missingTagParts.length)
      faults.push(`slide ${r.index + 1}: references missing tag part(s) ${r.missingTagParts.join(", ")}`);
    if (r.configMalformed) faults.push(`slide ${r.index + 1}: ${CONFIG_TAG} is not valid JSON`);
    if (r.slot !== null && typeof r.slot !== "number") faults.push(`slide ${r.index + 1}: slot tag is malformed`);
    // A config with nothing to hang it on would be unreachable from the pane.
    if (r.config && !r.chartObject)
      faults.push(`slide ${r.index + 1}: carries a config tag but no "${GROUP_NAME}" object to load it from`);
  }
  for (const p of allTagParts) {
    if (!referencedTagParts.has(p)) faults.push(`${p}: tag part is not referenced by any slide`);
    if (types && !types.includes(`PartName="/${p}"`))
      faults.push(`${p}: no [Content_Types].xml override — PowerPoint will reject the file`);
  }
  // Two runs' slides in one deck is NOT a fault — it is the case the run token
  // exists to survive — but a repeated (run, slot) pair means two slides claim
  // to be the same item of the same run, which nothing should produce.
  const seen = new Map();
  for (const r of rows) {
    if (r.run === null || r.slot === null) continue;
    const key = `${r.run}#${r.slot}`;
    if (seen.has(key)) faults.push(`slides ${seen.get(key) + 1} and ${r.index + 1}: same run and slot`);
    else seen.set(key, r.index);
  }
  return faults;
}

function report(deck, faults) {
  const { rows } = deck;
  const w = (s, n) =>
    String(s ?? "-")
      .padEnd(n)
      .slice(0, n);
  console.log(`${rows.length} slide(s)\n`);
  console.log(`  ${w("#", 4)}${w("title", 26)}${w("slot", 6)}${w("shapes", 7)}${w("object", 9)}${w("config", 7)}flags`);
  for (const r of rows) {
    const flags = [r.stamped && "NOT-COMPLETE", r.config && !r.origin && "no-origin"].filter(Boolean).join(" ");
    console.log(
      `  ${w(r.index + 1, 4)}${w(r.title, 26)}${w(r.slot, 6)}${w(r.shapes, 7)}` +
        `${w(r.picture ? "picture" : r.groups ? "group" : r.chartObject ? "shape" : "loose", 9)}` +
        `${w(r.config ? "yes" : "—", 7)}${flags}`,
    );
  }
  const runs = [...new Set(rows.map((r) => r.run).filter(Boolean))];
  const charts = rows.filter((r) => r.chartObject);
  console.log(
    `\n  runs in deck: ${runs.length ? runs.join(", ") : "none"}` +
      `\n  chart objects: ${charts.length} (${charts.filter((r) => r.config).length} re-editable, ` +
      `${charts.filter((r) => !r.config).length} NOT re-editable)` +
      `\n  pictures: ${rows.filter((r) => r.picture).length}` +
      `\n  stamped NOT COMPLETE: ${rows.filter((r) => r.stamped).length}`,
  );
  if (faults.length) {
    console.log(`\n  ${faults.length} FAULT(S):`);
    for (const f of faults) console.log(`   - ${f}`);
  } else {
    console.log(`\n  no structural faults`);
  }
}

// CLI only when invoked directly; importable as a library for the tests.
const invokedDirectly = process.argv[1] && process.argv[1].endsWith("verify-deck.mjs");
const [, , path, ...flags] = process.argv;
if (!invokedDirectly) {
  // imported — nothing to do
} else if (!path) {
  console.error("usage: node scripts/verify-deck.mjs <deck.pptx> [--json]");
  process.exit(2);
} else {
  let deck;
  try {
    deck = await readDeck(path);
  } catch (err) {
    console.error(`could not read ${path}: ${err?.message ?? err}`);
    process.exit(2);
  }
  const faults = faultsIn(deck);
  if (flags.includes("--json")) {
    console.log(JSON.stringify({ slides: deck.rows, faults }, null, 2));
  } else {
    report(deck, faults);
  }
  process.exit(faults.length ? 1 : 0);
}
