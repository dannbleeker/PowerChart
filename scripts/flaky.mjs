#!/usr/bin/env node
/**
 * Compare several runs of the same suite and name the tests that disagreed.
 *
 * A flaky test that passes on the run you are looking at is indistinguishable
 * from a good one, which is why this project has only ever found flakes by
 * accident. The one found this year — a 40ms budget racing a round trip, green
 * idle and red under load — was misdiagnosed first as a stale build artifact,
 * and cost an hour before anyone thought to run the suite twice.
 *
 * The comparison is per TEST, not per run: a suite that is red the same way
 * three times is a broken build and must not be filed as a flake, because the
 * two want completely different work. Exit 3 means "a test disagreed with
 * itself"; exit 0 means every run agreed, red or green.
 *
 * Usage: node scripts/flaky.mjs run-1.json run-2.json run-3.json
 */
import { readFileSync } from "fs";
import { basename } from "node:path";
import { isMain } from "./is-main.mjs";

/**
 * Every test in a vitest JSON report, by full name, with what it did.
 *
 * Full name rather than title, because two files may legitimately use the same
 * `it()` text and merging them would invent a disagreement that never happened.
 *
 * `basename` rather than a split on "/": vitest reports absolute paths, and on
 * Windows those are backslashed, so the split left the whole path in the key and
 * the report printed `C:\repo\test\foo.test.ts › …` for every row.
 */
export function outcomes(report) {
  const out = new Map();
  for (const file of report?.testResults ?? []) {
    for (const t of file.assertionResults ?? []) {
      const name = `${file.name ? basename(file.name) : "?"} › ${(t.ancestorTitles ?? []).join(" › ")} › ${t.title}`;
      out.set(name, t.status);
    }
  }
  return out;
}

/**
 * Tests whose outcome was not the same in every run.
 *
 * A test missing from a run counts as a disagreement too, and is worth more
 * than a status flip: it means the file did not load, which usually takes its
 * whole suite with it and is exactly the failure a summary count hides.
 */
export function disagreements(runs) {
  const names = new Set(runs.flatMap((r) => [...r.keys()]));
  const out = [];
  for (const name of names) {
    const seen = runs.map((r) => r.get(name) ?? "did-not-run");
    if (new Set(seen).size > 1) out.push({ name, seen });
  }
  return out.sort((a, b) => a.name.localeCompare(b.name));
}

/** The issue body a run posts. */
export function reportBody(found, runCount) {
  if (!found.length) {
    return `All ${runCount} runs agreed on every test. No flakes.\n`;
  }
  const rows = found.map((f) => `- \`${f.name}\`\n  - ${f.seen.join(" · ")}`).join("\n");
  return (
    `${found.length} test(s) did not do the same thing in every one of ${runCount} runs, ` +
    `on a deliberately loaded machine.\n\n` +
    `A flake is not a small problem here: it makes every future red run ambiguous, and this ` +
    `project has already spent an hour diagnosing one as something else. Fix the test or the ` +
    `budget it depends on — do not retry it.\n\n${rows}\n`
  );
}

function main() {
  const files = process.argv.slice(2);
  if (files.length < 2) throw new Error("usage: flaky.mjs <run-1.json> <run-2.json> [...]");
  const runs = files.map((f) => outcomes(JSON.parse(readFileSync(f, "utf8"))));
  const found = disagreements(runs);
  process.stdout.write(reportBody(found, runs.length));
  if (found.length) process.exitCode = 3;
}

if (isMain(import.meta.url, process.argv[1])) main();
