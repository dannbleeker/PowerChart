#!/usr/bin/env node
/**
 * The suite may grow. It may not silently shrink.
 *
 * `CLAUDE.md` records what this is for: a reorg once deleted 43 tests and the
 * suite still went green, because nothing compared the count before against the
 * count after. The rule written down in response — "pin the total first and
 * check it after" — is enforced by whoever remembers it, which is the same as
 * not being enforced.
 *
 * A high-water mark instead of an exact number, deliberately. An exact count
 * taxes every PR that adds a test with a second edit and a merge conflict, and a
 * gate people resent is a gate people route around. A floor costs nothing to
 * anyone adding tests and fails loudly for the one case it exists for.
 *
 * Usage:
 *   node scripts/test-count.mjs results.json          # check
 *   node scripts/test-count.mjs results.json --update # re-record the mark
 *
 * The JSON is vitest's own (`--reporter=json --outputFile=...`), so this never
 * runs the suite itself — it reads a run that already happened, which is what
 * keeps it free in CI.
 */
import { readFileSync, writeFileSync } from "fs";

const MARK_FILE = new URL("../test/fixtures/test-count.json", import.meta.url);

/**
 * How far below the mark is a rounding error rather than a deletion.
 *
 * Zero. There is no such thing as accidentally losing one test: a file moved,
 * a `describe` dropped, a merge that took the wrong side. If the number goes
 * down on purpose, `--update` says so in the diff, which is the whole point —
 * the deletion becomes something a reviewer sees rather than something the
 * suite absorbs.
 */
const SLACK = 0;

/** What a vitest JSON report says the suite ran. */
export function countOf(report) {
  const total = report?.numTotalTests;
  if (typeof total !== "number" || !Number.isFinite(total)) {
    throw new Error("that is not a vitest JSON report — no numTotalTests in it");
  }
  return total;
}

/** The verdict, with the sentence a reader needs. */
export function judgeCount(now, mark) {
  if (now < mark - SLACK) {
    return {
      ok: false,
      message:
        `the suite has SHRUNK: ${now} tests, against a recorded ${mark}. ` +
        `${mark - now} test(s) went missing.\n\n` +
        `If that was deliberate — tests moved out, a feature deleted — re-record it with\n` +
        `  node scripts/test-count.mjs <results.json> --update\n` +
        `and the drop shows up in the diff, where a reviewer can see it.`,
    };
  }
  return {
    ok: true,
    message: now > mark ? `${now} tests, up from ${mark}` : `${now} tests, the recorded mark`,
  };
}

function main() {
  const [file, ...flags] = process.argv.slice(2);
  if (!file) throw new Error("usage: test-count.mjs <vitest-json> [--update]");
  const now = countOf(JSON.parse(readFileSync(file, "utf8")));
  const mark = JSON.parse(readFileSync(MARK_FILE, "utf8")).tests;
  if (flags.includes("--update")) {
    writeFileSync(MARK_FILE, `${JSON.stringify({ tests: now }, null, 2)}\n`);
    process.stdout.write(`recorded ${now} tests\n`);
    return;
  }
  const verdict = judgeCount(now, mark);
  process.stdout.write(`${verdict.message}\n`);
  if (!verdict.ok) process.exitCode = 1;
  // The mark rises on its own. Writing it back here rather than asking every
  // test-adding PR to bump a number by hand is what keeps this cheap enough to
  // leave switched on — CI commits nothing, so the file only moves when a human
  // runs it, and a stale-but-lower mark is harmless.
  else if (now > mark) writeFileSync(MARK_FILE, `${JSON.stringify({ tests: now }, null, 2)}\n`);
}

if (process.argv[1] && import.meta.url.endsWith(process.argv[1].split("/").pop())) main();
