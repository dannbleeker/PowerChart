#!/usr/bin/env node
/**
 * Diff a real PowerPoint's answer sheet against the fake host's.
 *
 * This is the payoff for `src/render/host-probe.ts`. Every Office.js assertion
 * in this repo runs against a fake, and the fake has never been checked against
 * the thing it stands for. Its faults are behaviours a real host taught us; its
 * happy path is a set of things we assumed. Where an assumption is wrong, every
 * test resting on it is confidently wrong with it, and a green suite says
 * nothing about that.
 *
 *   node scripts/host-diff.mjs <real-host-answers.json> [fake-answers.json]
 *
 * The fake's sheet is the one CI freezes in `test/host-probe.test.ts`; pass it
 * explicitly, or let this read the baseline committed beside it.
 *
 * Only the `answer` field is compared. Timings and error text differ between
 * any two runs of the same host and would bury the signal.
 *
 * Exit 0 when the two agree, 1 when they do not — or when the real host never
 * got as far as answering — and 2 when a file cannot be read, so it can gate as
 * well as inform.
 */
import { readFileSync } from "fs";
import {
  FAKE_BASELINE,
  KNOWN_DIVERGENCES,
  PENDING_QUESTIONS,
  answersOf,
  diffAnswers,
  sheetOf,
  NEVER_ASKED,
} from "./host-baseline.mjs";
import { isMain } from "./is-main.mjs";

// Re-exported so the tables have ONE home and every existing importer keeps
// working: two copies of this data is how a claim quietly stops matching its
// check, which is the failure this whole file exists to prevent.
export { FAKE_BASELINE, KNOWN_DIVERGENCES, PENDING_QUESTIONS, answersOf, diffAnswers, sheetOf, NEVER_ASKED };

if (isMain(import.meta.url, process.argv[1])) {
  const [realPath, fakePath] = process.argv.slice(2).filter((a) => !a.startsWith("--"));
  if (!realPath) {
    console.error("usage: node scripts/host-diff.mjs <real-host-answers.json> [fake-answers.json]");
    process.exit(2);
  }
  let realSheet, fakeSheet;
  try {
    realSheet = JSON.parse(readFileSync(realPath, "utf8"));
    fakeSheet = fakePath ? JSON.parse(readFileSync(fakePath, "utf8")) : FAKE_BASELINE;
  } catch (err) {
    console.error(`could not read the answer sheets: ${err.message}`);
    process.exit(2);
  }
  const real = answersOf(realSheet);
  const fake = answersOf(fakeSheet);
  if (!real || !fake) {
    console.error(
      "that file is not an answer sheet (expected kind: ssf-charts-host-answers, or powerchart-host-answers from an older round)",
    );
    process.exit(2);
  }
  // The guard above cannot fire for any JSON OBJECT: `answersOf` deliberately
  // returns a bare map verbatim, because `FAKE_BASELINE` is one. So
  // `node scripts/host-diff.mjs package.json` was accepted as an answer sheet,
  // its top-level keys became question ids, and the run exited 0 — a wrong path
  // getting a green light from a gate.
  //
  // The honest test is not the shape of the file, it is whether it answered
  // anything: a sheet with no answers in this vocabulary is not a sheet.
  if (!Object.keys(real).length) {
    console.error("that file answers no host questions — is it an answer sheet, or did the probe never run?");
    process.exit(2);
  }

  const { agree, differ, notAsked, onlyReal, onlyFake } = diffAnswers(real, fake);
  // The header comes off the SAME object the answers did. Reading it off the
  // outer file left a whole round's upload reporting `REAL HOST ?` and
  // `requirement sets: unknown` over a page of real answers, because those
  // fields live one level down in that shape.
  const header = sheetOf(realSheet) ?? realSheet;
  const sets = Array.isArray(header.requirementSets) ? header.requirementSets.join(", ") : "unknown";
  console.log(
    `\n  REAL HOST ${header.source ?? realSheet.source ?? "?"} · build ${header.build ?? realSheet.build ?? "?"}` +
      `\n  requirement sets: ${sets}` +
      `\n  ${agree.length} of ${agree.length + differ.length} answers match the fake\n`,
  );

  if (differ.length) {
    console.log(`  ${differ.length} DIVERGENCE(S) — the fake and the real host disagree:\n`);
    for (const d of differ) {
      console.log(`    ${d.id}`);
      console.log(`      real host: ${d.real}`);
      console.log(`      fake:      ${d.fake}`);
      if (d.means) console.log(`      what rests on it: ${d.means}`);
      console.log("");
    }
    console.log("  Each of these is one of two things: the fake lies and the tests built on it");
    console.log("  are worth less than they look, or the host does something we did not know.\n");
  }
  if (notAsked.length) {
    console.log(`  ${notAsked.length} QUESTION(S) NEVER PUT — the host would not set the probe up:\n`);
    for (const n of notAsked) console.log(`    ${n.id}  (${n.why})`);
    console.log("\n  Not divergences. Nothing is known about what this host does here —");
    console.log("  the run could not get as far as asking. Fix the setup, run it again.\n");
  }
  // A question one side was never asked is a hole in the comparison, and
  // reporting it as agreement is how a diff stops meaning anything.
  if (onlyReal.length)
    console.log(`  ${onlyReal.length} question(s) the fake has no answer for: ${onlyReal.join(", ")}`);
  if (onlyFake.length)
    console.log(`  ${onlyFake.length} question(s) the real sheet is missing (older build?): ${onlyFake.join(", ")}`);
  if (!differ.length && !notAsked.length && !onlyReal.length && !onlyFake.length) {
    console.log("  The fake agrees with this host on every question it was asked.\n");
  }
  // A question that was never put is not agreement, so it fails the gate too —
  // an incomplete sheet exiting 0 is exactly the false all-clear this tool
  // spent a round learning not to give.
  //
  // `onlyReal`/`onlyFake` count as well, and their absence from this expression
  // contradicted the comment eight lines above ("a question one side was never
  // asked is a hole in the comparison, and reporting it as agreement is how a
  // diff stops meaning anything") and the file header ("exit 1 … when the real
  // host never got as far as answering"). A sheet whose questions no longer
  // match the build's is a hole of exactly that kind.
  process.exit(differ.length || notAsked.length || onlyReal.length || onlyFake.length ? 1 : 0);
}
