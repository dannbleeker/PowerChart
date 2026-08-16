#!/usr/bin/env node
/**
 * Fail when a scenario that WAS passing has stopped.
 *
 * The only automatic check this project has on a round's own result. Every other
 * number a round produces is read by a person and then filed, so rounds 070-072
 * — which took `same scale across the deck` from 35 consecutive failures to
 * three consecutive passes — bought a result that nothing guarded. A later build
 * could take it back and no gate would notice.
 *
 *     npm run rounds:gate
 *
 * NOT part of CI, and it cannot be: CI has no rounds. It runs after archiving,
 * against `rounds/`, and its exit code is the point.
 *
 * It is deliberately quiet about everything else. A gate that reports on a round
 * being merely worse is a gate that cries wolf on a host whose mood swings 4-of-5
 * to 1-of-5 with nothing changed — and `docs/BACKLOG.md` records what happens to
 * a gate like that: it gets switched off.
 */
import { readFileSync, readdirSync } from "fs";
import { isMain } from "./is-main.mjs";
import { scenarioRegressions, profileDivergence, roundProfile } from "./triage.mjs";

/** Every archived round, oldest first — the order `scenarioRegressions` expects. */
export function loadRounds(dir = "rounds", list = readdirSync, read = readFileSync) {
  return list(dir)
    .filter((f) => /^\d{3}-.*\.json$/.test(f))
    .sort()
    .map((f) => JSON.parse(read(`${dir}/${f}`, "utf8")));
}

if (isMain(import.meta.url, process.argv[1])) {
  const rounds = loadRounds();
  const gone = scenarioRegressions(rounds);
  // A SECOND, DIFFERENT QUESTION. The gate above asks whether a scenario fell
  // against its OWN history; this asks whether one slide size failed what
  // another passed on the same build. Round 077 was exactly that — 10 of 13 at
  // 4:3 against 13 of 13 at 16:9 — and nothing said so automatically.
  //
  // Reported, never fatal. A nightly cycle runs 16:9 twice and 4:3 once as
  // VALIDATION, and the agreed response to divergence is to run 4:3 again or on
  // its own, not to fail the build. Exiting non-zero here would turn a signal
  // that means "look closer" into one that means "stop", which is how a useful
  // report becomes an ignored one.
  const diverged = profileDivergence(rounds);
  if (diverged.length) {
    console.log(`  ${diverged.length} scenario(s) DIVERGED between slide sizes on the same build:`);
    for (const d of diverged)
      console.log(
        `    ${d.name} — passed at ${d.passedIn.join(", ")}, failed at ${d.failedIn.join(", ")} (${d.build})`,
      );
    console.log("  Run that profile again, or as a pair, before treating it as a property of the slide size.");
  }
  if (!gone.length) {
    console.log(
      `  no scenario regressed — checked the newest of ${rounds.length} archived round(s)` +
        ` at ${roundProfile(rounds[rounds.length - 1])}`,
    );
    process.exit(0);
  }
  console.error(`  ${gone.length} scenario(s) STOPPED PASSING in the newest round:`);
  for (const g of gone) console.error(`    ${g.name} — had passed the previous ${g.passedIn} rounds running`);
  console.error("  A round is evidence; this is the only thing that holds a build to it. See docs/ROUNDS.md.");
  process.exit(1);
}
