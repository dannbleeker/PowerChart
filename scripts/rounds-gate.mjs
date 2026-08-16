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
import { scenarioRegressions } from "./triage.mjs";

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
  if (!gone.length) {
    console.log(`  no scenario regressed — checked the newest of ${rounds.length} archived round(s)`);
    process.exit(0);
  }
  console.error(`  ${gone.length} scenario(s) STOPPED PASSING in the newest round:`);
  for (const g of gone) console.error(`    ${g.name} — had passed the previous ${g.passedIn} rounds running`);
  console.error("  A round is evidence; this is the only thing that holds a build to it. See docs/ROUNDS.md.");
  process.exit(1);
}
