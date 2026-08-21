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
import {
  scenarioRegressions,
  profileDivergence,
  roundProfile,
  traceNovelty,
  poolScenarioPopulations,
  poolGroupingOutcome,
  poolProfileDisagreements,
  poolPairPosition,
  poolFallbackRates,
  poolInPlaceUpdates,
  roundSpanSeconds,
  paneAgeAtStartSeconds,
} from "./triage.mjs";

/**
 * Every archived round, oldest first — the order `scenarioRegressions` expects.
 *
 * A FILE THAT WILL NOT PARSE IS NAMED, NOT THROWN OVER, and the difference
 * decides what a night does next. `archive` writes straight to the final path
 * rather than writing-then-renaming, so an interrupted write leaves a truncated
 * round behind; unguarded, one of those took the whole gate down with a
 * SyntaxError, node exited 1, and `cycle.mjs` reads any non-zero gate as a
 * REGRESSION. A corrupt download would have stopped the night reporting a fall
 * that never happened.
 *
 * Skipping it silently would be the other half of the same mistake: a round
 * missing from the comparison is a round whose regression cannot be seen, so
 * the caller is told which files were dropped and decides what that is worth.
 * `triage.mjs` already takes exactly this line for the same reason.
 */
export function loadRounds(dir = "rounds", list = readdirSync, read = readFileSync) {
  const unreadable = [];
  const rounds = list(dir)
    .filter((f) => /^\d{3}-.*\.json$/.test(f))
    .sort()
    .map((f) => {
      try {
        return JSON.parse(read(`${dir}/${f}`, "utf8"));
      } catch {
        unreadable.push(f);
        return null;
      }
    })
    .filter(Boolean);
  rounds.unreadable = unreadable;
  return rounds;
}

if (isMain(import.meta.url, process.argv[1])) {
  // EXIT 2 MEANS "I COULD NOT DO MY JOB", and it has to be its own code. This
  // gate's whole value is that exit 1 means a scenario stopped passing — so a
  // gate that fell over on its own reading, exiting 1 the way node does for any
  // uncaught throw, is indistinguishable from the finding it exists to report.
  // `cycle.mjs` acts on that difference: 1 stops the night saying a scenario
  // regressed, 2 stops it saying the gate needs looking at. Same convention
  // `triage.mjs` already uses for a file it cannot read.
  let rounds;
  try {
    rounds = loadRounds();
  } catch (err) {
    console.error(`  the gate could not read rounds/: ${err?.message ?? err}`);
    console.error("  Nothing was judged. This is not a regression — see docs/ROUNDS.md.");
    process.exit(2);
  }
  if (rounds.unreadable?.length) {
    // Loud, because a round missing from the comparison is a round whose
    // regression cannot be seen. Not fatal: the rest of the archive still
    // answers, and refusing to judge 57 good rounds over one bad file would be
    // the worse trade.
    console.error(`  ${rounds.unreadable.length} archived round(s) WOULD NOT PARSE and were left out:`);
    for (const f of rounds.unreadable) console.error(`    rounds/${f}`);
    console.error("  A regression inside one of those cannot be seen from here.");
  }
  if (!rounds.length) {
    console.error("  no readable rounds to judge — nothing was checked");
    process.exit(2);
  }
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
  const real = diverged.filter((d) => !d.flaky);
  const flaky = diverged.filter((d) => d.flaky);
  if (real.length) {
    console.log(`  ${real.length} scenario(s) DIVERGED between slide sizes on the same build:`);
    for (const d of real)
      console.log(
        `    ${d.name} — passed at ${d.passedIn.join(", ")}, failed at ${d.failedIn.join(", ")} (${d.build})`,
      );
    console.log("  Run that profile again, or as a pair, before treating it as a property of the slide size.");
  }
  // NAMED APART, because the response is different. A profile that disagrees
  // with ITSELF has said nothing about its slide size, and sending someone to
  // investigate an aspect ratio for a scenario that is simply flaky is how a
  // useful report teaches people to ignore it.
  //
  // This is the shape the check produced on its first live outing: `explode a
  // degraded picture` passed at 4:3, then passed once and failed once at 16:9
  // on build 17a8204. "Diverged between slide sizes" was true of the worst
  // reading and wrong about the cause.
  if (flaky.length) {
    console.log(`  ${flaky.length} scenario(s) were UNSTABLE WITHIN a slide size, which is not divergence:`);
    for (const d of flaky)
      console.log(`    ${d.name} — passed and failed at ${d.unstableIn.join(", ")} on the same build (${d.build})`);
    console.log("  Treat that as a flaky scenario, not a property of the slide size.");
  }
  // A SCENARIO CAN PASS ON LESS THAN IT USED TO, and none of the three questions
  // around this one can see it. `scenarioRegressions` compares PASS to PASS;
  // divergence compares slide sizes; novelty reads the trace. But `same scale
  // across the deck` scores itself `scaled === charts.length` against a
  // population it DISCOVERS — `probeCharts` returns whatever the deck scan finds
  // — so round 088's `6 of 6` and every earlier round's `8 of 8` are both a pass
  // and the gate said "no scenario regressed" between them.
  //
  // Reported, never fatal, for the same reason as the two above: round 088's six
  // is downstream of a host stall that skipped the scenario seeding the probe
  // charts, which is weather rather than a fault. It is a reason to read the
  // round, and a reason not to quote the pass without its denominator.
  // WHAT THE VERDICT CANNOT SEE. Rounds 092 and 093, one build run twice: 20
  // charts grouped and none refused, then 15 grouped and 4 refused with three
  // slides ending on 24 shapes each — and both reported 13/13 with the identical
  // verdict line. Reported every round, because a number only printed when it
  // looks bad is a number nobody has a baseline for.
  //
  // NEVER A REGRESSION, and the pair above is exactly why: 0 and 4 on the same
  // build is inside this project's own noise floor (1 vs 5, nothing changed). It
  // is a reason to read the round, which is all this line claims.
  // TWO READINGS OF THE SAME FACT, COMPARED. Everything below groups rounds by
  // profile, so a round filed under the wrong one silently contaminates every
  // comparison it appears in — and rounds 115 and 116 did exactly that while
  // `PW_EXPECT_SIZE` reported a match, because the guard read the live host and
  // the archive recorded the pane. Loud, because a wrongly-filed round is worse
  // than a missing one: it answers.
  // THE PAIR IS NOT TWO SAMPLES OF ONE CONDITION. Printed above everything that
  // compares rounds, because every such comparison assumes it is.
  // HOW LONG THIS ROUND TOOK, printed before anything that reads its counters.
  // A slow round is a degraded round in this archive, and the reader has never
  // been able to see which kind they were looking at.
  const newest = rounds[rounds.length - 1];
  const span = roundSpanSeconds(newest);
  const priorSpans = rounds
    .slice(0, -1)
    .map(roundSpanSeconds)
    .filter((n) => typeof n === "number");
  if (span !== null && priorSpans.length >= 3) {
    const sorted = [...priorSpans].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const ratio = median ? span / median : 1;
    const age = paneAgeAtStartSeconds(newest);
    console.log(`
  THIS ROUND TOOK ${span}s (median of ${priorSpans.length} prior round(s): ${median}s)`);
    // THE READING THAT PREDICTS THE COUNTERS. See `paneAgeAtStartSeconds`.
    if (age !== null)
      console.log(
        age < 200
          ? `    Pane was FRESH at the start (${age}s old). Fresh-pane rounds average 0.4 post-retry failures.`
          : `    PANE WAS REUSED — ${age}s old at the start. Reused-pane rounds average 4.6 post-retry
` +
              `    failures against 0.4, and leave 60+ deck shapes against 16. Read the counters below as a
` +
              `    degraded sample; reload the pane between rounds to avoid it.`,
      );
    if (ratio >= 1.5)
      console.log(
        `    ${ratio.toFixed(1)}x the usual. Slow rounds in this archive average roughly twice the post-retry
` + `    failures of fast ones — read this round's counters as a degraded sample, not as a change.`,
      );
  }

  // THE PATHS THE CODE TOOK BECAUSE ITS FIRST CHOICE FAILED. Recorded thousands
  // of times and read by nothing until now — see `poolFallbackRates` for the
  // thirtyfold win and the 40% drift that both went unnoticed.
  // A FEATURE THAT HAS NEVER RUN. See `poolInPlaceUpdates`.
  const ip = poolInPlaceUpdates(rounds);
  // fell + threw: a host-side refusal is a fallback too, and counting only the
  // rule-based declines hid three of them for two rounds.
  const ipDown = ip.fell + ip.threw;
  if (ipDown > 0 && ip.ok === 0)
    console.log(
      `
  IN-PLACE UPDATE HAS NEVER SUCCEEDED — 0 successes against ${ipDown} fallbacks over ${ip.rounds} round(s).
` +
        "    #406 was titled 'The in-place update fired zero times and would not say why' and added the trace" +
        "\n    that answers it. The answer has been sitting in every round file since. See FALLBACKS below for why.",
    );
  else if (ip.ok > 0) {
    console.log(`
  in-place update: ${ip.ok} succeeded, ${ip.fell} declined, ${ip.threw} refused by the host over ${ip.rounds} round(s)`);
    // THIS ROUND ON ITS OWN, WITH ITS DENOMINATOR. A pooled total hides the
    // thing that makes two rounds incomparable: round 148 scored the same 3
    // successes as 147 out of 11 attempts rather than 13, because two scenarios
    // failed and never reached the update. "3 again" is not the same evidence.
    const now = poolInPlaceUpdates(rounds.slice(-1));
    const attempts = now.ok + now.fell + now.threw;
    console.log(
      `    this round: ${now.ok} succeeded of ${attempts} attempt(s)` +
        (now.threw ? ` — ${now.threw} refused BY THE HOST, which is a defect, not a decline` : ""),
    );
  }
  // WHY, not just how many. A decline the differ made on purpose and a write the
  // host threw out read identically in a total.
  if (ip.reasons.length) {
    console.log("    why it fell back:");
    for (const r of ip.reasons.slice(0, 6)) console.log(`      ${String(r.n).padStart(4)}x  ${r.why}`);
  }
  if (ip.unexplained.length)
    console.log(
      `    ${ip.unexplained.length} carried NO reason at all — open these first, they are the ones no category fits`,
    );

  // WHAT IT TOOK TO GET THIS ROUND AT ALL. A successful recovery erases its own
  // evidence — the round that follows it looks like any other — so a round run
  // against a host that was already unwell read as clean. Round 148 took three
  // attempts and then failed two scenarios that had not failed in 109 rounds;
  // round 149's browser died 245 seconds in. Neither fact was archived.
  const dr = rounds[rounds.length - 1]?.driverRun;
  if (dr && Number(dr.attempts) > 1)
    console.log(
      `
  THIS ROUND NEEDED ${dr.attempts} ATTEMPTS — recovered from: ${(dr.recovered ?? []).join(", ") || "unrecorded"}
    A round the driver had to rescue is evidence taken from a host that was already unwell.
    Read a scenario failure here against that, not against a clean round.`,
    );

  const fb = poolFallbackRates(rounds);
  if (fb.length) {
    console.log(`
  FALLBACKS TAKEN — this round against the median of ${fb[0].rounds} prior round(s)`);
    for (const r of fb) {
      // DRIFT FIRST, because it is the reading a median cannot give. A signal
      // that climbs steadily looks NORMAL against its own history the whole way
      // up: `in-place update fell back to a redraw` went from 9 to 13 per round
      // across sixty rounds, and by the time anyone looked, "now" and "usually"
      // were both 13. The oldest third against the newest third sees the shape
      // a median absorbs.
      const drift =
        r.newest > r.oldest * 1.3 && r.newest - r.oldest >= 2
          ? `  <- RISING, ${r.oldest} to ${r.newest} across ${r.span}-round thirds`
          : r.oldest > r.newest * 1.3 && r.oldest - r.newest >= 2
            ? `  <- falling, ${r.oldest} to ${r.newest}`
            : "";
      console.log(`    ${r.label.padEnd(38)} ${String(r.now).padStart(4)}  (usually ${r.median})${drift}`);
    }
  }

  const pos = poolPairPosition(rounds);
  if (pos.pairs >= 4) {
    const moved = pos.worse + pos.better;
    console.log(
      `
  PAIR POSITION — over ${pos.pairs} build(s) run twice: the SECOND round was worse ` +
        `${pos.worse}x, better ${pos.better}x, unchanged ${pos.tied}x`,
    );
    if (moved > 0 && pos.worse > pos.better * 2)
      console.log(
        `    ${pos.worse} of the ${moved} pairs that moved went the same way. That is a direction, not a mood.
` +
          `    THE CAUSE IS THE PANE, not the position: a second round used to INHERIT the first round's
` +
          `    pane, and pane age separates post-retry 0.4 from 4.6. The driver reloads it between rounds
` +
          `    now, so this count is mostly history — ${pos.secondFresh} of ${pos.pairs} pairs had a fresh second round.`,
      );
  }

  const disagreed = poolProfileDisagreements(rounds);
  if (disagreed.length) {
    console.log(`
  SLIDE SIZE DISAGREES — ${disagreed.length} round(s) filed under a profile the driver did not measure`);
    for (const d of disagreed)
      console.log(`    ${d.build}  archive says ${d.pane} (from ${d.source}), the driver measured ${d.driver}`);
    console.log("    Every profile comparison below groups by the ARCHIVE's value. Treat these rounds as unfiled.");
  }

  const grouping = poolGroupingOutcome(rounds);
  if (grouping) {
    const { now, refusedMedian, rounds: priorRounds } = grouping;
    console.log(
      `  GROUPING, which no scenario verdict reports: ${now.grouped} chart(s) grouped, ` +
        // NO BASELINE IS NOT A BASELINE OF ZERO. This used to print `usually 0`
        // when there was no history at all, which reads as "clean until now".
        (refusedMedian === null
          ? `${now.refused} refused (no baseline — ${priorRounds} prior round(s) is too few to say what is usual)`
          : `${now.refused} refused (usually ${refusedMedian} over ${priorRounds} prior round(s))`),
    );
    console.log(`    the deck ended holding ${now.deck.join(",")} shape(s) per slide`);
    if (now.refused > 0)
      console.log(
        [
          "    A refused chart is left as loose shapes in its own box — it keeps its config",
          "    and looks identical to the scenario, which is why 13/13 can hide it. Inside the",
          "    noise floor unless a PAIR on one build agrees; read the deck line above.",
        ].join("\n"),
      );
  }
  const shrunk = poolScenarioPopulations(rounds);
  if (shrunk.length) {
    console.log(`  ${shrunk.length} scenario(s) PASSED ON A SMALLER POPULATION than they usually run:`);
    for (const p of shrunk)
      console.log(
        `    ${p.name} — ${p.now} this round, usually ${p.usual} over ${p.rounds} prior round(s)` +
          `${p.ok ? " (and it still reports PASS)" : ""}`,
      );
    console.log("  A ratio whose bottom half moved is not the same evidence. Read why before comparing it.");
  }
  // A THIRD QUESTION, and the cheapest of the three to answer wrongly. The two
  // above ask about scenarios — thirteen named outcomes a person already reads.
  // This asks about the TRACE, which is 95K characters nobody can count by eye,
  // and its whole job is to say which parts of it are worth the reading.
  //
  // Never fatal, for the same reason divergence is not: it reports difference,
  // and a build is not broken for being different. Exiting non-zero here would
  // make every landed fix fail the gate on the night it starts working.
  const nov = traceNovelty(rounds);
  if (nov.novel.length) {
    console.log(`  ${nov.novel.length} trace signature(s) NEVER SEEN in ${nov.priors} prior round(s):`);
    for (const s of nov.novel.slice(0, 8)) console.log(`    ${String(s.n).padStart(4)}x  ${s.sig}`);
    console.log("  Read the trace. A shape this archive has never produced is the reason to.");
  }
  if (nov.sinceBuild.length) {
    console.log(
      `  ${nov.sinceBuild.length} signature(s) are NEW BEHAVIOUR, not a spike (absent recently, common now):`,
    );
    // THE BUILD IT STARTED IN, per signature. This printed the build being
    // judged, for every entry — so one 064-era signature was blamed on nine
    // consecutive innocent commits, the newest of them a slide-counter fix.
    for (const s of nov.sinceBuild.slice(0, 8))
      console.log(`    ${String(s.n).padStart(4)}x  ${s.sig}${s.startedIn ? `  (first seen in ${s.startedIn})` : ""}`);
    console.log("  The shape a mechanism makes when it starts working — check it is one that build widened.");
  }
  if (nov.spikes.length) {
    console.log(`  ${nov.spikes.length} signature(s) SPIKED against their own history:`);
    for (const s of nov.spikes.slice(0, 8))
      console.log(`    ${String(s.n).padStart(4)}x (usually ${s.median})  ${s.sig}`);
    console.log("  These had a baseline and left it. Most likely to be the round's real story.");
  }
  if (!nov.novel.length && !nov.sinceBuild.length && !nov.spikes.length)
    console.log(
      `  nothing new in the trace — ${nov.vocabulary} known signature(s) across ${nov.priors} prior round(s).` +
        " The trace is still the evidence; this only says where to start.",
    );
  if (!gone.length) {
    console.log(
      `  no scenario regressed — checked the newest of ${rounds.length} archived round(s)` +
        ` at ${roundProfile(rounds[rounds.length - 1])}`,
    );
    process.exit(0);
  }
  console.error(`  ${gone.length} scenario(s) STOPPED PASSING in the newest round:`);
  for (const g of gone)
    console.error(
      `    ${g.name} — ${g.failed === 1 ? "FIRST failure" : `failed ${g.failed} times`} in ${g.ran} round(s) at this profile`,
    );
  console.error("  A round is evidence; this is the only thing that holds a build to it. See docs/ROUNDS.md.");
  process.exit(1);
}
