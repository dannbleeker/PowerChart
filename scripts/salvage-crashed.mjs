#!/usr/bin/env node
/**
 * File the rounds that already ran, from the crash records that kept them.
 *
 *     node scripts/salvage-crashed.mjs --dry-run
 *     node scripts/salvage-crashed.mjs --into rounds-salvaged
 *
 * THEY DO NOT GO IN `rounds/`, AND THE VERIFICATION IS WHY. `rounds/README.md`
 * opens with the contract: "`NNN-<build>.json`, **oldest first**". The number
 * IS the chronology, and every ordering-sensitive instrument leans on it —
 * `poolFallbackRates` compares the oldest third against the newest third, and
 * "this round" everywhere means the highest number.
 *
 * These records are from 2026-08-26 to 08-29. Filed at the end they would take
 * numbers 311-332 and claim to be the NEWEST rounds in the archive, which is
 * false, and they cannot be interleaved instead: no archived round records when
 * it ran, so there is nothing to interleave against.
 *
 * Measured rather than assumed, on a scratch copy: adding all 22 took
 * `poolGroupingOutcome` to null and `poolFallbackRates` from 5 signals to 0,
 * because the newest round became one with no trace. Those are the guards being
 * honest, and they are also the gate going quiet about its two headline numbers.
 * That is a shape change, and the plan's own rule was to file only if nothing
 * changed shape.
 *
 * So they live beside the archive rather than in it, and pooling them needs an
 * ordering key that does not exist yet. What they are good for TODAY is any
 * question that does not care about order: what a build's verdicts were, how
 * often a scenario has ever failed, whether the 4:3 arm agrees with 16:9. That
 * arm is where all 22 of them are, and it nearly doubles: 26 rounds to 48.
 *
 * WHY THIS EXISTS. A round's verdicts are all in before `collectDeckEvidence`
 * runs, and that function is where this host dies — 9 builds against 4 on the
 * per-phase traces, in a band of 441-572s. When it died there the pane never
 * assembled its run log, so a complete round became a crash record and nothing
 * else. On 2026-08-29 one 4:3 leg produced a full fourteen-scenario result FIVE
 * times, 14/14 in four of them, and archived none of the five.
 *
 * Across the archive: **33 of 49 crash records hold a complete round**, over 13
 * builds. That is about a tenth of `rounds/` again, already measured, already
 * paid for in host time, sitting unread.
 *
 * `src/taskpane/app.ts` now banks the log BEFORE that scan, so this should stop
 * finding new work. It exists for the records written before that landed, and as
 * the belt to that braces.
 *
 * WHAT A SALVAGED ROUND DELIBERATELY LACKS, and why it is still safe to pool:
 *
 *   deck    the scan never returned. `crashes/README.md` is right that a partial
 *           record must never be pooled — an INVENTED inventory is exactly the
 *           fabricated evidence it warns about, and everything downstream reads
 *           that field as measured. So it is absent, not guessed.
 *   trace   the crash keeps `steps` as formatted strings; a round's trace is
 *           structured entries with `message` and `data`. Reshaping prose into
 *           structure would be inventing the same way. Absent.
 *
 * Both omissions make the pooling functions SKIP these rounds — and that is only
 * safe because the stale-`now` class was fixed first (see `isTheRoundBeingJudged`
 * in `scripts/triage.mjs`). Before that fix, a wave of skipped rounds would have
 * made three emitters report an older round's numbers as the newest round's,
 * every run, for ever.
 */
import { readFileSync, writeFileSync, readdirSync, existsSync, unlinkSync, mkdirSync } from "fs";
import { isMain } from "./is-main.mjs";
import { archive } from "./round.mjs";

/** The verdicts a crash record kept, in the order it recorded them. */
export function verdictsFrom(crash) {
  return (crash?.findings ?? [])
    .filter((f) => typeof f?.key === "string" && f.key.startsWith("selftest:"))
    .map((f) => f.value)
    .filter((v) => v && typeof v === "object" && typeof v.name === "string");
}

/** The probe sheet, which is recorded under its own key. */
export function sheetFrom(crash) {
  return (crash?.findings ?? []).find((f) => f?.key === "hostAnswers")?.value;
}

/**
 * WHAT PROFILE THIS ROUND RAN AT, which is the one field a salvage must not
 * guess and must not omit.
 *
 * The archive pools 16:9 and 4:3 in one directory, and averaging two aspect
 * ratios into one number is the rounds 24-and-25 mistake this repo already
 * records — a nightly 4:3 round would repeat it every night. So a round that
 * cannot say what it ran at is refused rather than filed as unknown.
 *
 * Two sources, newest first:
 *   `runLogHead`   written by the pane before the scan, since 2026-08-29. Exact.
 *   the steps      `host  slide size read  width=720 height=540 source=pageSetup`
 *                  — the host's own answer, traced during the round. The LAST
 *                  reading, because the pane reads it at the end and that is the
 *                  one a filed round would have carried.
 *
 * AND THE RUNG IT CAME FROM DECIDES WHETHER IT COUNTS — see `trustedSlideSize`.
 */
export function slideSizeFrom(crash) {
  const head = (crash?.findings ?? []).find((f) => f?.key === "runLogHead")?.value;
  if (head?.slideSize?.width && head?.slideSize?.height) return head.slideSize;
  let found = null;
  for (const step of crash?.steps ?? []) {
    const m = /slide size read\s+width=(\d+)\s+height=(\d+)(?:\s+source=(\S+))?/.exec(String(step));
    if (m) found = { width: Number(m[1]), height: Number(m[2]), ...(m[3] ? { source: m[3] } : {}) };
  }
  return found;
}

/**
 * IS THAT READING GOOD ENOUGH TO FILE A ROUND UNDER? Usually not.
 *
 * `slideSize()` resolves through three rungs and they are not equally true.
 * `pageSetup` is the live host answering. `exportedSlide` is the LAST rung,
 * reached only after the host gave up answering, and it reads a saved file
 * PowerPoint may not have updated yet.
 *
 * `archive` carries the case in its own docstring: **rounds 115 and 116 filed
 * themselves as 720x540 while running on a 960x540 deck**, because the pane fell
 * to that rung at an unlucky moment. What caught it was `driverSlideSize` — the
 * driver's independent reading, recorded beside the pane's.
 *
 * A salvaged round has no driver reading to be caught by. Nobody was holding the
 * other end: the driver filed a crash, not a round. So the fallback rung is
 * refused outright rather than filed and hoped about — this repo's own rule is
 * that filing a round under the wrong profile is worse than not running it, and
 * 18 of 46 records here rest on that rung.
 */
export function trustedSlideSize(size) {
  return size?.source === "pageSetup" ? size : null;
}

/**
 * Build a round from a crash record, or say why it cannot be one.
 *
 * `expectedNames` is the scenario set from an ARCHIVED round of the same build —
 * real evidence rather than a hard-coded count. A literal `>= 14` would refuse
 * the genuinely complete twelve-scenario rounds from earlier in the archive and
 * accept a future partial one the day a fifteenth scenario lands. Checked as a
 * superset by NAME, because a count cannot tell a missing scenario from a
 * renamed one.
 */
export function roundFromCrash(crash, expectedNames, source) {
  const build = String(crash?.build ?? "");
  if (!build) return { ok: false, why: "no build stamp — it cannot be filed against anything" };
  const verdicts = verdictsFrom(crash);
  if (!verdicts.length) return { ok: false, why: "no scenario verdicts at all" };
  if (!expectedNames?.length)
    return { ok: false, why: `no archived round on build ${build.split(" ")[0]} to say what a complete round is` };
  const have = new Set(verdicts.map((v) => v.name));
  const missing = expectedNames.filter((n) => !have.has(n));
  if (missing.length)
    return { ok: false, why: `${missing.length} scenario(s) never reached a verdict: ${missing.join("; ")}` };
  const read = slideSizeFrom(crash);
  if (!read) return { ok: false, why: "nothing in it says what slide size it ran at" };
  const slideSize = trustedSlideSize(read);
  if (!slideSize)
    return {
      ok: false,
      why:
        `its only slide-size reading came from \`${read.source ?? "an unnamed rung"}\`, the fallback the host ` +
        "falls to after giving up — rounds 115 and 116 filed 720x540 while running at 960x540 that way, and a " +
        "salvaged round has no driver reading to be caught by",
    };
  const sheet = sheetFrom(crash);
  return {
    ok: true,
    round: {
      build,
      host: String(crash.host ?? ""),
      slideSize,
      runs: [],
      ...(sheet ? { hostAnswers: sheet } : {}),
      selftest: verdicts,
      /**
       * SAYS WHAT IT IS, on the round itself. A reader who finds a round with no
       * deck and no trace should learn why from the file rather than infer it,
       * and anything that later wants to exclude salvaged rounds needs a field
       * to test. Named for the record it came from so the original is one
       * `ls` away.
       */
      salvagedFrom: source,
    },
  };
}

/** Scenario names per build, from what has actually been filed. */
export function expectedByBuild(dir = "rounds", read = readFileSync, list = readdirSync) {
  const byBuild = new Map();
  for (const f of list(dir).filter((n) => /^\d{3}-.*\.json$/.test(n))) {
    let r;
    try {
      r = JSON.parse(read(`${dir}/${f}`, "utf8"));
    } catch {
      continue;
    }
    if (r?.salvagedFrom) continue; // never take the reference from a salvage
    const build = String(r?.build ?? "").split(" ")[0];
    const names = Object.values(r?.selftest ?? {})
      .map((s) => s?.name)
      .filter(Boolean);
    if (!build || !names.length) continue;
    // The LONGEST set seen for the build: a round of that build that itself lost
    // a scenario to a stall must not lower the bar for everything after it.
    if ((byBuild.get(build)?.length ?? 0) < names.length) byBuild.set(build, names);
  }
  return byBuild;
}

if (isMain(import.meta.url, process.argv[1])) {
  const dry = process.argv.includes("--dry-run");
  // NOT `rounds` — see the header. Defaulting to the real archive is the one
  // mistake this tool could make that nobody would notice until a drift reading
  // went wrong months later.
  const intoAt = process.argv.indexOf("--into");
  const into = intoAt === -1 ? "rounds-salvaged" : process.argv[intoAt + 1];
  if (into === "rounds") {
    console.error("  refusing --into rounds: these are old rounds and would claim the newest numbers.");
    console.error("  See this script's header, and `rounds/README.md` — the archive is oldest-first.");
    process.exit(2);
  }
  if (!dry && !existsSync(into)) mkdirSync(into, { recursive: true });
  const expected = expectedByBuild();
  const files = readdirSync("crashes")
    .filter((f) => f.endsWith("-crashed-run.json"))
    .sort();
  let filed = 0;
  const refused = [];
  for (const f of files) {
    let crash;
    try {
      crash = JSON.parse(readFileSync(`crashes/${f}`, "utf8"));
    } catch (err) {
      refused.push([f, `unreadable: ${err?.message ?? err}`]);
      continue;
    }
    const build = String(crash?.build ?? "").split(" ")[0];
    const out = roundFromCrash(crash, expected.get(build), f);
    if (!out.ok) {
      refused.push([f, out.why]);
      continue;
    }
    if (dry) {
      console.log(`  would file  ${f}  ${build}  ${out.round.slideSize.width}x${out.round.slideSize.height}`);
      filed++;
      continue;
    }
    // Through `archive`, not a hand-rolled write: it owns the numbering, the
    // twin check and the "every round ever filed" lister that stopped two
    // different rounds being filed as 064.
    const tmp = `${into}/.salvage-tmp.json`;
    writeFileSync(tmp, `${JSON.stringify(out.round, null, 2)}\n`);
    try {
      const name = archive(tmp, into);
      console.log(`  ${f}  ->  ${into}/${name}`);
      filed++;
    } catch (err) {
      refused.push([f, `archive refused it: ${err?.message ?? err}`]);
    } finally {
      if (existsSync(tmp)) unlinkSync(tmp);
    }
  }
  console.log(`\n  ${filed} salvaged, ${refused.length} refused, of ${files.length} crash record(s)`);
  // NAMED, never a bare count. A refusal here is the interesting half: it says
  // which rounds died before they had anything worth keeping.
  for (const [f, why] of refused) console.log(`    ${f}: ${why}`);
}
