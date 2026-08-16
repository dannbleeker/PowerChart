#!/usr/bin/env node
/**
 * Run a night's rounds without a person between them.
 *
 *     npm run cycle
 *
 * Two rounds at 16:9 and one at 4:3, which is the schedule agreed with the repo
 * owner: 16:9 is the measurement, 4:3 is VALIDATION. The noise floor here is
 * 1-versus-5 for the same fault with nothing changed, so a single round is never
 * evidence and the pair is the unit — `docs/ROUNDS.md` says two per build
 * minimum, three where a claim depends on it.
 *
 * WHAT THIS DOES NOT DO is as important as what it does.
 *
 * It does not decide whether a round found something. It does not decide whether
 * a divergence is real. It does not set a slide size — `RECOVERABLE_STOPS`
 * excludes `wrong-size` precisely because changing the size would change what
 * the round measures rather than restore it, and a cycle runner that "helpfully"
 * fixed the profile would be doing the one thing recovery is forbidden to do.
 * And it never goes near sign-in: that needs a password, none of that is the
 * agent's to do, and the driver returns before any of this on that path.
 *
 * It also does not reimplement `shouldRetry`. Retrying is the driver's job and
 * there must be exactly one implementation of "is this worth another attempt";
 * this reads the receipt the driver leaves and stops when the driver gave up.
 */
import { spawnSync } from "child_process";
import { readFileSync, existsSync, rmSync } from "fs";
import { isMain } from "./is-main.mjs";
// Its own line: the grouped-import trap is documented in test/triage.test.ts and
// has been paid for four times.
import { RECEIPT_PATH } from "./round.mjs";

/**
 * The night's legs, in order.
 *
 * 4:3 LAST and only once. It is the cheap validation arm, and putting it first
 * would mean a night that dies early loses the measurement rather than the
 * check. `PW_EXPECT_SIZE` is what makes the driver refuse a deck in the wrong
 * profile — without it a 4:3 leg run against a 16:9 deck files a round under a
 * profile it was not measured at, and every later comparison inherits that.
 */
export function cyclePlan({ wide = "Presentation64", tall = "Presentation66" } = {}) {
  return [
    { leg: 1, deck: wide, size: "16:9", why: "the measurement" },
    { leg: 2, deck: wide, size: "16:9", why: "the pair — one round is never evidence" },
    { leg: 3, deck: tall, size: "4:3", why: "validation" },
  ];
}

/** The driver's account of the round that just ran, or null if it left none. */
export function readReceipt(path = RECEIPT_PATH, exists = existsSync, read = readFileSync) {
  if (!exists(path)) return null;
  try {
    return JSON.parse(read(path, "utf8"));
  } catch {
    return null;
  }
}

/**
 * Carry on, or stop and fetch a person?
 *
 * The three stopping conditions are different in kind and the messages have to
 * keep them apart, because the response to each is different: a build that
 * regressed wants reading, a refusal the driver could not clear wants a hand on
 * the machine, and a missing receipt means something went wrong beneath both.
 *
 * A round that FINISHED always continues, even when its scenarios failed. A
 * scenario failing is the measurement working — stopping the night on it would
 * throw away the second half of the pair, which is the only thing that could
 * tell a real fault from this host's 1-in-5 mood.
 */
export function nextStep({ exitCode, receipt, gateStatus }) {
  // TWO WAYS FOR THE GATE TO EXIT NON-ZERO, and they are opposite findings. 1 is
  // the thing it exists to say: a scenario that was passing has stopped. 2 is
  // the gate reporting that it could not judge at all — an unreadable archive,
  // an empty one — and calling that a regression would send someone hunting a
  // fall that never happened. Before the gate had its own code for this, an
  // interrupted write to a round file did exactly that: node exited 1 on a
  // SyntaxError and the night stopped blaming the build.
  if (gateStatus === 2) {
    return {
      go: false,
      why: "the gate could not judge this round — nothing was checked, and that is not a regression",
    };
  }
  if (gateStatus !== 0) {
    return {
      go: false,
      why: "a scenario that WAS passing has stopped — that is the one fatal check, and it wants reading now",
    };
  }
  if (!receipt) {
    return { go: false, why: `the driver left no ${RECEIPT_PATH} — it did not reach the end of its own run` };
  }
  if (exitCode === 0) return { go: true, why: null };
  // The driver has ALREADY retried everything recovery addresses by this point.
  // A non-zero exit carrying recoverable codes means the retries ran out, not
  // that another one would help.
  const codes = receipt.codes?.length ? ` (${receipt.codes.join(", ")})` : "";
  return {
    go: false,
    why: receipt.recoverable
      ? `the driver retried and still could not clear ${receipt.reason}${codes}`
      : `${receipt.reason}${codes} is not something recovery addresses — it needs a person`,
  };
}

if (isMain(import.meta.url, process.argv[1])) {
  const dirArg = process.argv.indexOf("--dir");
  const dir = dirArg === -1 ? ".pw-session" : process.argv[dirArg + 1];
  const retryArg = process.argv.indexOf("--retry");
  const retry = retryArg === -1 ? "6" : process.argv[retryArg + 1];
  const plan = cyclePlan({
    wide: process.env.PW_DECK_16_9 ?? undefined,
    tall: process.env.PW_DECK_4_3 ?? undefined,
  });

  console.log(`  a cycle of ${plan.length} round(s): ${plan.map((p) => p.size).join(", ")}\n`);
  let ran = 0;
  for (const leg of plan) {
    console.log(`\n  ROUND ${leg.leg} of ${plan.length} — ${leg.deck} at ${leg.size}, ${leg.why}`);
    // A CHILD PROCESS PER ROUND, not an imported function. A round that wedges
    // its own process cannot then poison the next one, and the per-round
    // environment is the natural way to say which arm this is.
    // CLEAR IT FIRST, or the check below reads the LAST round's outcome. A leg
    // that dies before writing leaves the previous leg's receipt on disk, and
    // nothing distinguishes the two — which made the "the driver left no
    // receipt" branch, written for exactly this, unreachable from leg 2 onward.
    // A stale `finished` would have been read as this round's success.
    try {
      rmSync(RECEIPT_PATH, { force: true });
    } catch {
      /* nothing to clear is the normal case on the first leg */
    }
    const r = spawnSync(process.execPath, ["scripts/round.mjs", "--dir", dir, "--retry", retry], {
      stdio: "inherit",
      env: { ...process.env, PW_DECK: leg.deck, PW_EXPECT_SIZE: leg.size },
    });
    const receipt = readReceipt();
    // The gate runs after EVERY round rather than once at the end, so a
    // regression stops the night at the round that caused it instead of being
    // found three rounds later against a changed deck.
    const gate = spawnSync(process.execPath, ["scripts/rounds-gate.mjs"], { stdio: "inherit" });
    const step = nextStep({ exitCode: r.status, receipt, gateStatus: gate.status });
    if (receipt?.roundFile) ran++;
    if (!step.go) {
      console.error(`\n  CYCLE STOPPED after ${ran} archived round(s): ${step.why}`);
      console.error("  Nothing here is fixed by running it again — see docs/ROUNDS.md.");
      process.exit(1);
    }
  }
  console.log(`\n  cycle done — ${ran} round(s) archived. The rounds are the evidence; read them.`);
  process.exit(0);
}
