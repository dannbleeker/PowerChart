#!/usr/bin/env node
/**
 * PUT ONE QUESTION TO THE LIVE HOST, FROM THE COMMAND LINE.
 *
 * `round.mjs` drives the same browser to run a fourteen-minute round. This runs
 * ONE experiment — see `src/render/experiments.ts` — and prints what the host
 * said. Seconds, not minutes, and no archive entry: an experiment settles a
 * decision rather than watching the product.
 *
 * It reuses the round driver's `cli` and `refFor` deliberately. Those two carry
 * a year of hard-won detail about finding the pane inside the right frame — the
 * miss line that can be matched as a hit, the outer iframe whose ref lands in
 * OneDrive where `PowerPoint` is undefined — and a second copy of that would be
 * wrong in a way nobody notices until an experiment reports a healthy host as
 * dead.
 *
 * Usage:  node scripts/experiment.mjs [experiment-id]
 */
import { spawnSync } from "child_process";
import { cli, refFor, recover } from "./round.mjs";

/** How long to wait for the pane to finish answering before giving up on it. */
const ANSWER_BUDGET_MS = 90_000;

/**
 * Click the button and wait for the pane's own note to settle.
 *
 * IN ONE `eval`, on purpose. The alternative is click-then-poll from out here,
 * which spends a CLI round trip per poll and races the pane's own rendering.
 * The pane already writes its result into the status line; this waits for that
 * line to stop saying "busy" and hands it back.
 */
function askScript(id, budgetMs) {
  return (
    "async () => {" +
    "  const pick = document.getElementById('experiment-pick');" +
    "  const run = document.getElementById('experiment-run');" +
    "  if (!pick || !run) return 'no-controls';" +
    `  const want = ${JSON.stringify(id)};` +
    "  if (want) {" +
    "    const has = [...pick.options].some((o) => o.value === want);" +
    "    if (!has) return 'no-such-experiment:' + [...pick.options].map((o) => o.value).join(',');" +
    "    pick.value = want;" +
    "  }" +
    // The status line the pane writes into. Captured BEFORE the click so a
    // stale result from a previous press cannot be read as this one's answer.
    // `#host-note` by id, not by a role or class guess. The pane has ONE status
    // line and it is aria-live; a selector list that merely happens to match it
    // today is how this reads some other element after a markup tidy.
    "  const el = document.getElementById('host-note');" +
    "  if (!el) return 'no-status-line';" +
    "  const line = () => el.textContent || '';" +
    "  const before = line();" +
    "  run.click();" +
    `  const deadline = Date.now() + ${budgetMs};` +
    "  while (Date.now() < deadline) {" +
    "    await new Promise((r) => setTimeout(r, 400));" +
    "    const now = line();" +
    "    if (now && now !== before && !/Asking:/.test(now)) return now;" +
    "  }" +
    "  return 'timed-out:' + line();" +
    "}"
  );
}

async function main(argv) {
  const id = argv[2] ?? "";
  const sh = cli(spawnSync, process.cwd());
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  let pane = refFor(sh, "Chart", /tab "Chart"/);
  if (!pane) {
    // OPEN IT RATHER THAN REFUSING. A tool that needs a round to have been run
    // first is a tool nobody reaches for, and this exists precisely so a
    // question can be put without one. `recover` is the round driver's own — it
    // reopens the browser, re-sideloads the add-in and opens the pane — so this
    // inherits every fix that path has had rather than growing a second one that
    // is subtly different.
    console.error("  the pane is not open — reopening the browser and the add-in first...");
    await recover(sh, sleep);
    pane = refFor(sh, "Chart", /tab "Chart"/);
  }
  if (!pane) {
    console.error(
      "  could not open the add-in pane. If a sign-in prompt is showing, that one is yours —\n" +
        "  sign in and run this again.",
    );
    return 2;
  }
  const said = sh("eval", askScript(id, ANSWER_BUDGET_MS), pane).trim();
  if (!said) {
    console.error("  the pane returned nothing — is it still loading?");
    return 2;
  }
  console.log(`\n  ${said}\n`);
  // A NON-ZERO EXIT for an answer that is not an answer, so a caller scripting
  // several of these can tell "the host said no" from "the question never ran".
  return /^(no-controls|no-status-line|no-such-experiment|timed-out)/.test(said) ? 1 : 0;
}

process.exitCode = await main(process.argv);
