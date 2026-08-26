#!/usr/bin/env node
/**
 * WAIT FOR *THIS* COMMIT TO BE LIVE — CI green, then Pages deployed.
 *
 * Written because I got it wrong four times in one afternoon, the same way each
 * time: `gh run list ... | select(.name == "Deploy Pages") | first` picks the
 * NEWEST run, which right after a push is still the PREVIOUS commit's. Watching
 * that one finish reports success for a deploy that never happened, and the
 * round then runs against stale code — or refuses, which is how it was finally
 * caught: `the site is serving 55630e7 but HEAD is acb5a9e`.
 *
 * That is worse than not checking at all. A missing check leaves you unsure; a
 * check keyed to the wrong row manufactures evidence.
 *
 * CI FIRST, though NOT for the reason I first wrote here. I claimed Pages does
 * not deploy a failed build; the archive says otherwise — `acb5a9e` deployed
 * successfully with CI red. Deploys and CI are independent.
 *
 * The ordering is still right, for a different reason: a build that is live but
 * failing CI is a worse thing to measure than one that has not landed. A round
 * against it produces numbers everyone will trust and a suite nobody has
 * satisfied. So CI is a gate on USING the build, not a precondition of its
 * existence — and when it is red this says so instead of leaving the reader to
 * wonder about the deploy.
 *
 * Usage:  node scripts/await-deploy.mjs [--timeout-min 20]
 * Exit:   0 live · 1 CI failed · 2 deploy failed · 3 timed out
 */
import { execFileSync } from "child_process";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** The workflow run for THIS commit, or null while GitHub has not made one yet. */
function runFor(name, sha) {
  const out = execFileSync(
    "gh",
    ["run", "list", "--limit", "20", "--json", "databaseId,name,headSha,status,conclusion"],
    { encoding: "utf8" },
  );
  const rows = JSON.parse(out);
  return rows.find((r) => r.name === name && r.headSha === sha) ?? null;
}

async function waitFor(name, sha, deadline) {
  // Absent is NOT failed. GitHub takes a few seconds to create the run, and
  // treating "no row yet" as a failure would make this flap on every push.
  let lastSeen = "(not created yet)";
  while (Date.now() < deadline) {
    const row = runFor(name, sha);
    if (row) {
      lastSeen = `${row.status} ${row.conclusion ?? ""}`.trim();
      if (row.status === "completed") return row.conclusion === "success" ? "success" : `failed: ${row.conclusion}`;
    }
    process.stdout.write(`  ${name}: ${lastSeen}\r`);
    await sleep(15_000);
  }
  return `timed out (last seen: ${lastSeen})`;
}

const argv = process.argv.slice(2);
const at = argv.indexOf("--timeout-min");
const minutes = at === -1 ? 20 : Number(argv[at + 1]) || 20;
const sha = execFileSync("git", ["rev-parse", "HEAD"], { encoding: "utf8" }).trim();
const short = sha.slice(0, 7);
const deadline = Date.now() + minutes * 60_000;

console.log(`  waiting for ${short} to go live`);
const ci = await waitFor("CI", sha, deadline);
console.log(`  CI: ${ci}`.padEnd(60));
if (ci !== "success") {
  // NAMED, because the round that follows would otherwise refuse with
  // `site-behind` and send the reader looking at the deploy instead.
  console.error(`  ${short} did not pass CI. It may well be LIVE — deploys do not wait for CI — but a round`);
  console.error("  against a build the suite rejects produces numbers people will quote and tests nobody has passed.");
  process.exit(ci.startsWith("timed out") ? 3 : 1);
}
const pages = await waitFor("Deploy Pages", sha, deadline);
console.log(`  Deploy Pages: ${pages}`.padEnd(60));
if (pages !== "success") process.exit(pages.startsWith("timed out") ? 3 : 2);
console.log(`  ${short} is live`);
