import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";

/**
 * The CI configuration is code, and one line of it took the site down for an
 * afternoon.
 *
 * Nothing else in this suite reads `.github/workflows`, so every default in
 * them is unexamined until it costs something. On 2026-08-06 one did: six
 * consecutive Pages deployments were cancelled by `actions/deploy-pages@v4`'s
 * ten-minute default timeout while GitHub was still reporting
 * `deployment_in_progress`, and the site served an eight-hour-old build through
 * four merged pull requests.
 *
 * These are not tests of GitHub. They pin the handful of workflow settings this
 * repo has had to learn the hard way, so the next person to touch the file has
 * to do it deliberately.
 */
describe("the Pages deploy workflow", () => {
  const pages = readFileSync(".github/workflows/pages.yml", "utf8");

  it("waits far longer than the action's default before giving up", () => {
    // The failure mode is worse than "the deploy is slow". On timeout the
    // action CANCELS the deployment GitHub is still working on, so a slow Pages
    // backend does not cost one deploy — it costs every deploy, and each re-run
    // restarts the same race. Five re-runs produced five cancellations.
    const match = pages.match(/timeout:\s*(\d+)/);
    expect(match, "the deploy step accepts the 10-minute default timeout again").toBeTruthy();
    const ms = Number(match![1]);
    // Strictly greater than the default, not merely present: a `timeout: 600000`
    // would satisfy "is set" while changing nothing at all.
    expect(ms, `timeout is ${ms}ms — the default that cancelled six deployments is 600000`).toBeGreaterThan(600_000);
  });

  it("never cancels a Pages deploy that is already running", () => {
    // `cancel-in-progress: true` would mean a second merge kills the first
    // merge's deployment mid-flight, which is how a green pipeline ends up
    // serving neither commit.
    expect(pages).toMatch(/concurrency:/);
    expect(pages, "a queued Pages deploy may not cancel the one in flight").toMatch(/cancel-in-progress:\s*false/);
  });
});
