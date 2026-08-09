import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { execFileSync } from "child_process";

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

  it("can be run by hand when the automatic trigger does not fire", () => {
    // Not hypothetical. On 2026-08-06 the `pull_request` event produced no CI
    // run at all for PR #300 — CodeQL ran on the same event seconds later, and
    // this workflow was `active` and unmodified; closing and reopening the PR
    // did not fire it either. `test` is the only required check, so a missed
    // trigger blocks the merge outright, and without a dispatch the only lever
    // left is pushing a commit nobody wanted.
    const ci = readFileSync(".github/workflows/ci.yml", "utf8");
    expect(ci, "CI has no manual trigger — a missed event would be unrecoverable").toMatch(/workflow_dispatch:/);
  });

  it("never cancels a Pages deploy that is already running", () => {
    // `cancel-in-progress: true` would mean a second merge kills the first
    // merge's deployment mid-flight, which is how a green pipeline ends up
    // serving neither commit.
    expect(pages).toMatch(/concurrency:/);
    expect(pages, "a queued Pages deploy may not cancel the one in flight").toMatch(/cancel-in-progress:\s*false/);
  });
});

describe("the source tree stays searchable", () => {
  /**
   * One NUL byte makes grep and ripgrep classify a whole FILE as binary.
   *
   * `trace.ts` carried one for months — a deliberate separator, `${scope}\0
   * ${message}`, joining a Map key. The property it wanted was right: a
   * character that cannot appear in either half. The character was wrong, and
   * the cost was not the key: every codebase search silently skipped the file,
   * printing "binary file matches" instead of a line. The step-line formatter's
   * drop rule lived in that file and was wrong the whole time; a sweep for it
   * would never have matched.
   *
   * A file nothing can search is a file nothing will fix, so this is checked
   * rather than remembered.
   */
  it("has no NUL byte in any tracked source file", () => {
    const files = execFileSync("git", ["ls-files", "*.ts", "*.mjs", "*.js", "*.json", "*.md", "*.html", "*.css"], {
      encoding: "utf8",
      maxBuffer: 8 << 20,
    })
      .split("\n")
      .filter(Boolean);
    const withNul = files.filter((f) => readFileSync(f).includes(0));
    expect(withNul, "grep reports these as binary and skips them entirely").toEqual([]);
    // The scan is only worth anything if it actually read the tree.
    expect(files.length, "the file list came back empty, so nothing was checked").toBeGreaterThan(100);
  });
});
