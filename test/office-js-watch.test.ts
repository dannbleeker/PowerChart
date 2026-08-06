import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types, deliberately independent
// of src/ so it cannot inherit anything from the code it watches.
import { freshIssues, reportBody, KNOWN_ISSUES, WATCHED_APIS } from "../scripts/office-js-watch.mjs";

/**
 * The tracker sweep.
 *
 * Five of the defects this repo now guards against were found in a single manual
 * sweep of the office-js tracker, after months of not looking — every one of
 * them open upstream for a year or more, under code that ships. The sweep is
 * automated so the next one is not luck.
 *
 * Only the matching half is testable, and that is the half worth testing: the
 * fetch is one API call, and an agent session cannot reach that repository at
 * all. So the network is not mocked here, it is simply not involved — issues go
 * in as data, exactly as `--from` feeds them.
 */

type Issue = Record<string, unknown>;
const issue = (n: number, title: string, body = "", extra: Issue = {}): Issue => ({
  number: n,
  title,
  body,
  state: "open",
  html_url: `https://github.com/OfficeDev/office-js/issues/${n}`,
  updated_at: "2026-08-01T00:00:00Z",
  ...extra,
});

describe("watching the office-js tracker", () => {
  it("reports an issue that names an API this add-in calls", () => {
    const fresh = freshIssues([issue(9001, "PowerPoint web", "shapes.addTextBox drops the font")]);
    expect(fresh).toHaveLength(1);
    expect(fresh[0].hits).toContain("addTextBox");
    // And says WHY it matters, so the report is actionable without opening it.
    expect(fresh[0].why[0]).toBeTruthy();
  });

  it("ignores an issue about something this add-in never calls", () => {
    // The whole value is signal. A sweep that reports every PowerPoint issue is
    // a sweep nobody reads by the second week.
    expect(freshIssues([issue(9002, "Word table borders", "nothing we touch")])).toHaveLength(0);
  });

  it("never re-reports an issue already answered", () => {
    // Without this the report is the same twenty rows forever.
    expect(freshIssues([issue(2775, "text box deletes selection", "addTextBox")])).toHaveLength(0);
  });

  it("drops pull requests", () => {
    // The issues endpoint returns them too, and a PR against office-js is not a
    // defect report about it.
    const pr = issue(9003, "fix", "context.sync", { pull_request: { url: "x" } });
    expect(freshIssues([pr])).toHaveLength(0);
  });

  it("puts the most recently active first", () => {
    const older = issue(9004, "a", "addGroup", { updated_at: "2026-01-01T00:00:00Z" });
    const newer = issue(9005, "b", "addGroup", { updated_at: "2026-08-01T00:00:00Z" });
    expect(freshIssues([older, newer]).map((f: { number: number }) => f.number)).toEqual([9005, 9004]);
  });

  it("says plainly when there is nothing new, rather than staying silent", () => {
    // A quiet week and a broken sweep look identical unless the quiet week says
    // so — and the counts are what make the "nothing new" believable.
    const body = reportBody([], 300);
    expect(body).toMatch(/No office-js issue/);
    expect(body).toContain("300");
  });

  it("records what was DONE about every issue it already knows", () => {
    // The same discipline `KNOWN_DIVERGENCES` lives by. An entry that only says
    // "seen" is a to-do wearing a passing test's clothes, and the value of this
    // table is that an issue number met in a code comment can be traced to a
    // decision without re-reading the thread.
    for (const [number, why] of Object.entries(KNOWN_ISSUES) as [string, string][]) {
      expect(why.length, `#${number} is listed with no account of what was done`).toBeGreaterThan(40);
    }
  });

  it("knows about every office-js issue the codebase cites", async () => {
    // The table has to cover what the repo already refers to, or the first sweep
    // re-reports a dozen issues someone dealt with months ago — and the report
    // gets ignored on its first run, which is the only run that sets the habit.
    const { readdirSync, readFileSync, statSync } = await import("fs");
    const cited = new Set<number>();
    const walk = (dir: string) => {
      for (const name of readdirSync(dir)) {
        if (name === "node_modules" || name === ".git" || name === "dist-lib") continue;
        const path = `${dir}/${name}`;
        if (statSync(path).isDirectory()) walk(path);
        else if (/\.(ts|mjs|md)$/.test(name)) {
          const text = readFileSync(path, "utf8");
          for (const m of text.matchAll(/office-js[#/](?:issues\/)?(\d+)/g)) cited.add(Number(m[1]));
          for (const m of text.matchAll(/office-js\/issues\/(\d+)/g)) cited.add(Number(m[1]));
        }
      }
    };
    for (const dir of ["src", "docs", "scripts", "test"]) walk(dir);
    const missing = [...cited].filter((n) => !(n in KNOWN_ISSUES)).sort((a, b) => a - b);
    expect(missing, "cited in the codebase but absent from KNOWN_ISSUES").toEqual([]);
  });

  it("watches only calls this add-in actually makes", async () => {
    // A term that matches nothing here produces noise, and noise is how a weekly
    // report stops being read. Checked against the renderer itself rather than
    // against a list someone maintained by hand beside it.
    const { readFileSync } = await import("fs");
    const source = ["src/render/powerpoint.ts", "src/render/host-probe.ts"]
      .map((f) => readFileSync(f, "utf8"))
      .join("\n")
      .toLowerCase();
    for (const { term } of WATCHED_APIS as { term: string }[]) {
      // The bare member name — `shape.group` is written `.group` at the call
      // site, `fill.setImage` as `setImage`.
      const member = term.split(".").pop()!.toLowerCase();
      expect(source, `${term} is watched but never called here`).toContain(member);
    }
  });
});
