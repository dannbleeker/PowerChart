import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "fs";
import { spawnSync } from "child_process";
// @ts-expect-error — plain .mjs tool, no types.
import { poolEveryDraw } from "../scripts/triage.mjs";
// Its own line: the grouped-import + `@ts-expect-error` trap is documented at
// the top of `triage.test.ts` and has now bitten twice.
// @ts-expect-error — as above.
import { poolRasteriseArms } from "../scripts/triage.mjs";

/**
 * The archive is evidence, so it has to stay readable and honestly labelled.
 *
 * Not a snapshot of the findings — those live in the round files themselves and
 * in the commits that acted on them. This checks the two things that would make
 * the archive quietly wrong: a file that no longer parses as a round, and a file
 * whose name claims a build it does not carry. The second is not hypothetical.
 * Round 27 ran on `fef1c2a` while `162f80a` was the last commit merged, and the
 * first attempt at this directory named it `027-162f80a.json`. An archive that
 * misattributes a round to a build is worse than no archive: every later
 * comparison inherits the error and nothing on the machine disagrees.
 */
describe("the round archive", () => {
  const dir = new URL("../rounds/", import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, "$1");
  // `NNN-<build>.json` only, exactly as `triage.mjs` expands the directory.
  // "Every .json in rounds/" is the obvious spelling and it is wrong, because
  // `predictions.json` lives here too — it bit triage's directory expansion and
  // then bit this test within the hour, both times by counting the ledger as a
  // round. One shape, two places; if a third reader of this directory appears it
  // needs the same filter.
  const files = readdirSync(dir).filter((f) => /^\d{3}-.*\.json$/.test(f));

  it("has rounds in it", () => {
    expect(files.length, "the archive is empty — see rounds/README.md").toBeGreaterThan(0);
  });

  it("names every file for the build that round actually reported", () => {
    for (const f of files) {
      const round = JSON.parse(readFileSync(dir + f, "utf8"));
      const build = String(round.build ?? "").split(" ")[0];
      expect(build, `${f} carries no build stamp`).toMatch(/^[0-9a-f]{7}$/);
      expect(f, `${f} is named for a build it does not carry (${build})`).toContain(build);
      expect(f, `${f} does not start with a round number`).toMatch(/^\d{3}-/);
    }
  });

  it("keeps every round readable by the tools that pool them", () => {
    const logs = files.map((f) => JSON.parse(readFileSync(dir + f, "utf8")));
    for (const [i, log] of logs.entries()) {
      expect(Array.isArray(log?.trace?.entries), `${files[i]} has no trace entries`).toBe(true);
      expect(Array.isArray(log?.hostAnswers?.answers), `${files[i]} has no answer sheet`).toBe(true);
    }
    // The whole point of keeping them: the pooled view must see more than one
    // round's worth. A single round contributes 4 arm draws and ~40 total, so
    // anything at or below those numbers means the pooling silently read one
    // file — which is the state this directory exists to end.
    const arms = poolRasteriseArms(logs);
    const every = poolEveryDraw(logs);
    const armDraws = Object.values(arms.arms as Record<string, { ok: number; stall: number }>).reduce(
      (n, a) => n + a.ok + a.stall,
      0,
    );
    const allDraws = Object.values(every.after as Record<string, { ok: number; stall: number }>).reduce(
      (n, a) => n + a.ok + a.stall,
      0,
    );
    expect(arms.rounds, "pooling read fewer rounds than the archive holds").toBe(files.length);
    expect(armDraws, "the counterbalanced arms pooled no further than one round").toBeGreaterThan(4);
    expect(allDraws, "every-draw pooling saw no more than a single round").toBeGreaterThan(40);
  });

  it("reports on the NEWEST round when handed the directory", () => {
    // `npm run rounds` is the one command a loop glances at between rounds, and
    // it expanded the directory sorted then read `[0]` — the OLDEST round. So
    // the deck evidence and self-test at the top came from two days before the
    // grid underneath, with nothing saying so. Nineteen rounds went by without
    // it being noticed, because the grid was right.
    //
    // Asserted through the CLI rather than a unit, because the defect lived in
    // argument handling and a unit test of the reporter would have passed
    // throughout.
    const sorted = [...files].sort();
    const newestBuild = /^\d{3}-(.*)\.json$/.exec(sorted[sorted.length - 1])![1];
    const out = spawnSync(process.execPath, ["scripts/triage.mjs", "rounds"], { encoding: "utf8" }).stdout ?? "";
    const header = out.split("\n").find((l) => l.trimStart().startsWith("build ")) ?? "";
    expect(header, "reported some round other than the newest").toContain(newestBuild);
    // An explicitly named file still means THAT file: two rounds named on the
    // command line means the first of them, and only the directory case flips.
    const firstBuild = /^\d{3}-(.*)\.json$/.exec(sorted[0])![1];
    const named =
      spawnSync(process.execPath, ["scripts/triage.mjs", dir + sorted[0], dir + sorted[sorted.length - 1]], {
        encoding: "utf8",
      }).stdout ?? "";
    expect(named.split("\n").find((l) => l.trimStart().startsWith("build ")) ?? "").toContain(firstBuild);
  });

  it("does not carry the slide images, which are half the bytes", () => {
    for (const f of files) {
      const raw = readFileSync(dir + f, "utf8");
      // A base64 slide picture is thousands of characters. Anything that long
      // and base64-shaped means an unstripped round got committed.
      expect(/"[A-Za-z0-9+/=]{2000,}"/.test(raw), `${f} still carries an embedded image — strip it`).toBe(false);
    }
  });
});
