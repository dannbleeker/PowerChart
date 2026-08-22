import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import * as driver from "../scripts/round.mjs";

/**
 * THE HAND-ROLLED WAIT, and why it is not hand-rolled any more.
 *
 * Every round is launched behind a "wait for Pages" loop, and until 2026-08-22
 * that loop was retyped at the shell each time. That day's copy had two silent
 * faults at once: it polled `dannbleeker.github.io`, which GitHub Pages has
 * 301'd to the custom domain since July, and it read a `.commit` field the
 * document does not have — its only field is `build`. Either fault alone prints
 * the same tick a not-yet-deployed site prints, so ten identical ticks were
 * indistinguishable from patience, and the round sat there for ten minutes
 * after its deploy had already landed.
 */
describe("waiting for the site to serve the commit under test", () => {
  const capture = () => {
    const lines: string[] = [];
    return { lines, log: (s: string) => lines.push(s) };
  };
  const sleep = async () => {};

  it("returns as soon as the served build matches, without sleeping past it", async () => {
    let polls = 0;
    const fetchBuild = async () => {
      polls++;
      return polls < 3 ? "0000000 · 2026-08-22 18:00Z" : "abc1234 · 2026-08-22 19:16Z";
    };
    const { lines, log } = capture();
    expect(await driver.waitForDeploy("abc1234", { fetchBuild, sleep, log, max: 40 })).toBe(true);
    expect(polls).toBe(3);
    expect(lines.some((l) => l.includes("deployed"))).toBe(true);
  });

  it("reads the stamp with the driver's own parser, not a field the document lacks", async () => {
    // The document is `{"build":"9b0cdc2 · 2026-08-22 19:16Z"}` — there is no
    // `.commit`, and reaching for one throws on EVERY poll including the
    // successful ones.
    const fetchBuild = async () => JSON.stringify({ build: "9b0cdc2 · 2026-08-22 19:16Z" });
    expect(await driver.waitForDeploy("9b0cdc2", { fetchBuild, sleep, log: () => {}, max: 2 })).toBe(true);
  });

  it("says a poll could not read a build stamp, rather than printing the same tick as a stale one", async () => {
    // THE WHOLE DEFECT. A wrong URL and an undeployed commit produced identical
    // output, so ten minutes of waiting looked like ten minutes of progress.
    const { lines, log } = capture();
    await driver.waitForDeploy("abc1234", { fetchBuild: async () => "", sleep, log, max: 2 });
    expect(lines.some((l) => l.includes("did not answer with a build stamp"))).toBe(true);
    expect(lines.some((l) => l.includes("0000000"))).toBe(false);

    const stale = capture();
    await driver.waitForDeploy("abc1234", {
      fetchBuild: async () => "0000000 · 2026-08-22 18:00Z",
      sleep,
      log: stale.log,
      max: 2,
    });
    expect(
      stale.lines.some((l) => l.includes("0000000")),
      "a stale poll names the hash it read",
    ).toBe(true);
    expect(stale.lines.some((l) => l.includes("did not answer with a build stamp"))).toBe(false);
  });

  it("gives up and runs anyway rather than throwing — readiness is what judges a stale site", async () => {
    const { lines, log } = capture();
    expect(await driver.waitForDeploy("abc1234", { fetchBuild: async () => "", sleep, log, max: 3 })).toBe(false);
    expect(lines.some((l) => l.includes("running anyway"))).toBe(true);
    expect(lines.filter((l) => l.includes("poll ")).length).toBe(3);
  });
});
