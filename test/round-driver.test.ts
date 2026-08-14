import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import { readiness, buildOf, nextRoundNumber, stripImages, cliEntry, sessionDir } from "../scripts/round.mjs";

/**
 * The driver's whole value is refusing to start a round that cannot prove
 * anything, so the refusals are what is worth testing. A round that runs on the
 * wrong build is worse than no round: it produces a file that looks like
 * evidence and is not, and nothing downstream can tell.
 */
describe("deciding whether a round is worth running", () => {
  const ready = { head: "abc1234", deployed: "abc1234", stamp: "abc1234", slides: 1, verbose: true, pictures: true };

  it("runs when the site, the pane and HEAD all agree and the deck is clean", () => {
    expect(readiness(ready).ok).toBe(true);
  });

  it("refuses when the site has not published HEAD yet", () => {
    // The one that looks harmless. Pages lags a merge by a minute or two, and a
    // round started inside that window measures the PREVIOUS build while
    // appearing to test the new one.
    const r = readiness({ ...ready, deployed: "0000000" });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toContain("Deploy Pages");
  });

  it("refuses when the pane is older than the site, and says how to clear it", () => {
    // `Cache-Control: max-age=600` on the pane HTML. Reopening the pane does not
    // clear it; the whole PowerPoint tab has to be hard-reloaded. Round 24 was
    // nearly run this way.
    const r = readiness({ ...ready, stamp: "0000000" });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toMatch(/hard-reload/);
  });

  it("refuses a deck that still holds the last round's slides", () => {
    // Not a correctness problem, a comparability one: rounds 24 and 25 differed
    // only in this, and were compared as though they did not.
    const r = readiness({ ...ready, slides: 7 });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toContain("7 slides");
  });

  it("refuses when either toggle is off, because they decide what the round can prove", () => {
    expect(readiness({ ...ready, verbose: false }).stop.join(" ")).toContain("Verbose trace");
    expect(readiness({ ...ready, pictures: false }).stop.join(" ")).toContain("empty");
  });

  it("does not refuse merely because something could not be read", () => {
    // Unknown is not the same as wrong. A missing toggle reading means the pane
    // was not open at that moment, and turning that into a hard stop would make
    // the driver unusable while telling nobody anything true.
    expect(readiness({ ...ready, verbose: null, pictures: null, slides: null }).ok).toBe(true);
  });
});

describe("talking to the browser at all", () => {
  it("refuses everything, loudly, when the CLI could not be run", () => {
    // The failure this driver shipped with: every browser read came back empty
    // and it reported the pane as closed while the pane sat open on screen.
    // Unreachable has to short-circuit, because nothing below it was measured.
    const r = readiness({
      head: "a",
      deployed: "a",
      stamp: null,
      slides: null,
      verbose: null,
      pictures: null,
      reachable: false,
    });
    expect(r.ok).toBe(false);
    expect(r.stop).toHaveLength(1);
    expect(r.stop[0]).toContain("nothing below was actually measured");
  });

  it("finds the CLI entry beside node, or says it cannot", () => {
    expect(cliEntry("C:/n/node.exe", () => true)).toContain("playwright-cli.js");
    expect(cliEntry("C:/n/node.exe", () => false)).toBe(null);
  });

  it("normalises a session directory, because 8.3 short names are a different session", () => {
    // "C:/Users/DANN~1.PED/..." and "C:/Users/dann.pedersen/..." are one folder
    // and two strings, and the daemon keys its sessions by the string — so from
    // the short-name form it answers "(no browsers)" with the browser open.
    expect(sessionDir("C:/Users/DANN~1.PED/x", () => "C:/Users/dann.pedersen/x")).toBe("C:/Users/dann.pedersen/x");
  });

  it("falls back to the path it was given when it cannot be resolved", () => {
    expect(
      sessionDir("/nope", () => {
        throw new Error("ENOENT");
      }),
    ).toBe("/nope");
  });
});

describe("archive housekeeping", () => {
  it("reads a build out of a stamp, and nothing out of prose", () => {
    expect(buildOf("32a6987 · 2026-08-14 08:03Z")).toBe("32a6987");
    expect(buildOf("no build here")).toBe(null);
  });

  it("numbers the next round above the highest already kept", () => {
    expect(nextRoundNumber(["023-0e22b31.json", "028-32a6987.json", "README.md"])).toBe("029");
    expect(nextRoundNumber([])).toBe("001");
  });

  it("takes the highest number, not the last one listed", () => {
    // `readdirSync` order is not guaranteed, and taking the last entry would
    // renumber a round on top of an existing one the first time it differed.
    expect(nextRoundNumber(["028-32a6987.json", "023-0e22b31.json"])).toBe("029");
  });

  it("strips slide images but leaves the evidence alone", () => {
    const round = {
      build: "abc1234",
      deck: { picture: "A".repeat(3000) },
      hostAnswers: { answers: [{ id: "a", answer: "yes", detail: "short" }] },
    };
    const out = stripImages(structuredClone(round));
    expect(out.deck.picture).toMatch(/^<image stripped/);
    expect(out.hostAnswers.answers[0].detail, "a short string is not an image").toBe("short");
  });
});
