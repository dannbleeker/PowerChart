import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import * as driver from "../scripts/round.mjs";
// A NAMESPACE import, and destructured below, because the directive only covers
// the line that carries `from`. A named-import list this long is wrapped by
// prettier onto ten lines, which moves `from` five lines down, silently uncovers
// the import and turns the directive itself into an "unused directive" error.
// This is the grouped-import trap in `triage.test.ts` reached from the other
// side: there the directive covered too much, here formatting moved it off.
const {
  readiness,
  buildOf,
  nextRoundNumber,
  stripImages,
  cliEntry,
  sessionDir,
  pingScript,
  readPing,
  refFor,
  sawCrashDialog,
} = driver;

const READY = { head: "abc1234", deployed: "abc1234", stamp: "abc1234", slides: 1, verbose: true, pictures: true };

/**
 * The driver's whole value is refusing to start a round that cannot prove
 * anything, so the refusals are what is worth testing. A round that runs on the
 * wrong build is worse than no round: it produces a file that looks like
 * evidence and is not, and nothing downstream can tell.
 */
describe("deciding whether a round is worth running", () => {
  const ready = READY;

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

describe("asking the host whether it is awake", () => {
  /**
   * The ping is a STRING sent to another JavaScript engine, so nothing the
   * compiler knows reaches it. These run the generated source against a stub
   * `PowerPoint` — the only way to find out whether what the driver sends is
   * the thing it means to send.
   */
  const runPing = (budgetMs: number, host: unknown) =>
    new Function("PowerPoint", `return (${pingScript(budgetMs)})()`)(host) as Promise<string>;

  it("comes back ok when the host answers", async () => {
    const ctx = { presentation: { slides: { getCount: () => {} } }, sync: async () => {} };
    expect(await runPing(2000, { run: async (cb: (c: unknown) => unknown) => cb(ctx) })).toMatch(/^ok:\d+$/);
  });

  it("comes back, not hangs, when the host never answers", async () => {
    // The whole point. A host that will not answer `getCount()` is exactly the
    // state rounds 24, 25 and 29 each spent most of an hour inside, and an
    // unraced `PowerPoint.run` would sit here as long as it sat there.
    //
    // Verdict only, never the number: Windows' clock quantises to ~15.6ms, so
    // asserting elapsed milliseconds would flake for reasons unrelated to this.
    const out = await runPing(20, { run: () => new Promise(() => {}) });
    expect(out).toMatch(/^no:\d+$/);
  });

  it("comes back when the host throws instead of hanging", async () => {
    // A 5010 or a torn-down proxy is a NO, not a crash in the driver.
    expect(
      await runPing(2000, {
        run: async () => {
          throw new Error("GeneralException");
        },
      }),
    ).toMatch(/^no:\d+$/);
  });

  it("takes the ref off the matching line, not the first one in the frame tree", () => {
    // What made the first live ping lie. `find` prints the frame hierarchy above
    // the hit, so the first ref in its output belongs to the OUTER iframe — the
    // OneDrive document, where `Office` and `PowerPoint` are both undefined.
    // Evaluating there returns "no" for every host, healthy or not, which is the
    // worst possible failure for a precondition: it refuses correct work and
    // teaches the reader to ignore it.
    const found = [
      "iframe [ref=f1a2b3c]",
      "  iframe [ref=f4d5e6f]",
      '    checkbox "Verbose trace" [checked] [ref=f7a8b9c]',
    ].join("\n");
    const sh = (() => found) as unknown as Parameters<typeof refFor>[0];
    expect(refFor(sh, "Verbose trace", /checkbox "Verbose trace"/)).toBe("f7a8b9c");
  });

  it("has no ref when nothing matched", () => {
    const sh = (() => "iframe [ref=f1a2b3c]") as unknown as Parameters<typeof refFor>[0];
    expect(refFor(sh, "Verbose trace", /checkbox "Verbose trace"/)).toBe(null);
  });

  it("reads a verdict back out of the CLI's quoting, and nothing out of noise", () => {
    expect(readPing('"ok:37"')).toEqual({ answered: true, ms: 37 });
    expect(readPing("no:8001")).toEqual({ answered: false, ms: 8001 });
    expect(readPing(""), "an empty result is not a dead host — see `cli`").toBe(null);
    expect(readPing("undefined")).toBe(null);
  });

  it("names the crash when PowerPoint's own dialog is up", () => {
    // The wedge, finally identified. Four rounds died against it and the report
    // they got was a 90-second timeout on whatever call came next; what had
    // actually happened is that PowerPoint crashed and put up a modal, behind
    // which every Office.js call — including an empty sync — hangs forever
    // without throwing. Saying so is the difference between a refusal the reader
    // can act on and one they wait out.
    const r = readiness({ ...READY, crashed: true, ping: { answered: false, ms: 8002 } });
    expect(r.ok).toBe(false);
    expect(r.stop[0], "the crash outranks the silence it causes").toContain("Refresh");
    expect(r.stop.join(" ")).toContain("Add-ins");
  });

  it("does not see a dialog in find's own echo of the query", () => {
    // `find` answers a miss with `No matches found for "<query>"` — the query
    // included. A detector that tested for the phrase it searched for would
    // report a crash on every healthy host forever. Same shape as the ref
    // `buildOf` used to read out of its own haystack; that one cost a round.
    expect(sawCrashDialog('No matches found for "Sorry, we ran into a problem".')).toBe(false);
    // The echo with the word `dialog` INSIDE the query, which is what a
    // contains-the-word check gets wrong. The first version of this test used a
    // query without it, so it passed against a detector that had no guard at
    // all — green, and proving nothing.
    expect(sawCrashDialog('No matches found for "crash dialog".')).toBe(false);
    expect(sawCrashDialog(""), "unreachable CLI is not a crash — see `reachable`").toBe(false);
    expect(sawCrashDialog('- dialog [ref=f21e737]:\n  - button "Refresh" [ref=f21e750]')).toBe(true);
  });

  it("refuses a round when the host will not answer the cheapest call there is", () => {
    const r = readiness({ ...READY, ping: { answered: false, ms: 8001 } });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toContain("8001ms");
  });

  it("runs when the host answers, and when nobody asked it", () => {
    // Not asked is not unwell. The pane can be closed at check time for reasons
    // the driver already reports separately, and a second hard stop built on a
    // reading that was never taken would just be noise.
    expect(readiness({ ...READY, ping: { answered: true, ms: 120 } }).ok).toBe(true);
    expect(readiness({ ...READY, ping: null }).ok).toBe(true);
  });
});

describe("archive housekeeping", () => {
  it("reads a build out of a stamp, and nothing out of prose", () => {
    expect(buildOf("32a6987 · 2026-08-14 08:03Z")).toBe("32a6987");
    expect(buildOf("no build here")).toBe(null);
  });

  it("does not mistake a playwright ref for a build", () => {
    // `f14e735` is an element ref, and the accessibility dump this reads is full
    // of them. Taking a bare 7-hex token reported the pane as showing a build
    // the site had never served and refused to start a round on a pane that was
    // showing exactly the right commit.
    expect(buildOf('button "Insert chart" [ref=f14e735]')).toBe(null);
    expect(buildOf("generic [ref=f9e12]: 2601703 · 2026-08-14 13:58Z")).toBe("2601703");
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
