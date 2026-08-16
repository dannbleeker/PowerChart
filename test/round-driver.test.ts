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
  quietStreak,
  archive,
  signedOut,
  signInIsPopup,
  shouldRetry,
  recover,
  cli,
  isOverflow,
  recoveryFor,
  noBrowser,
  slideResolveScript,
  readSlideResolve,
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
    expect(readiness({ ...ready, slideOk: true }).ok, "a resolved slide changes nothing").toBe(true);
  });

  it("refuses a host that answers the cheap call but will not resolve a slide", () => {
    // THE 2s CRASH, moved to where it is cheap. Four rounds running, PowerPoint
    // crashed two seconds into the FIRST attempt and never into one that followed
    // a recovery, and its own log named the call: `OnServerFindSucceeded could
    // not find target slide`. The ping cannot see that state — `getCount` is a
    // count, and this host answers it in single-digit ms while in it.
    const r = readiness({ ...ready, ping: { answered: true, ms: 3 }, slideOk: false });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toContain("would not resolve slide 1");
    // And it must not fire on the state the ping already covers, or the two
    // messages contradict each other on the same sheet.
    expect(readiness({ ...ready, ping: { answered: false, ms: 8011 }, slideOk: null }).stop).toHaveLength(1);
  });

  it("asks for a slide by position, inside a budget", () => {
    const script = slideResolveScript(20000);
    expect(script, "the call the host dies on, not a count").toContain("getItemAt(0)");
    expect(script, "no budget means a wedged host hangs the check forever").toContain("20000");
    expect(readSlideResolve("slide:287#62081387")).toBe(true);
    expect(readSlideResolve("slide-failed:GeneralException")).toBe(false);
    // Silence is not a refusal: an empty answer means the eval never landed, and
    // calling that a refused slide would refuse rounds on a healthy host.
    expect(readSlideResolve("")).toBe(null);
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

  it("tells a dead browser from a closed pane", () => {
    // SEVEN ATTEMPTS, 2026-08-15. A round wedged and the browser process died
    // with it; `recover` then reloaded and reopened a pane inside a window that
    // did not exist, seven times, while the check said "could not read the pane's
    // build stamp — is the add-in open?" The add-in was fine. There was nothing
    // to open it in.
    expect(noBrowser("  (no browsers)")).toBe(true);
    expect(noBrowser("### Browsers\n- ms:\n  - status: open")).toBe(false);
    expect(noBrowser("")).toBe(false);

    const r = readiness({ ...READY, browserGone: true, stamp: null, slides: null });
    expect(r.ok).toBe(false);
    expect(r.stop).toHaveLength(1);
    expect(r.stop[0], "sent the reader to the add-in again").not.toContain("is the add-in open?");
    expect(r.stop[0]).toContain("no browser");
    // Recoverable WITHOUT a password, because the profile keeps the sign-in —
    // so the loop reopens it rather than waiting for a person.
    expect(r.codes).toEqual(["browser-gone"]);
    expect(shouldRetry("not-ready", 0, 3, r.codes), "a dead browser is the loop's to fix").toBe(true);
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

  it("keeps a failed spawn's doubt inside the sweep that saw it", () => {
    // ROUND 044. `--retry` cleared the host's crash, `recover` reloaded the tab
    // and reopened the pane, and one of its spawns lost a race. The latch was
    // set for the rest of the PROCESS, so the next attempt read a healthy setup
    // — the host answered in 4ms, on the line above — and refused it with
    // "nothing below was actually measured", which was false of every value it
    // had just printed. One flaky spawn cost the whole round.
    let fail = true;
    const run = () => {
      if (fail) return { error: new Error("EAGAIN") };
      return { status: 0, stdout: "ok" };
    };
    // THE ENTRY IS INJECTED, and leaving it to the default is what turned this
    // test green here and red on CI. `cliEntry()` finds the real playwright-cli
    // on the machine that runs rounds and finds nothing on a CI runner, where
    // every call then takes the no-entry path and latches whatever the fake
    // `run` was going to do. A test about spawn results must not depend on
    // whether a tool is installed.
    const sh = cli(run, ".", "C:/fake/playwright-cli.js");
    // Sticky WITHIN a sweep is deliberate and stays: a call that never ran and a
    // page that answered nothing are the same empty string.
    sh("find", "x");
    expect(sh.state.unreachable).toBe(true);
    fail = false;
    sh("find", "x");
    expect(sh.state.unreachable, "a later success does not retro-clear this sweep").toBe(true);
    // A new sweep is a new question.
    sh.startSweep();
    expect(sh.state.unreachable).toBe(false);
    sh("find", "x");
    expect(sh.state.unreachable, "a sweep whose calls all ran is reachable").toBe(false);
  });

  it("still reports unreachable in every sweep when the tool is not installed at all", () => {
    // The permanent case the latch was written for. With no entry no spawn is
    // attempted at all, and `startSweep` must not paper over it.
    const sh = cli(
      () => {
        throw new Error("should never spawn");
      },
      ".",
      null,
    );
    sh.startSweep();
    sh("find", "x");
    expect(sh.state.unreachable).toBe(true);
  });

  it("does not call an answer too big to hold a tool that could not be run", () => {
    // ROUND 044's ROOT CAUSE. `requests` on a live PowerPoint tab is bigger than
    // spawnSync's 1 MiB default, so it came back ENOBUFS — an error, on a call
    // that ran perfectly. Read as "playwright-cli could not be run", it emptied
    // two crash reports and then refused the retry on a healthy host.
    const enobufs = Object.assign(new Error("spawnSync C:/node.exe ENOBUFS"), { code: "ENOBUFS" });
    expect(isOverflow(enobufs), "the code").toBe(true);
    expect(isOverflow(new Error("spawnSync C:/node.exe ENOBUFS")), "the message, when there is no code").toBe(true);
    expect(isOverflow(new Error("spawn ENOENT")), "a tool that is not there").toBe(false);
    expect(isOverflow(undefined)).toBe(false);

    const sh = cli(() => ({ error: enobufs }), ".", "C:/fake/playwright-cli.js");
    sh("requests");
    expect(sh.state.unreachable, "an overflow is not an unreachable tool").toBe(false);
    expect(sh.state.overflowed).toBe("requests");
    expect(sh.state.lastError).toBe("overflow");
  });

  it("asks for a buffer big enough to hold a live tab's request log", () => {
    // The fix for the cause rather than the symptom. Pinned because the default
    // is invisible: nothing about `spawnSync(...)` says 1 MiB, and the failure it
    // produces names the executable rather than the size.
    let opts: { maxBuffer?: number } = {};
    const sh = cli(
      (_file: string, _args: string[], o: { maxBuffer?: number }) => {
        opts = o;
        return { status: 0, stdout: "" };
      },
      ".",
      "C:/fake/playwright-cli.js",
    );
    sh("requests");
    expect(opts.maxBuffer).toBeGreaterThanOrEqual(64e6);
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

  it("does not end a running round because the CLI lost a race", () => {
    // What killed round 29's first archive-worthy run: the CLI serves one
    // command per session, an agent read the trace while the driver polled, one
    // poll exited non-zero, and the empty string that came back was reported as
    // "PowerPoint has probably crashed". The round was fine — it went on to pass
    // 10 of 12 scenarios. A failed call is not a silent page.
    expect(quietStreak(1, "", true), "a failed call measured nothing").toBe(0);
    expect(quietStreak(1, '- button "Download run log" [disabled]', false)).toBe(0);
    // Genuinely silent, twice, is the real thing — and it has to reach two.
    expect(quietStreak(0, "", false)).toBe(1);
    expect(quietStreak(1, "   ", false)).toBe(2);
  });

  it("retries what recovery can actually put right", () => {
    // A crash is the state the driver has always put right on its own, six times
    // in one night. The rest of this is 2026-08-15, when retrying ONLY a crash
    // cost two rounds: mid-round the quiet wedge exits as `silent` and the driver
    // went home, though `docs/ROUNDS.md` says a reload clears both forms of the
    // wedge and `recover` already does exactly that.
    expect(shouldRetry("crashed", 0, 3)).toBe(true);
    expect(shouldRetry("crashed", 3, 3), "the budget is a budget").toBe(false);
    expect(shouldRetry("silent", 0, 3), "the quiet wedge is a wedge").toBe(true);
    // A WEDGE MID-ROUND, which stopped an unattended run dead in its first hour
    // on 2026-08-15: the round timed out at thirty minutes, `recover` could have
    // cleared it in eighty seconds, and `timeout` was not on the list. The cost
    // of retrying is bounded by the `--retry` the caller chose; the cost of not
    // retrying was nine idle hours.
    expect(shouldRetry("timeout", 0, 3), "a wedged round is recoverable").toBe(true);
    expect(shouldRetry("timeout", 3, 3), "and still bounded by the budget").toBe(false);
    expect(shouldRetry("finished", 0, 3)).toBe(false);
  });

  it("retries a check whose every refusal a reload would clear, and no other", () => {
    // The hand-recovery of 2026-08-15, encoded. `--check` refused with a silent
    // host, a pane one build behind and an eight-slide deck; a person then did
    // the five steps `recover` does and every one of them went green. The driver
    // had the same five steps and would not take them.
    expect(shouldRetry("not-ready", 0, 3, ["host-silent", "pane-stale", "deck-dirty"])).toBe(true);
    expect(shouldRetry("not-ready", 0, 3, ["crashed"])).toBe(true);
    expect(shouldRetry("not-ready", 0, 3, ["slide-refused"])).toBe(true);

    // And the refusals a reload does NOT touch, which is what the old comment
    // was right about: retrying a stale build just repeats itself until the
    // night is gone. ONE unrecoverable code is enough to stop, even beside three
    // recoverable ones.
    expect(shouldRetry("not-ready", 0, 3, ["site-behind"]), "waiting on Pages is not a reload").toBe(false);
    expect(shouldRetry("not-ready", 0, 3, ["verbose-off"]), "a pane toggle is a person's choice").toBe(false);
    expect(shouldRetry("not-ready", 0, 3, ["host-silent", "deck-dirty", "site-behind"])).toBe(false);

    // No codes is not a licence: it means nothing was recorded about why the
    // check refused, and retrying on no evidence is how a loop spins.
    expect(shouldRetry("not-ready", 0, 3, []), "an empty list is not a reason to retry").toBe(false);
    expect(shouldRetry("not-ready", 0, 3), "and neither is no list at all").toBe(false);
    expect(shouldRetry("not-ready", 3, 3, ["deck-dirty"]), "the budget still binds").toBe(false);
  });

  it("names what it is recovering from, rather than announcing a crash that did not happen", () => {
    // ROUND 047 refused on a dirty deck alone and was told "clearing the crash
    // and starting again". The line was written when a crash was the only thing
    // retried; the moment the retry covered more, it began inventing one. A
    // recovery line is read while someone is debugging.
    expect(recoveryFor("crashed", ["crashed"])).toContain("clearing the crash");
    expect(recoveryFor("silent", undefined)).toContain("went quiet");
    expect(recoveryFor("timeout", undefined)).toContain("wedged");
    expect(recoveryFor("not-ready", ["deck-dirty"]), "the round-047 line").toBe(
      "recovering from a dirty deck, then starting again",
    );
    expect(recoveryFor("not-ready", ["host-silent", "pane-stale", "deck-dirty"])).toBe(
      "recovering from a silent host and a stale pane and a dirty deck, then starting again",
    );
    // An unrecoverable code never reaches this line (shouldRetry stops first),
    // so it must not be named as something being recovered from.
    expect(recoveryFor("not-ready", ["site-behind"])).toBe("recovering and starting again");
    expect(recoveryFor("not-ready", [])).toBe("recovering and starting again");
  });

  it("gives every refusal a code, and every code a meaning", () => {
    // The codes are what the retry decision reads, so a stop that forgets one is
    // a stop the driver will treat as no reason at all. Asserted against the
    // real function rather than a list, because a hand-kept list is the thing
    // that goes stale.
    const wedged = readiness({ ...READY, ping: { answered: false, ms: 8011 }, slides: 8, stamp: "old1234" });
    expect(wedged.ok).toBe(false);
    expect(wedged.codes).toEqual(["host-silent", "pane-stale", "deck-dirty"]);
    expect(wedged.codes.length, "a message without a code is invisible to the retry").toBe(wedged.stop.length);
    expect(shouldRetry("not-ready", 0, 3, wedged.codes), "the state a person fixed by hand").toBe(true);

    const stale = readiness({ ...READY, head: "abc1234", deployed: "def5678", stamp: "def5678" });
    expect(stale.codes).toEqual(["site-behind"]);
    expect(shouldRetry("not-ready", 0, 3, stale.codes)).toBe(false);
  });

  it("recovers in the order the host needs, not the order that reads well", async () => {
    // The sequence is load-bearing: Refresh reloads the document and closes the
    // pane, so the pane must be reopened AFTER it, the Automation tab must be
    // showing before the run button exists, and the deck has to be cleaned last
    // because cleaning needs the pane's frame to evaluate in.
    const calls: string[] = [];
    const sh = ((...args: string[]) => {
      calls.push(args[0] + (args[1] ? ` ${args[1].slice(0, 22)}` : ""));
      if (args[0] === "find" && args[1] === "Sorry, we ran into a problem") return "- dialog [ref=f1]:";
      if (args[0] === "find") return `button "${args[1]}" [ref=f2]\ntab "${args[1]}" [ref=f2]`;
      return "ok";
    }) as unknown as Parameters<typeof recover>[0];

    await recover(sh, async () => {});

    expect(calls.filter((c) => c.startsWith("eval")).length, "clicked and cleaned").toBeGreaterThan(2);
    const order = calls.join(" | ");
    expect(order.indexOf("find Insert chart"), "the pane is reopened after the reload").toBeGreaterThan(
      order.indexOf("find Sorry"),
    );
    expect(order.indexOf("find Automation")).toBeGreaterThan(order.indexOf("find Insert chart"));
  });

  it("reloads when there is no dialog to clear", async () => {
    // The quiet form of the wedge — no modal, session severed by the network. A
    // recovery that only knew how to click Refresh would do nothing at all here.
    const calls: string[] = [];
    const sh = ((...args: string[]) => {
      calls.push(args[0]);
      if (args[0] === "find" && args[1] === "Sorry, we ran into a problem") return 'No matches found for "…".';
      return `button "x" [ref=f2]`;
    }) as unknown as Parameters<typeof recover>[0];

    await recover(sh, async () => {});
    expect(calls).toContain("reload");
  });

  it("says SIGNED OUT rather than blaming a pane that was never there", async () => {
    // The state the owner walks back into after the browser dies. A signed-out
    // browser answers every pane read with nothing, which looks exactly like an
    // add-in nobody opened — and the check duly said "is the add-in open?" while
    // the tab was showing login.live.com, sending the reader to hunt for a pane
    // in a window with no document in it.
    //
    // It also short-circuits, for the same reason `reachable` does: if this is
    // true then nothing below it was measured and every other line is noise.
    const r = readiness({ ...READY, stamp: null, slides: null, loggedOut: true });
    expect(r.ok).toBe(false);
    expect(r.stop, "said more than the one thing that is true").toHaveLength(1);
    expect(r.stop[0]).toContain("sign-in page");
    expect(r.stop[0], "the reader has to know this one is theirs").toContain("needs a password");
  });

  it("knows a sign-in page from a document", () => {
    expect(signedOut("- 0: (current) [Sign in](https://login.live.com/login.srf?wa=wsignin1)")).toBe(true);
    expect(signedOut("- 0: (current) [x](https://login.microsoftonline.com/common/oauth2)")).toBe(true);
    expect(signedOut("- 0: (current) [Presentation63.pptx](https://onedrive.live.com/personal/x/doc.aspx)")).toBe(
      false,
    );
    expect(signedOut(""), "no tabs is not a sign-in page").toBe(false);
  });

  it("knows a sign-in POPUP beside a live deck from a browser sitting on a login page", () => {
    // What actually ended the overnight run of 2026-08-15/16. Office opened an
    // auth prompt NEXT TO a deck tab that was still there, and the driver
    // reported "the browser is on a Microsoft sign-in page" — while the screen
    // showed PowerPoint with a small dialog over it.
    //
    // Stopping the round is right either way: if the host is asking for
    // credentials, nothing measured past that point can be trusted. What was
    // wrong is a message that does not match what the reader sees, which on an
    // overnight run is the only account of the night they get.
    const popup = [
      "- 0: [Presentation63.pptx](https://onedrive.live.com/personal/x/doc.aspx)",
      "- 1: (current) [Sign in](https://login.live.com/login.srf?wa=wsignin1)",
    ].join("\n");
    expect(signInIsPopup(popup), "a deck tab was still open and this called it a login page").toBe(true);
    // The whole browser on a login page — no document tab anywhere.
    expect(signInIsPopup("- 0: (current) [Sign in](https://login.live.com/login.srf)")).toBe(false);
    // And it must not fire on a healthy deck with no prompt at all.
    expect(signInIsPopup("- 0: (current) [Presentation63.pptx](https://onedrive.live.com/x)")).toBe(false);
    expect(signInIsPopup("")).toBe(false);

    // The messages are genuinely different, and each names what the reader can
    // see. Both hand the job back — it needs a password either way.
    const asPopup = readiness({ ...READY, stamp: null, slides: null, loggedOut: true, authPopup: true });
    const asPage = readiness({ ...READY, stamp: null, slides: null, loggedOut: true });
    expect(asPopup.stop[0], "the popup case still described a browser on a login page").toContain("beside the deck");
    expect(asPage.stop[0]).toContain("sign-in page");
    expect(asPopup.stop[0]).not.toBe(asPage.stop[0]);
    for (const r of [asPopup, asPage]) {
      expect(r.ok, "a host asking for credentials cannot be measured either way").toBe(false);
      expect(r.stop[0], "the reader has to know this one is theirs").toContain("needs a password");
    }
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
  it("refuses to file the same log twice under a new number", async () => {
    // Round 042 wedged, `Download run log` was disabled so the click did
    // nothing, and the PREVIOUS round's file was still sitting at the same path.
    // It was archived as 039 and was byte-identical to 038 — a whole round of
    // evidence that never happened, caught by a checksum and nothing else.
    //
    // A fabricated round is the worst thing that directory can hold. Everything
    // downstream pools these files — verdict histories, the rasterise arms, the
    // scenario flip detector — so one duplicate silently doubles the weight of
    // whatever the real round happened to say.
    const round = { build: "abc1234 · 2026-08-15 05:00Z", trace: { entries: [] } };
    const body = `${JSON.stringify(round, null, 2)}\n`;
    const read = () => body;
    expect(() =>
      archive("log.json", "rounds", read as never, (() => {}) as never, (() => ["038-abc1234.json"]) as never),
    ).toThrow(/byte-identical to 038-abc1234\.json/);
  });

  it("does not hand out a number a round committed elsewhere already has", () => {
    // WHAT ACTUALLY WENT WRONG on 2026-08-16, and it is not an overwrite. Round
    // 064 was archived on `main` and its findings committed to a branch;
    // `git checkout main` then removed the file from the working tree, because
    // it is tracked only on that branch. The next round was archived from
    // `main`, where the directory ends at 063, and was numbered 064 as well.
    //
    // Nothing was destroyed — `nextRoundNumber` is max+1, so it never lands on a
    // file it can SEE. The damage is two different rounds both called 064 in two
    // git contexts: a merge collision, and a pooled report that reads one of
    // them twice or not at all. Every downstream number here is a pool over this
    // directory.
    //
    // The fix is what the lister is given, so that is what this pins.
    const round = { build: "abc1234 · 2026-08-16 05:00Z", trace: { entries: [] } };
    // Distinct bodies, or the byte-identical twin guard fires first and this
    // passes without ever reaching the numbering under test.
    const read = ((p: string) =>
      p === "log.json"
        ? `${JSON.stringify(round, null, 2)}\n`
        : `${JSON.stringify({ ...round, trace: { entries: [{ message: p }] } }, null, 2)}\n`) as never;
    const written: string[] = [];
    const write = ((p: string) => written.push(p)) as never;

    // The working tree alone — the bug. It cannot see 064 and reissues it.
    archive("log.json", "rounds", read as never, write, (() => ["063-9dc3bc8.json"]) as never);
    expect(written[0], "the working tree alone reissued a number git already holds").toBe("rounds/064-abc1234.json");

    // The union of disk and git — the fix. 064 is committed on a branch, so the
    // next round is 065 even though the file is nowhere on disk.
    written.length = 0;
    archive("log.json", "rounds", read as never, write, (() => ["063-9dc3bc8.json", "064-bcd5773.json"]) as never);
    expect(written[0], "a round committed on another branch was numbered over").toBe("rounds/065-abc1234.json");
  });

  it("files a log that differs from everything already kept", () => {
    const round = { build: "abc1234 · 2026-08-15 05:00Z", trace: { entries: [] } };
    const written: string[] = [];
    const name = archive(
      "log.json",
      "rounds",
      ((p: string) => (String(p) === "log.json" ? JSON.stringify(round) : "{}")) as never,
      ((p: string) => written.push(String(p))) as never,
      (() => ["038-abc1234.json"]) as never,
    );
    expect(name).toBe("039-abc1234.json");
    expect(written).toHaveLength(1);
  });

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
