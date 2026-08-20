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
  namePresent,
  setSlideSizeScript,
  readSlideSize,
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
  browserDiedMidRound,
  onlyDirtyDeck,
  DEAD_BROWSER_POLLS,
  slideResolveScript,
  readSlideResolve,
  slideSizeScript,
  outcomeReceipt,
  RECOVERABLE_STOPS,
  selectDeck,
  sideloadAddIn,
  sweepDeck,
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

describe("the Automation tab holds everything the driver needs", () => {
  it("does not treat an unread toggle as permission to run", () => {
    // `Verbose trace`, `Picture every slide` and the run button all live on the
    // Automation tab, and a pane does not open there — a freshly sideloaded or
    // reopened one comes up on `Chart`, where none of the three is in the DOM.
    //
    // 2026-08-20: the driver put the add-in back by itself for the first time,
    // reached `ready` with `verbose trace ?` in the same block, and then stopped
    // because the run button was not there either. Both readings had one cause.
    //
    // `null` means "could not read it". A round started on that produces a trace
    // too thin to mine, and the point of selecting the tab first is that the
    // answer becomes a reading.
    const unread = readiness({ ...READY, verbose: null });
    expect(unread.ok, "an unread Verbose trace is not evidence that it is on").toBe(true);
    // The refusal that DOES exist stays: an explicit false still stops the round.
    const off = readiness({ ...READY, verbose: false });
    expect(off.ok).toBe(false);
    expect(off.stop.join(" ")).toMatch(/Verbose trace is off/);
  });
});

describe("a sign-in TAB is not a sign-in PROMPT", () => {
  it("refuses when the host is silent — the real prompt still stops the round", () => {
    // Unchanged behaviour, and it has to be: a genuine credential prompt makes
    // everything after it unmeasurable, and only a person can clear it.
    const r = readiness({ ...READY, loggedOut: true, authPopup: true, ping: null });
    expect(r.ok).toBe(false);
    expect(r.stop.join(" ")).toMatch(/needs a password/);
  });

  it("carries on when the host is ANSWERING, because that proves the session is signed in", () => {
    // 2026-08-20: Chrome's "Restore pages?" brought back an `oauth20_authorize`
    // popup from a crashed session, so the tab list showed a login page beside a
    // deck that was working perfectly. The check printed
    // `host answered in 7ms · slide 1 resolved` and refused in the same breath,
    // telling the owner to enter a password they had already entered.
    //
    // A host that answers Office.js and resolves a slide cannot be
    // unauthenticated. The ping is better evidence than the tab list because it
    // asks the thing we care about rather than looking at the furniture.
    const r = readiness({ ...READY, loggedOut: true, authPopup: true, ping: { answered: true, ms: 7 }, slideOk: true });
    expect(r.ok, "a stale login tab blocked a working deck").toBe(true);
  });

  it("still refuses when the slide will not resolve, even if the ping answered", () => {
    // Half-answering is not answering. `slideOk === false` is the host taking a
    // call and failing the one that matters, which is exactly the state the
    // original guard existed to stop a round on.
    const r = readiness({
      ...READY,
      loggedOut: true,
      authPopup: true,
      ping: { answered: true, ms: 7 },
      slideOk: false,
    });
    expect(r.ok).toBe(false);
  });
});

describe("setting the deck's slide size, only when asked", () => {
  it("writes BOTH dimensions and reads them back rather than assuming", () => {
    // `PowerPoint.PageSetup.slideWidth`/`slideHeight` are writable at
    // PowerPointApi 1.10, which round 096's `environment` line shows this host
    // advertising — so the `wrong-size` refusal was asking a person to do
    // something the API can do.
    //
    // The read-back is the point. The driver's own refusal text has warned for
    // months that a size change "made while the document is loading is accepted
    // and does nothing", and an API write deserves the same suspicion.
    const four = setSlideSizeScript("4:3", 20_000);
    expect(four, "4:3 is 720x540 at 96dpi").toMatch(/slideWidth = 720/);
    expect(four).toMatch(/slideHeight = 540/);
    expect(four, "wrote without reading the result back").toMatch(/load\("slideWidth,slideHeight"\)/);
    expect(four).toMatch(/return "size:"/);

    const wide = setSlideSizeScript("16:9", 20_000);
    expect(wide).toMatch(/slideWidth = 960/);
    expect(wide).toMatch(/slideHeight = 540/);
  });

  it("carries a budget, so a host that will not answer cannot hang the leg", () => {
    expect(setSlideSizeScript("4:3", 12_345)).toMatch(/12345/);
  });

  it("round-trips through readSlideSize, or the driver cannot tell whether it took", () => {
    // The setter returns the same `size:WxH` shape the reader parses. If those
    // two ever drift, the driver would resize the deck and then report that it
    // could not read the result — and carry on with the old value.
    expect(readSlideSize("size:720x540")).toBe("4:3");
    expect(readSlideSize("size:960x540")).toBe("16:9");
    expect(readSlideSize("size-failed:budget"), "a failure must not read as a size").toBeNull();
  });
});

describe("telling a control that is absent from one that is merely disabled", () => {
  // Playwright hands out a `ref` only for something it could act on, so a
  // greyed-out button matches its line and carries none. `refFor` answers null
  // for both, and the driver read that single null as "the add-in is gone".
  const findReturning = (out: string) => ((cmd: string) => (cmd === "find" ? out : "")) as never;

  it("sees a disabled control that refFor cannot", () => {
    const sh = findReturning('        - button "Insert chart" [disabled]');
    expect(refFor(sh, "Insert chart", /button "Insert chart"/), "a disabled control has no ref").toBeNull();
    expect(namePresent(sh, "Insert chart", /button "Insert chart"/), "but it IS in the ribbon").toBe(true);
  });

  it("still says absent when the ribbon really does not carry it", () => {
    const sh = findReturning('        - button "Design Suggestions" [ref=f12e42]');
    expect(namePresent(sh, "Insert chart", /button "Insert chart"/)).toBe(false);
  });

  it("agrees with refFor when the control is usable", () => {
    const sh = findReturning('        - button "Insert chart" [ref=f12e99]');
    expect(refFor(sh, "Insert chart", /button "Insert chart"/)).toBe("f12e99");
    expect(namePresent(sh, "Insert chart", /button "Insert chart"/)).toBe(true);
  });
});

describe("talking to the browser at all", () => {
  it("bounds every CLI call, because the round's own deadline cannot", () => {
    // THE STALL THIS DRIVER COULD NOT SURVIVE. The poll loop checks a 30-minute
    // limit at the TOP of each pass, so it only ever fires if the call below it
    // returned. An unbounded spawn meant a wedged CLI — a browser that stopped
    // answering, a tab mid-crash — hung the driver with that deadline sitting
    // there unreachable and nothing printed since.
    const seen: Record<string, unknown>[] = [];
    const run = (_exe: string, _args: string[], opts: Record<string, unknown>) => {
      seen.push(opts);
      return { status: 0, stdout: "" };
    };
    cli(run, ".", "some-cli.js")("list");
    expect(seen).toHaveLength(1);
    expect(seen[0].timeout, "an unbounded call can hang the whole night").toBeGreaterThan(0);
    // Generous, or it kills the slow calls that are legitimate: an `eval`
    // carries a 20s page-side budget and `requests` can return tens of MB.
    expect(seen[0].timeout).toBeGreaterThanOrEqual(60_000);
  });

  it("fronts the deck a round was told to run against, and refuses when it is not open", () => {
    // `PW_DECK` used to reach only `recover`, in the branch that reopens a dead
    // browser — so a cycle setting it per leg chose which deck a RECOVERY would
    // hunt for and nothing else. The ordinary path never selected a tab, so the
    // nightly cycle's 4:3 leg measured whichever document leg two left open and
    // refused with `wrong-size` every single night.
    const calls: string[][] = [];
    const shWith = (tabs: string) => {
      const sh = ((...args: string[]) => {
        calls.push(args);
        return args[0] === "tab-list" ? tabs : "";
      }) as never as { (...a: string[]): string; state: unknown };
      return sh;
    };
    const open = shWith('0: https://x/ "Presentation64"\n1: https://x/ "Presentation66"');
    expect(selectDeck(open, "Presentation66")).toBe(true);
    expect(
      calls.some((c) => c[0] === "tab-select" && c[1] === "1"),
      "did not front the deck it was given",
    ).toBe(true);
    // Not open at all: say so rather than measure the wrong document.
    expect(selectDeck(shWith('0: https://x/ "Presentation64"'), "Presentation66")).toBe(false);
    // No deck asked for is the behaviour this has always had.
    expect(selectDeck(shWith(""), null)).toBe(true);
  });

  /**
   * Putting the add-in back after a browser death took the sideload with it.
   *
   * Ten steps of Office ribbon automation, which will break the day Microsoft
   * moves a control. What these hold is not the walk but the two things that
   * decide whether a breakage is survivable: it must never leave a dialog over
   * the document, and it must never claim success it did not verify.
   */
  describe("putting the add-in back", () => {
    /** A fake `sh` whose answers are keyed by the query it is given. */
    const shWith = (answers: Record<string, string>, missing: string[] = []) => {
      const calls: string[][] = [];
      const sh = ((...args: string[]) => {
        calls.push(args);
        const q = args[1] ?? "";
        // EXACT, not substring. `"Manage My Add-ins".includes("Add-ins")` is
        // true, so a loose match hands three different queries the same answer
        // and the walk "passes" having clicked the wrong control twice.
        if (missing.includes(q)) return "";
        return answers[q] ?? "";
      }) as never as { (...a: string[]): string; calls: string[][] };
      sh.calls = calls;
      return sh;
    };
    // Every control the walk needs, each answering with a ref line `refFor` parses.
    const ALL: Record<string, string> = {
      "Slide List": 'listbox "Slide List" [ref=r1]',
      "Add-ins": 'button "Add-ins" [ref=r2]',
      "See all": 'menuitem "See all installed add-ins" [ref=r3]',
      "More Add-ins": 'menuitem "More Add-ins" [ref=r4]',
      "MY ADD-INS": 'tab "MY ADD-INS" [ref=r5]',
      "Manage My Add-ins": 'button "Manage My Add-ins" [ref=r6]',
      "Upload My Add-in": 'menuitem "Upload My Add-in" [ref=r7]',
      Browse: 'button "Browse..." [ref=r8]',
      Upload: 'button "Upload" [ref=r9]',
      "Insert chart": 'button "Insert chart" [ref=r10]',
    };
    const now = () => Promise.resolve();

    it("walks the whole flow and uploads the PROD manifest", async () => {
      const sh = shWith(ALL);
      expect(await sideloadAddIn(sh, now, "C:/x/manifest-prod.xml")).toBe(true);
      const uploaded = sh.calls.find((c) => c[0] === "upload");
      expect(uploaded?.[1], "sideloaded the dev manifest").toMatch(/manifest-prod\.xml$/);
      // The document must be focused FIRST — the ribbon ignores clicks until it
      // is, which cost three failed attempts to discover and reports itself as
      // `aria-expanded=true` over a menu that is plainly shut.
      const firstClick = sh.calls.findIndex((c) => c[0] === "eval");
      const listLookup = sh.calls.findIndex((c) => (c[1] ?? "").includes("Slide List"));
      expect(listLookup).toBeLessThan(firstClick);
    });

    it("never leaves a dialog over the document when a step is missing", async () => {
      // THE RISK THAT MATTERS MOST. A half-walked dialog is invisible to
      // `readiness` and fatal to every round after it — the deck reads fine and
      // nothing can be clicked.
      for (const step of ["More Add-ins", "Upload My Add-in", "Insert chart"]) {
        // Cancel has to be FINDABLE, or this asserts nothing: looking for a
        // dismiss control and clicking one are different acts, and the first
        // version of this test checked only that the lookup happened — it
        // passed with the clicks deleted.
        const sh = shWith({ ...ALL, Cancel: 'button "Cancel" [ref=rcancel]' }, [step]);
        expect(await sideloadAddIn(sh, now, "C:/x/m.xml"), `${step} should have failed`).toBe(false);
        const clickedCancel = sh.calls.some((c) => c[0] === "eval" && c[2] === "rcancel");
        expect(clickedCancel, `left a dialog open after failing at ${step}`).toBe(true);
      }
    });

    it("will not call it done when the host never accepted the manifest", async () => {
      // The Upload button is disabled until a file is accepted, so its enabled
      // state is the host's own receipt — better evidence than the upload call
      // not throwing.
      const sh = shWith({ ...ALL, Upload: 'button "Upload" [disabled] [ref=r9]' });
      expect(await sideloadAddIn(sh, now, "C:/x/m.xml")).toBe(false);
    });

    it("refuses to walk a document that is not up", async () => {
      // Ten ribbon steps against a loading tab would leave a dialog on it.
      const sh = shWith(ALL, ["Slide List"]);
      expect(await sideloadAddIn(sh, now, "C:/x/m.xml")).toBe(false);
      expect(
        sh.calls.some((c) => (c[1] ?? "").includes("Add-ins")),
        "started clicking anyway",
      ).toBe(false);
    });
  });

  it("believes the sweep's own answer rather than that it was called", () => {
    // "deck swept — the next round starts clean" printed whether the sweep had
    // cleaned the deck, failed outright, or left slides on it. The next round
    // then refuses with `deck-dirty` for a state the previous round's output
    // said could not exist. A tool that exits in silence reads as a pass.
    const shWith = (evalAnswer: string) =>
      ((...args: string[]) => {
        if (args[0] === "find") return 'tab "Chart" [ref=r1]';
        return evalAnswer;
      }) as never as (...a: string[]) => string;
    expect(sweepDeck(shWith("deck:1")), "a deck down to its one slide is clean").toBe(true);
    expect(sweepDeck(shWith("deck:7")), "seven slides left is not a swept deck").toBe(false);
    expect(sweepDeck(shWith("deck-failed")), "the sweep said it failed").toBe(false);
    expect(sweepDeck(shWith("")), "no answer is not a clean deck").toBe(false);
  });

  it("stops on the first attempt when the document has no add-in to open", () => {
    // SEVEN ATTEMPTS AND FIFTEEN MINUTES, on 2026-08-16. A 4:3 leg switched to a
    // deck PowerChart was not registered for; `recover` reopens the pane from
    // the ribbon's `Insert chart` control, that deck offers no such control, and
    // the loop rediscovered this six more times. A refusal recovery cannot
    // address must not be retried.
    // `commandPresent: false` is what "no add-in here" actually means — the
    // ribbon does not carry the control at all. Saying only `canOpenPane: false`
    // is a weaker statement (it may be present and greyed out) and this fixture
    // used to make it, which is how the two states got conflated in the driver.
    const r = readiness({ ...READY, stamp: null, canOpenPane: false, commandPresent: false });
    expect(r.ok).toBe(false);
    expect(r.codes).toContain("addin-missing");
    expect(r.codes, "still reported as a merely-closed pane").not.toContain("pane-closed");
    expect(shouldRetry("not-ready", 0, 6, r.codes), "retrying this cannot help").toBe(false);
    // A DISABLED CONTROL IS NOT AN ABSENT ONE, and this is the case that cost
    // a round on 2026-08-19. PowerPoint's document went `Disconnected` — the
    // network, the same fault that crashed round 089 — which greys the WHOLE
    // ribbon. The accessibility tree still said `button "Insert chart"
    // [disabled]`, the add-in was loaded and fine, and the driver called it
    // `addin-missing`: the one refusal recovery may not retry, needing the
    // owner, for a state a reload clears.
    const greyed = readiness({ ...READY, stamp: null, canOpenPane: false, commandPresent: true });
    expect(greyed.codes, "a greyed ribbon is not a missing add-in").not.toContain("addin-missing");
    expect(greyed.codes).toContain("host-disconnected");
    expect(shouldRetry("not-ready", 0, 6, greyed.codes), "a reload clears this — it must be retried").toBe(true);

    // And a pane that is simply shut on a deck that CAN open one stays
    // recoverable, or the loop stops doing the thing it is good at.
    const shut = readiness({ ...READY, stamp: null, canOpenPane: true });
    expect(shut.codes).toContain("pane-closed");
    expect(shouldRetry("not-ready", 0, 6, shut.codes)).toBe(true);
  });

  it("will not call a tab that is merely mid-reload an absent add-in", () => {
    // THE DANGEROUS HALF OF THE CHECK ABOVE. A reloading tab answers nothing to
    // every read, so the ribbon looks as bare as a document with no add-in — and
    // `addin-missing` is deliberately un-retryable, so getting this wrong ends a
    // night on a state a reload would have cleared in twenty seconds.
    //
    // A readable slide list is the proof the document is actually up. The deck
    // that motivated the check had one; a loading tab does not.
    const loading = readiness({ ...READY, stamp: null, canOpenPane: false, slides: null });
    expect(loading.codes, "condemned a tab that was only still loading").not.toContain("addin-missing");
    expect(loading.codes).toContain("pane-closed");
    expect(shouldRetry("not-ready", 0, 6, loading.codes), "a reload would have fixed this").toBe(true);
  });

  it("names the missing DECK rather than blaming its slide size", () => {
    // The message matters as much as the refusal. "the deck is 16:9 and this
    // round was asked for 4:3" sends the owner to Design ▸ Slide Size for a
    // deck that was simply never opened.
    const r = readiness({ ...READY, wantDeck: "Presentation66", deckFronted: false });
    expect(r.ok).toBe(false);
    expect(r.codes).toContain("deck-missing");
    expect(r.stop.join(" ")).toContain("Presentation66");
    // And it must not fire when no deck was asked for.
    expect(readiness({ ...READY }).ok).toBe(true);
  });

  it("treats a timed-out call as nothing measured, not as an answer", () => {
    // A timeout arrives as `error`, which is the same shape as any other spawn
    // failure — and the difference that matters is already drawn: nothing was
    // measured, so the sweep must not trust anything it read.
    const run = () => ({ error: Object.assign(new Error("timed out"), { code: "ETIMEDOUT" }), status: null });
    const sh = cli(run, ".", "some-cli.js");
    expect(sh("list")).toBe("");
    expect(sh.state.unreachable, "a wedged CLI must not read as a healthy empty answer").toBe(true);
  });

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

  it("notices a browser that dies UNDER a running round, which the quiet counter cannot", async () => {
    // 24 MINUTES ON 2026-08-16, and the cause is a good guard with a hole in it.
    //
    // `quietStreak` resets to zero whenever a CLI call FAILED — deliberately,
    // because one failed call means nothing was measured, and folding that into
    // "the pane is gone" once killed a healthy round that went on to pass 10 of
    // 12. But a dead browser makes EVERY call fail, permanently: the quiet
    // counter can never reach its threshold, the crash dialog cannot be read
    // either, and the loop polls a corpse until the thirty-minute limit. `pw
    // list` said `(no browsers)` outright while the driver sat there.
    expect(browserDiedMidRound(DEAD_BROWSER_POLLS, "  (no browsers)"), "a dead browser went unnoticed").toBe(true);

    // BOTH CONDITIONS ARE LOAD-BEARING, and each guards against re-making the
    // other's mistake.
    //
    // A streak without the list check would end a round on the absence of
    // evidence — exactly the contention bug above.
    expect(browserDiedMidRound(DEAD_BROWSER_POLLS, "### Browsers\n- ms:\n  - status: open")).toBe(false);
    expect(browserDiedMidRound(DEAD_BROWSER_POLLS, ""), "an unreadable list is not a dead browser").toBe(false);
    // And the list check without a streak would fire on a single blip, which is
    // what a second terminal touching the CLI looks like.
    expect(browserDiedMidRound(1, "  (no browsers)"), "one failed poll is contention, not a death").toBe(false);
    expect(browserDiedMidRound(0, "  (no browsers)")).toBe(false);
  });

  it("retries a browser that died mid-round, because reopening needs no password", () => {
    // The profile keeps the sign-in — a dead browser is not a lost sign-in — so
    // `recover` reopens it unattended. This is also the CHEAPEST of the
    // retryable reasons: it costs the minute already spent, where a wedge costs
    // the full thirty before anyone knows.
    expect(shouldRetry("browser-gone", 0, 3, undefined), "an unattended run stopped dead on a recoverable state").toBe(
      true,
    );
    // Still bounded by the caller's --retry N, like every other reason.
    expect(shouldRetry("browser-gone", 3, 3, undefined)).toBe(false);
  });

  it("sweeps a dirty deck rather than refusing, but only when it is the ONLY thing wrong", () => {
    // A dirty deck is not a fault — it is the last round's slides, and the
    // driver already sweeps them on every recovery. Refusing over it made a
    // person do by hand the one step the machine does better, and by hand is
    // how a deck reached ZERO slides on 2026-08-16: a fixed count of deletes
    // against a deck that held fewer, leaving `slide 1 REFUSED`, which is the
    // state the 2s crash starts from.
    expect(onlyDirtyDeck(["deck-dirty"]), "the one stop the driver can clear itself").toBe(true);

    // ONLY, and that word is the whole guard. Dirty AND stale is a round that
    // would measure the wrong build; healing the cheap half moves it closer to
    // running while still being wrong.
    expect(onlyDirtyDeck(["deck-dirty", "pane-stale"]), "healed its way into measuring the wrong build").toBe(false);
    expect(onlyDirtyDeck(["pane-stale"])).toBe(false);
    // No codes is no evidence — the same rule `shouldRetry` already applies.
    expect(onlyDirtyDeck([])).toBe(false);
    expect(onlyDirtyDeck(undefined)).toBe(false);
  });

  it("refuses a deck that is not the slide size this round was asked for", () => {
    // A CLICK THAT DID NOTHING. Setting Widescreen during the 2026-08-16 control
    // run silently did not take — it landed while the document was in its greyed
    // "Loading" state, the menu accepted it, and nothing changed. It was caught
    // only by reopening the menu and reading which box was ticked.
    //
    // A round that believes it is 4:3 and is not proves nothing, which is the
    // same harm as a round on a stale pane — already a hard stop. With a nightly
    // cycle running 16:9 twice and 4:3 once, an unverified size files a round
    // under the wrong profile, which is worse than not running it.
    const r = readiness({ ...READY, size: "16:9", expectSize: "4:3" });
    expect(r.ok).toBe(false);
    expect(r.codes).toContain("wrong-size");
    expect(r.stop.join(" "), "did not warn that a click can be accepted and do nothing").toMatch(/CHECK IT TOOK/);

    // Matching is silent, and asking for nothing is silent — an ordinary 16:9
    // round must read exactly as it always has.
    expect(readiness({ ...READY, size: "4:3", expectSize: "4:3" }).ok).toBe(true);
    expect(readiness({ ...READY, size: null, expectSize: null }).ok).toBe(true);
    // AND A HOST THAT WOULD NOT ANSWER IS NOT A MISMATCH. Null is no evidence,
    // and refusing on no evidence is the mistake `reachable` exists to prevent.
    expect(readiness({ ...READY, size: null, expectSize: "4:3" }).ok, "refused on an unanswered question").toBe(true);

    // NOT RECOVERABLE, and that is deliberate: `recover` could set the size in
    // two clicks, and doing so would change what the round measures rather than
    // restore it.
    expect(shouldRetry("not-ready", 0, 3, ["wrong-size"]), "recovery tried to change the deck's slide size").toBe(
      false,
    );
  });

  it("reads a slide size the way the add-in does, and calls silence silence", () => {
    expect(readSlideSize("size:960x540")).toBe("16:9");
    expect(readSlideSize("size:720x540")).toBe("4:3");
    // Anything else is its own profile, never folded into the nearest named one.
    expect(readSlideSize("size:1000x500")).toBe("1000x500");
    // A failure or a silence answers nothing — see the null case above.
    expect(readSlideSize("size-failed:GeneralException")).toBe(null);
    expect(readSlideSize("")).toBe(null);
    // Points via pageSetup, which is what the add-in itself reads.
    expect(slideSizeScript(15000)).toContain("slideWidth");
    expect(slideSizeScript(15000), "no budget means a wedged host hangs the check").toContain("15000");
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

  it("refuses to archive the PREVIOUS round's log when a wedge left the download doing nothing", () => {
    // THE FAILURE THE TWIN CHECK CANNOT SEE. `Download run log` is DISABLED
    // while a round is running, so clicking it after one that wedged does
    // nothing at all — and the previous round's file is still sitting at the
    // same path, waiting to be filed under a new number. Round 039 was archived
    // byte-identical to 038 that way: a whole round of evidence that never took
    // place, caught by a checksum and nothing else.
    //
    // The twin check only catches it when the two logs are IDENTICAL. A stale
    // log that merely DIFFERS — an older round, an earlier build — sails past
    // it. The pane's build stamp is what closes that, and it costs nothing: the
    // driver already read it to decide the round was worth running.
    const round = { build: "aaaaaaa · 2026-08-16 05:00Z", trace: { entries: [] } };
    const read = () => `${JSON.stringify(round, null, 2)}
`;
    const written: string[] = [];
    const write = ((p: string) => written.push(p)) as never;
    expect(
      () => archive("log.json", "rounds", read as never, write, (() => []) as never, "bbbbbbb"),
      "filed a log from a different build as this round",
    ).toThrow(/is the PREVIOUS round's file/);
    expect(written, "it refused and wrote anyway").toEqual([]);

    // And the matching case still files, or the guard would block every round.
    archive("log.json", "rounds", read as never, write, (() => []) as never, "aaaaaaa");
    expect(written[0]).toBe("rounds/001-aaaaaaa.json");
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

/**
 * The driver exits 0 or 1 and nothing else, so whatever runs the next round can
 * either be told what happened or go parsing sentences. `rounds-gate.mjs`
 * already refused to parse prose once; this is the same refusal with a file
 * behind it.
 */
describe("the account the driver leaves of how a round ended", () => {
  it("says which file the round produced, not merely that one was produced", () => {
    // "The newest file in rounds/" is the obvious substitute and it is the
    // assumption that produced a wrong overwrite diagnosis in this repo once.
    const r = outcomeReceipt({ reason: "finished", codes: [], roundFile: "082-0f7eadc.json", build: "0f7eadc" });
    expect(r.roundFile).toBe("082-0f7eadc.json");
    expect(r.build).toBe("0f7eadc");
    expect(r.reason).toBe("finished");
  });

  it("decides recoverability with the same set the driver retries on", () => {
    // Two implementations of "is this worth another attempt" is one too many.
    const one = [...RECOVERABLE_STOPS][0];
    expect(outcomeReceipt({ reason: "not-ready", codes: [one] }).recoverable).toBe(true);
    expect(outcomeReceipt({ reason: "not-ready", codes: [one, "wrong-size"] }).recoverable).toBe(false);
    // `wrong-size` must never read as recoverable: recovery COULD set a slide
    // size, and doing so would change what the round measures rather than
    // restore it.
    expect(RECOVERABLE_STOPS.has("wrong-size")).toBe(false);
  });

  it("hands back an array of codes even when there were none", () => {
    // A reader doing `codes.includes(...)` on a finished round should get
    // `false`, not a crash.
    expect(outcomeReceipt({ reason: "finished" }).codes).toEqual([]);
    expect(outcomeReceipt({ reason: "finished" }).recoverable).toBe(false);
    expect(outcomeReceipt({ reason: "crashed", codes: undefined }).codes).toEqual([]);
  });

  it("records the slide size in the shape the driver actually produces", () => {
    // THE PROFILE STRING, not an object. This test used to hand in
    // `{ width, height }` and assert it came back — which tested the
    // assumption, not the driver. `readSlideSize` returns "16:9", "4:3", or
    // "960x540", so the receipt recorded `"size": {}` on every cycle leg and
    // the field naming the arm a round belonged to said nothing at all.
    for (const [raw, profile] of [
      ["960x540", "16:9"],
      ["720x540", "4:3"],
      // Neither named ratio — a custom deck must be visibly its own profile
      // rather than folded into whichever one it is nearest.
      ["1000x500", "1000x500"],
    ]) {
      expect(outcomeReceipt({ reason: "finished", size: readSlideSize(`size:${raw}`) }).size).toBe(profile);
    }
    expect(outcomeReceipt({ reason: "finished" }).size).toBeNull();
  });

  it("numbers a round against every round ever filed, not just the working tree", () => {
    // TWO ROUNDS WERE BOTH FILED AS 064 on 2026-08-16. `everyRoundEverFiled`
    // was written for it and wired to the `--archive` subcommand — the path a
    // person uses by hand. `collectRound`, which archives EVERY round, went on
    // passing `readdirSync`: the working tree alone, blind to a round committed
    // on a branch that is not checked out, which is the exact collision.
    //
    // Guarded at the DEFAULT rather than at a call site, so the next caller
    // cannot repeat the omission. Read off the function itself because the
    // behaviour is a default value, and a test that passed its own lister —
    // as every other test here does — could never see it.
    const src = String(driver.archive);
    expect(src, "the automatic path can reuse a number again").toMatch(/list\s*=\s*everyRoundEverFiled/);
    expect(src, "readdirSync as the default is what filed 064 twice").not.toMatch(/list\s*=\s*readdirSync/);
  });

  it("files a round even when the archive names one that git has and the disk does not", () => {
    // THE TWIN CHECK OPENS EVERY NAME THE LISTER RETURNS, and since
    // `everyRoundEverFiled` became the default that list deliberately includes
    // rounds committed on a branch nobody has checked out. Unguarded it threw
    // ENOENT on exactly the HEALTHY path — `.find` reaches a git-only name only
    // when nothing matched, which is the case of a genuinely new round — and
    // `collectRound` swallowed it, so the driver exited 0 claiming a finished
    // round with nothing written and the next leg overwrote the evidence.
    const onDisk = '{\n  "build": "aaa1111 · x"\n}\n';
    const read = ((p: string) => {
      if (p.endsWith("log.json")) return JSON.stringify({ build: "ccc3333 · a fresh round" });
      if (p.endsWith("063-aaa1111.json")) return onDisk;
      // The git-only name: on nobody's disk, so opening it throws exactly as
      // `readFileSync` does.
      throw Object.assign(new Error("ENOENT: no such file or directory"), { code: "ENOENT" });
    }) as never;
    const written: string[] = [];
    const write = ((p: string) => written.push(p)) as never;
    const list = (() => ["063-aaa1111.json", "064-bbb2222.json"]) as never;
    const name = archive("log.json", "rounds", read, write, list);
    // Numbered past the round git knows about, which is the whole reason the
    // union lister is the default.
    expect(name).toBe("065-ccc3333.json");
    expect(written).toHaveLength(1);
  });

  it("still refuses a byte-identical twin that IS on disk", () => {
    // The guard the fix must not cost: a log the pane never rewrote means the
    // round did not finish, and filing it would double the weight of whatever
    // the real round said.
    const body = '{\n  "build": "aaa1111 · x"\n}\n';
    const read = ((p: string) => {
      if (p.endsWith("log.json")) return '{"build":"aaa1111 · x"}';
      if (p.endsWith("063-aaa1111.json")) return body;
      throw Object.assign(new Error("ENOENT"), { code: "ENOENT" });
    }) as never;
    expect(() =>
      archive(
        "log.json",
        "rounds",
        read,
        (() => {}) as never,
        (() => ["063-aaa1111.json", "064-bbb2222.json"]) as never,
      ),
    ).toThrow(/byte-identical to 063-aaa1111\.json/);
  });

  it("names a round that threw where nothing expected it to", () => {
    // An unexpected exception used to kill the process outright: no receipt, no
    // retry, and a night that ended on its first surprise.
    const r = outcomeReceipt({ reason: "threw", codes: [], threw: "Cannot read properties of undefined" });
    expect(r.reason).toBe("threw");
    expect(r.threw).toMatch(/Cannot read properties/);
    // RECOVERY DOES ADDRESS IT, and the receipt has to say so. `recoverable`
    // answers one question — did the driver retry this — and the cycle reads it
    // only to choose its wording; it stops either way. So the honest answer is
    // the one `shouldRetry` gives.
    //
    // This asserted `false` while `recoverable` looked at the CODES alone, and
    // most stops carry none: a crash, a wedge, a dead browser and a throw are
    // all retried on their reason. A night that recovered from a crash six times
    // ended by printing "crashed is not something recovery addresses — it needs
    // a person", directly under six lines saying "clearing the crash and
    // starting again".
    expect(r.recoverable).toBe(true);
  });

  it("agrees with shouldRetry about what recovery addresses", () => {
    // One question, one implementation. Two of these carry no codes at all, so a
    // receipt reading only `codes` gets every one of them wrong.
    for (const reason of ["crashed", "silent", "timeout", "browser-gone", "threw"])
      expect(outcomeReceipt({ reason, codes: [] }).recoverable, `${reason} is retried by the driver`).toBe(true);
    // And a stop recovery genuinely cannot touch still reads false.
    expect(outcomeReceipt({ reason: "not-ready", codes: ["wrong-size"] }).recoverable).toBe(false);
    expect(outcomeReceipt({ reason: "finished", codes: [] }).recoverable).toBe(false);
  });

  it("leaves `threw` off a round that did not throw", () => {
    expect(outcomeReceipt({ reason: "finished" })).not.toHaveProperty("threw");
  });

  it("retries a round that threw, and stops once the retries are spent", () => {
    // Bounded by the caller's own --retry, so a deterministic bug fails that
    // many times and stops rather than spinning all night.
    expect(shouldRetry("threw", 0, 3, [])).toBe(true);
    expect(shouldRetry("threw", 3, 3, [])).toBe(false);
  });
});
