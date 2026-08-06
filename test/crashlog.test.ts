// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  beginCrashLog,
  clearCrashLog,
  endCrashLog,
  flushCrashLog,
  markCrashLogSaved,
  recordCrashStep,
  recoverCrashLog,
  _resetCrashLogForTest,
} from "../src/taskpane/crashlog";
// @ts-expect-error — a plain .mjs tool with no types, deliberately independent
// of src/ so it cannot inherit a bug from the code it audits.
import { crashLogIn } from "../scripts/triage.mjs";

/**
 * The record that has to survive the run dying.
 *
 * Everything here is about one property: a run that never reported finishing
 * must still be readable afterwards. Two real-host rounds were lost because it
 * was not — one wedged at 1819 seconds and was killed by closing the tab, one
 * taken out at 108 seconds by PowerPoint's own crash dialog. Both produced no
 * log at all, because the only log this project had was built from a value
 * assigned after the last await.
 *
 * The tests deliberately do NOT reach into storage keys. What matters is the
 * behaviour at the boundary — write steps, do not finish, come back, get them —
 * and a test that asserted the key names would pass while the pane showed
 * nothing.
 */

const meta = { build: "abc1234", host: "PowerPoint · web · 16.0" };

/** Put a deliberately broken `window.localStorage` back, if a test swapped one in. */
let restoreStorage: (() => void) | undefined;

beforeEach(() => {
  window.localStorage.clear();
  _resetCrashLogForTest();
});
afterEach(() => {
  restoreStorage?.();
  restoreStorage = undefined;
  window.localStorage.clear();
  _resetCrashLogForTest();
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("the crash-surviving run log", () => {
  it("hands back a run that never reported finishing", async () => {
    beginCrashLog({ ...meta, label: "host self-test" });
    recordCrashStep("   0.1s  selftest  scenario started  name=insert on top of an earlier run");
    recordCrashStep("   9.4s  error  drawing the chart's shapes  error=host stopped answering");
    // The tab going away is the last synchronous moment an add-in gets, and
    // the pane spends it here. No endCrashLog — this is the run that died.
    flushCrashLog();

    // A fresh pane: nothing in memory, only what storage kept.
    _resetCrashLogForTest();
    const found = recoverCrashLog();
    expect(found, "the crashed run was not recoverable").not.toBeNull();
    expect(found?.finishedAt, "a run that never ended was marked finished").toBeUndefined();
    expect(found?.label).toBe("host self-test");
    expect(found?.build).toBe("abc1234");
    // Oldest first: a file is read from the top, unlike the on-screen list.
    expect(found?.steps[0]).toContain("scenario started");
    expect(found?.steps.at(-1)).toContain("host stopped answering");
  });

  it("lands on its own timer, without anyone asking it to", async () => {
    // The crash that matters most gives no warning at all: PowerPoint's own
    // "Sorry, we ran into a problem" takes the frame with no `pagehide`, no
    // unload, nothing to hang a flush on. The timer is the ONLY thing covering
    // that case, so a debounce that silently stopped firing would leave the
    // whole module passing its other tests and useless in the field.
    vi.useFakeTimers();
    beginCrashLog({ ...meta, label: "a run nobody gets to close" });
    recordCrashStep("   3.0s  selftest  the last thing it did");
    await vi.advanceTimersByTimeAsync(1000);
    vi.useRealTimers();

    _resetCrashLogForTest();
    const found = recoverCrashLog();
    expect(found?.steps, "the debounced write never landed").toContain("   3.0s  selftest  the last thing it did");
  });

  it("offers nothing back when the run ended normally AND its file was saved", () => {
    // The other half, and the half that makes it a test. A recovery banner
    // that appears after every healthy run is noise, and noise on this
    // particular banner is expensive: it is the one that means "something died".
    //
    // Both conditions, because finishing is not the same as the user having the
    // file — see the case below, which is the one a real round lost.
    beginCrashLog({ ...meta, label: "demo deck" });
    recordCrashStep("   0.1s  demo  inserting 38 slides");
    endCrashLog(true);

    _resetCrashLogForTest();
    expect(recoverCrashLog(), "a run that finished and saved was offered as a crash").toBeNull();
  });

  it("still offers a run that FINISHED but never saved its file", () => {
    // The round that was lost. It completed, printed "Saved as one file", and
    // marked itself finished BEFORE the download was attempted — so when
    // PowerPoint died moments later and the pane reopened, the record was
    // sitting in storage and the recovery would not look at it.
    //
    // A blocked download does the same thing without any crash: a task pane is
    // a nested cross-origin frame, and the browser can refuse to save from one
    // exactly as it already refuses the clipboard.
    beginCrashLog({ ...meta, label: "the whole round" });
    recordCrashStep("   9.9s  round  probe done, self-test done");
    endCrashLog();

    _resetCrashLogForTest();
    const found = recoverCrashLog();
    expect(found?.label, "a finished run whose file never landed was written off").toBe("the whole round");
    // And the pane can tell the two apart, because they mean different things
    // to whoever reads the banner: one says the host died, the other says the
    // host was fine and the file never arrived.
    expect(found?.finishedAt, "the record forgot that this run actually finished").toBeTruthy();
  });

  it("stops offering it once the user has actually pressed save", () => {
    beginCrashLog({ ...meta, label: "the whole round" });
    recordCrashStep("   9.9s  round  probe done, self-test done");
    endCrashLog();
    _resetCrashLogForTest();
    expect(recoverCrashLog(), "the setup did not leave anything on offer").toBeTruthy();

    markCrashLogSaved();
    _resetCrashLogForTest();
    expect(recoverCrashLog(), "the run was still offered after the user saved it").toBeNull();
  });

  it("does not let the next run destroy the crashed one's evidence", () => {
    // The reaction to a crash is to reload and try again. If starting a run
    // overwrote the only slot, that reaction would erase exactly what was
    // worth keeping — before anyone had a chance to look at it.
    //
    // Observed by running the second one to a clean finish AND a saved file. A
    // run in that state is not offered, so whatever comes back afterwards can
    // only be the crashed one — which says it survived, without this test
    // needing to know a storage key. The first version instead asserted the
    // crashed run was the one OFFERED, which is a different property and, as it
    // turned out, the wrong one: it passed happily while two consecutive
    // crashes handed back the older of them.
    beginCrashLog({ ...meta, label: "the run that died" });
    recordCrashStep("   5.0s  selftest  wedged here");
    flushCrashLog();

    _resetCrashLogForTest();
    beginCrashLog({ ...meta, label: "the run after it" });
    recordCrashStep("   0.1s  selftest  started again");
    endCrashLog(true);

    const found = recoverCrashLog();
    expect(found?.label, "the new run buried the crashed one").toBe("the run that died");
    expect(found?.steps.join("\n")).toContain("wedged here");
  });

  it("offers the MOST RECENT crash when two runs in a row died", () => {
    // The sequence that actually happens. A run dies; you reopen, see the
    // offer, and do the obvious thing — try again — rather than downloading
    // first. The second run dies too. Now there are two unfinished records, and
    // the one worth having is the SECOND: it is the one you were watching, on
    // the build you were testing.
    //
    // The first version handed back the older one, because promoting the
    // previous record to a safe slot and reading that slot first are both
    // right on their own and wrong together.
    beginCrashLog({ ...meta, label: "first crash" });
    recordCrashStep("   1.0s  selftest  died the first time");
    flushCrashLog();

    _resetCrashLogForTest();
    beginCrashLog({ ...meta, label: "second crash" });
    recordCrashStep("   1.0s  selftest  died again");
    flushCrashLog();

    _resetCrashLogForTest();
    const found = recoverCrashLog();
    expect(found?.label, "handed back the older crash, not the one just watched").toBe("second crash");
    // And the older one is still THERE — promoting it was not pointless, it
    // just must not win the tie.
    expect(recoverCrashLog()).not.toBeNull();
  });

  it("keeps the newest steps when the record outgrows its cap", () => {
    // A stalled run emits for as long as someone lets it — the 1819-second one
    // would have written tens of thousands of lines. What anyone reads is the
    // END: the last thing it did before it stopped. Dropping from the front is
    // the only truncation that preserves that.
    beginCrashLog({ ...meta, label: "a very long run" });
    for (let i = 0; i < 2500; i++) recordCrashStep(`step ${i}`);
    flushCrashLog();

    _resetCrashLogForTest();
    const found = recoverCrashLog();
    expect(found?.steps.length).toBeLessThanOrEqual(2000);
    expect(found?.steps.at(-1), "the newest step was dropped").toBe("step 2499");
    expect(found?.dropped, "dropped steps were not counted").toBeGreaterThan(0);
  });

  it("survives a storage that refuses to store", () => {
    // A task pane is a third-party frame. Storage can be absent, disabled by
    // policy, or full, and none of that is a reason for a RUN to fail —
    // logging must never be the thing that breaks the thing it is logging.
    const angry = {
      getItem: () => {
        throw new Error("SecurityError");
      },
      setItem: () => {
        throw new Error("SecurityError");
      },
      removeItem: () => {
        throw new Error("SecurityError");
      },
      clear: () => {},
      key: () => null,
      length: 0,
    } as unknown as Storage;
    // Both, because the module reaches through `window` and the rest of the
    // suite reaches through the global. The original descriptor is put back by
    // hand: `vi.unstubAllGlobals` restores what IT stubbed, and an own property
    // defined here is not that — leaving it in place would hand every later
    // test in this worker a storage that throws.
    const original = Object.getOwnPropertyDescriptor(window, "localStorage");
    vi.stubGlobal("localStorage", angry);
    Object.defineProperty(window, "localStorage", { value: angry, configurable: true });
    restoreStorage = () => {
      if (original) Object.defineProperty(window, "localStorage", original);
    };

    expect(() => {
      beginCrashLog({ ...meta, label: "no storage here" });
      recordCrashStep("a step nobody can keep");
      flushCrashLog();
      endCrashLog();
    }).not.toThrow();
    expect(recoverCrashLog()).toBeNull();
  });

  it("can still read a crashed run back when the store has no room left", () => {
    // Reading needs no quota. Writing does. Gating both behind a write probe
    // meant a full store made a record that was RIGHT THERE unreadable — and
    // the way a store gets full is our own 2000-step log, so this is the
    // scenario a long crashed run creates for the next one.
    beginCrashLog({ ...meta, label: "recorded while there was room" });
    recordCrashStep("   1.0s  selftest  the last thing it did");
    flushCrashLog();
    _resetCrashLogForTest();

    // Now the store is full: every write throws, reads are fine.
    const real = window.localStorage;
    const original = Object.getOwnPropertyDescriptor(window, "localStorage");
    const full = {
      getItem: (k: string) => real.getItem(k),
      setItem: () => {
        throw new DOMException("QuotaExceededError", "QuotaExceededError");
      },
      removeItem: (k: string) => real.removeItem(k),
      clear: () => real.clear(),
      key: (i: number) => real.key(i),
      get length() {
        return real.length;
      },
    } as unknown as Storage;
    Object.defineProperty(window, "localStorage", { value: full, configurable: true });
    restoreStorage = () => {
      if (original) Object.defineProperty(window, "localStorage", original);
    };

    const found = recoverCrashLog();
    expect(found?.label, "a full store hid a record it could have read").toBe("recorded while there was room");
    expect(found?.steps.join("\n")).toContain("the last thing it did");
  });

  it("keeps the newest half rather than nothing when a write is too big", () => {
    // The recovery that could never run. `flush` halves the record and retries
    // once on a rejected write — worth having, because the newest half of a
    // long run beats none of it — but every path first went through a write
    // probe, so a store that rejected writes bailed out before reaching it.
    const real = window.localStorage;
    const original = Object.getOwnPropertyDescriptor(window, "localStorage");
    let budget = 0;
    const tight = {
      getItem: (k: string) => real.getItem(k),
      setItem: (k: string, v: string) => {
        if (v.length > budget) throw new DOMException("QuotaExceededError", "QuotaExceededError");
        real.setItem(k, v);
      },
      removeItem: (k: string) => real.removeItem(k),
      clear: () => real.clear(),
      key: (i: number) => real.key(i),
      get length() {
        return real.length;
      },
    } as unknown as Storage;
    Object.defineProperty(window, "localStorage", { value: tight, configurable: true });
    restoreStorage = () => {
      if (original) Object.defineProperty(window, "localStorage", original);
    };

    budget = 1_000_000; // room while the run starts
    beginCrashLog({ ...meta, label: "outgrew the store" });
    for (let i = 0; i < 200; i++) recordCrashStep(`step ${i} ${"x".repeat(40)}`);
    // Now only half of it fits.
    budget = 6_000;
    flushCrashLog();

    _resetCrashLogForTest();
    const found = recoverCrashLog();
    expect(found, "gave up entirely instead of keeping the newest half").not.toBeNull();
    expect(found!.steps.length, "kept everything, so the write cannot have been refused").toBeLessThan(200);
    expect(found!.steps.at(-1), "kept the OLDEST half — the wrong half").toContain("step 199");
    expect(found!.dropped, "did not count what halving threw away").toBeGreaterThan(0);
  });

  it("forgets the crashed run once it has been saved", () => {
    beginCrashLog({ ...meta, label: "saved and done" });
    recordCrashStep("   1.0s  selftest  something");
    flushCrashLog();
    _resetCrashLogForTest();
    expect(recoverCrashLog()).not.toBeNull();

    clearCrashLog();
    expect(recoverCrashLog(), "the banner would come back on every open").toBeNull();
  });

  it("does not clear the run that is recording right now", () => {
    // `clearCrashLog` runs from the download button, which a user can press
    // while a NEW run is already going. Clearing the live slot then would
    // delete that run's evidence — this module committing the exact failure it
    // exists to prevent.
    beginCrashLog({ ...meta, label: "crashed earlier" });
    recordCrashStep("   1.0s  old  step");
    flushCrashLog();
    _resetCrashLogForTest();
    beginCrashLog({ ...meta, label: "running now" });
    recordCrashStep("   0.1s  new  step");
    flushCrashLog();

    clearCrashLog();
    // The live run is untouched: reading it back from a fresh context still
    // finds it, because it was never marked finished.
    _resetCrashLogForTest();
    const still = recoverCrashLog();
    expect(still?.label, "clearing the old record killed the live one").toBe("running now");
  });

  it("is a shape npm run triage can read on its own", () => {
    // The point of the file. A crashed run has no deck to be joined to — it
    // never got that far — so the tool has to answer from the log alone.
    beginCrashLog({ ...meta, label: "host self-test" });
    recordCrashStep("  90.0s  error  reading the selected chart  error=PowerPoint did not respond");
    flushCrashLog();
    _resetCrashLogForTest();
    const found = recoverCrashLog();

    const parsed = crashLogIn(JSON.parse(JSON.stringify(found)));
    expect(parsed, "triage would not recognise its own crash log").not.toBeNull();
    expect(parsed.steps).toHaveLength(1);
    // And it must NOT mistake the other two log shapes for this one.
    expect(crashLogIn({ runs: [], build: "x" })).toBeNull();
    expect(crashLogIn({ selftest: [], build: "x" })).toBeNull();
    expect(crashLogIn(null)).toBeNull();
  });
});
