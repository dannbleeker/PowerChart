// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  beginCrashLog,
  clearCrashLog,
  endCrashLog,
  flushCrashLog,
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

beforeEach(() => {
  window.localStorage.clear();
  _resetCrashLogForTest();
});
afterEach(() => {
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

  it("offers nothing back when the run ended normally", () => {
    // The other half, and the half that makes it a test. A recovery banner
    // that appears after every healthy run is noise, and noise on this
    // particular banner is expensive: it is the one that means "something died".
    beginCrashLog({ ...meta, label: "demo deck" });
    recordCrashStep("   0.1s  demo  inserting 38 slides");
    endCrashLog();

    _resetCrashLogForTest();
    expect(recoverCrashLog(), "a run that finished cleanly was offered as a crash").toBeNull();
  });

  it("does not let the next run destroy the crashed one's evidence", () => {
    // The reaction to a crash is to reload and try again. If starting a run
    // overwrote the only slot, that reaction would erase exactly what was
    // worth keeping — before anyone had a chance to look at it.
    beginCrashLog({ ...meta, label: "the run that died" });
    recordCrashStep("   5.0s  selftest  wedged here");
    flushCrashLog();

    _resetCrashLogForTest();
    beginCrashLog({ ...meta, label: "the run after it" });
    recordCrashStep("   0.1s  selftest  started again");
    flushCrashLog();

    const found = recoverCrashLog();
    expect(found?.label, "the new run buried the crashed one").toBe("the run that died");
    expect(found?.steps.join("\n")).toContain("wedged here");
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
    vi.stubGlobal("localStorage", angry);
    Object.defineProperty(window, "localStorage", { value: angry, configurable: true });

    expect(() => {
      beginCrashLog({ ...meta, label: "no storage here" });
      recordCrashStep("a step nobody can keep");
      flushCrashLog();
      endCrashLog();
    }).not.toThrow();
    expect(recoverCrashLog()).toBeNull();
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
