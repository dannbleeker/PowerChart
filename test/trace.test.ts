import { afterEach, describe, expect, it } from "vitest";
import {
  formatTraceLine,
  formatTraceValue,
  onTrace,
  setTracing,
  trace,
  traceLog,
  traceMark,
  tracing,
} from "../src/core/trace";

afterEach(() => setTracing(false));

describe("the activity trace", () => {
  it("records nothing at all while it is off", () => {
    setTracing(false);
    trace("demo", "something happened", { a: 1 });
    expect(tracing()).toBe(false);
    expect(traceLog().entries).toEqual([]);
  });

  it("records scope, message and data once switched on", () => {
    setTracing(true);
    trace("host", "gave up waiting", { what: "drawing shapes", afterMs: 45000 });
    const { entries } = traceLog();
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({ scope: "host", message: "gave up waiting" });
    expect(entries[0].data).toEqual({ what: "drawing shapes", afterMs: 45000 });
    // Relative to when tracing started, so a log reads as a timeline.
    expect(entries[0].ms).toBeGreaterThanOrEqual(0);
    expect(entries[0].ms).toBeLessThan(60_000);
  });

  it("omits the data key entirely when there is none", () => {
    setTracing(true);
    trace("pane", "action started");
    expect(traceLog().entries[0]).not.toHaveProperty("data");
  });

  it("starts clean each time it is switched on", () => {
    setTracing(true);
    trace("demo", "from the first run");
    setTracing(true);
    trace("demo", "from the second");
    expect(traceLog().entries.map((e) => e.message)).toEqual(["from the second"]);
  });

  it("drops the oldest entries rather than growing without bound", () => {
    // A stalled web run can emit steps for fifteen minutes. The last few
    // hundred are what anyone reads, and an unbounded array in a struggling
    // browser tab is the last thing this should add to.
    setTracing(true);
    for (let i = 0; i < 2100; i++) trace("draw", `step ${i}`);
    const { entries, dropped } = traceLog();
    expect(entries).toHaveLength(2000);
    expect(dropped).toBe(100);
    expect(entries[0].message).toBe("step 100");
    expect(entries.at(-1)?.message).toBe("step 2099");
  });

  it("hands back a copy, so a caller cannot corrupt the log it is reading", () => {
    setTracing(true);
    trace("demo", "one");
    const first = traceLog().entries;
    first.push({ ms: 0, scope: "forged", message: "not ours" });
    expect(traceLog().entries).toHaveLength(1);
  });
});

describe("the trace is a record, not a live view", () => {
  it("keeps what was true when the entry was written, not what the caller did next", () => {
    // A log is evidence. Holding the caller's object meant a call site that
    // reused or mutated it afterwards rewrote history — in a file someone may
    // already have downloaded and be reading as fact.
    setTracing(true);
    const payload: Record<string, unknown> = { shapes: 10 };
    trace("draw", "batch issued", payload);
    payload.shapes = 999;
    expect(traceLog().entries[0].data).toEqual({ shapes: 10 });
  });

  it("hands each reader its own copy of the data, not a shared one", () => {
    setTracing(true);
    trace("draw", "batch issued", { shapes: 10 });
    const first = traceLog().entries[0].data!;
    first.shapes = 999;
    expect(traceLog().entries[0].data).toEqual({ shapes: 10 });
  });
});

describe("reading back one operation's slice of the log", () => {
  it("returns only what happened after the mark", () => {
    // The buffer spans every operation since tracing was switched on, with
    // nothing but a "pane: action started" line between them. A run log that
    // carried the whole buffer carried other runs too — and reading one run's
    // per-item numbers against another run's trace is an expensive mistake.
    setTracing(true);
    trace("demo", "an earlier run");
    const mark = traceMark();
    trace("demo", "this run");
    trace("demo", "still this run");
    expect(traceLog(mark).entries.map((e) => e.message)).toEqual(["this run", "still this run"]);
    // …and the whole buffer is still available to anyone who wants it.
    expect(traceLog().entries).toHaveLength(3);
  });

  it("stays correct after entries have fallen off the front", () => {
    // A mark is an absolute position, so it has to survive the ring buffer
    // dropping older entries out from under it — otherwise a long run's slice
    // silently starts in the wrong place.
    setTracing(true);
    const mark = traceMark();
    for (let i = 0; i < 2100; i++) trace("draw", `step ${i}`);
    const { entries, dropped } = traceLog(mark);
    expect(dropped).toBe(100);
    expect(entries).toHaveLength(2000); // everything still held, none skipped twice
    expect(entries[0].message).toBe("step 100");
  });

  it("hands back everything when the mark predates the buffer entirely", () => {
    setTracing(true);
    trace("demo", "one");
    expect(traceLog(0).entries).toHaveLength(1);
  });
});

describe("the summary that makes a long trace readable", () => {
  it("tallies each distinct scope+message, commonest first", () => {
    // A real run's log is 160 KB of pretty-printed JSON holding a few hundred
    // entries — more than fits in one read, let alone on a screen. The two
    // questions ever asked of it are "what did this run do" and "what went
    // wrong", and both are answered by a tally rather than by the entries.
    setTracing(true);
    trace("demo", "item finished");
    trace("group", "tagging failed");
    trace("demo", "item finished");
    trace("demo", "item finished");
    expect(traceLog().summary.steps).toEqual([
      { scope: "demo", message: "item finished", n: 3 },
      { scope: "group", message: "tagging failed", n: 1 },
    ]);
  });

  it("keeps the same message under different scopes apart", () => {
    setTracing(true);
    trace("demo", "read the deck back");
    trace("repair", "read the deck back");
    expect(traceLog().summary.steps).toHaveLength(2);
  });

  it("tallies failure reasons whichever word the call site used for them", () => {
    // `error`, `why` and `reason` are all in live use. A histogram that knew
    // only `error` missed two thirds of one real run's problems.
    setTracing(true);
    trace("group", "tagging failed", { error: "InvalidParam passed to GetItem(id)" });
    trace("group", "tagging failed", { error: "InvalidParam passed to GetItem(id)" });
    trace("repair", "left alone", { reason: "chart is one object but carries no config tag" });
    trace("demo", "degraded", { why: "an item took 53s" });
    expect(traceLog().summary.problems).toEqual([
      { text: "InvalidParam passed to GetItem(id)", n: 2 },
      { text: "chart is one object but carries no config tag", n: 1 },
      { text: "an item took 53s", n: 1 },
    ]);
  });

  it("truncates a problem string so one error's debugInfo blob cannot swamp the tally", () => {
    setTracing(true);
    const long = `InvalidParam passed to GetItem(id) | code=5010 | debugInfo=${"x".repeat(400)}`;
    trace("group", "tagging failed", { error: long });
    const [only] = traceLog().summary.problems;
    expect(only.text).toHaveLength(120);
    expect(only.text.startsWith("InvalidParam passed to GetItem(id)")).toBe(true);
  });

  it("ignores non-string and empty reasons rather than tallying [object Object]", () => {
    setTracing(true);
    trace("repair", "applying delete", { reason: "" });
    trace("repair", "applying delete", { reason: { code: 5010 } as unknown as string });
    expect(traceLog().summary.problems).toEqual([]);
  });

  it("summarises only the slice asked for, never the whole buffer", () => {
    // The mark exists because a log that carried the whole buffer carried
    // other runs; a summary computed over the whole buffer would put those
    // runs' failures back into this one's headline count.
    setTracing(true);
    trace("group", "tagging failed", { error: "an earlier run's problem" });
    const mark = traceMark();
    trace("demo", "item finished");
    const { summary } = traceLog(mark);
    expect(summary.steps).toEqual([{ scope: "demo", message: "item finished", n: 1 }]);
    expect(summary.problems).toEqual([]);
  });
});

describe("watching the trace as it happens", () => {
  afterEach(() => onTrace(undefined));

  it("hands each entry to a watcher as it is recorded, not at the end", () => {
    // The pane's live step list is fed from here. It exists because the run
    // LOG does not survive the failures worth explaining: it becomes
    // downloadable only when a run ends, and real-host rounds have been lost
    // to a run that never ended and to a PowerPoint killed outright. Whatever
    // is already on screen survives both — so a watcher must see each step at
    // the moment it happens, not receive a batch afterwards.
    setTracing(true);
    const seen: string[] = [];
    onTrace((e) => seen.push(`${e.scope}:${e.message}`));
    trace("selftest", "scenario starting", { name: "alpha" });
    expect(seen, "the watcher was not called during the run").toEqual(["selftest:scenario starting"]);
    trace("selftest", "scenario passed", { name: "alpha" });
    expect(seen).toHaveLength(2);
  });

  it("keeps the record when the watcher throws", () => {
    // The window is a convenience; the log is the evidence. A broken renderer
    // must never cost an entry — least of all during the crash it is there to
    // photograph.
    setTracing(true);
    onTrace(() => {
      throw new Error("the pane blew up");
    });
    expect(() => trace("host", "gave up waiting", { what: "rasterising a slide" })).not.toThrow();
    expect(traceLog().entries).toHaveLength(1);
  });

  it("says nothing to a watcher while tracing is off", () => {
    setTracing(false);
    const seen: string[] = [];
    onTrace((e) => seen.push(e.message));
    trace("demo", "something happened");
    expect(seen).toEqual([]);
  });
});

describe("the one line a step list and a crash log share", () => {
  /**
   * The formatter lived inside `wireInsert`, a 1113-line DOM closure, where
   * nothing could reach it — and it dropped every value whose `typeof` was
   * "object". That is every array, and the arrays are the payloads that matter:
   * `degradation curves` carries the two timing series its whole experiment
   * exists to produce, and `tag pass over a page` carries `withoutTag`, the
   * list of charts that lost their config.
   *
   * Both go to the screen AND to the crash log through this one function. So an
   * experiment that killed the tab left behind its verdict and none of its
   * numbers — in the one artefact that survives a dead tab.
   */
  it("keeps an array payload instead of dropping it", () => {
    const line = formatTraceLine({
      ms: 12_340,
      scope: "selftest",
      message: "degradation curves",
      data: { oneContext: [2339, 3177, 3704], freshContext: [2852, 3397], suspect: "the slide" },
    });
    expect(line, "the timing series the experiment exists to produce was dropped").toContain(
      "oneContext=[2339,3177,3704]",
    );
    expect(line).toContain("freshContext=[2852,3397]");
    expect(line).toContain("suspect=the slide");
  });

  it("keeps the list of charts a repair pass could not tag", () => {
    const line = formatTraceLine({
      ms: 100,
      scope: "repair",
      message: "tag pass over a page",
      data: { slides: 4, withoutTag: ["17", "23"] },
    });
    expect(line, "the only part of this line worth reading was dropped").toContain('withoutTag=["17","23"]');
  });

  it("caps one value rather than letting it push the line off the pane", () => {
    const long = Array.from({ length: 200 }, (_, i) => i);
    const line = formatTraceLine({ ms: 0, scope: "host", message: "many", data: { ids: long, after: "kept" } });
    expect(line.length, "a breadcrumb became a heap dump").toBeLessThan(260);
    expect(line, "the cap ate the values after it").toContain("after=kept");
    expect(line).toMatch(/ids=\[0,1,2/);
    expect(line).toContain("…");
  });

  it("keeps 0, false and null, and leaves out only what has nothing to say", () => {
    // `0` is an answer — `settled=0` is the whole finding in a settle-pass line
    // — and so is `false`. Only undefined and the empty string mean "nothing to
    // report", and `key=` with nothing after it is noise on a line read at a
    // glance.
    expect(formatTraceValue(0)).toBe("0");
    expect(formatTraceValue(false)).toBe("false");
    expect(formatTraceValue(null)).toBe("null");
    expect(formatTraceValue(undefined)).toBeUndefined();
    expect(formatTraceValue("")).toBeUndefined();
    const line = formatTraceLine({ ms: 0, scope: "group", message: "settle pass:", data: { settled: 0, lost: 1 } });
    expect(line).toContain("settled=0 lost=1");
  });

  it("survives a payload that cannot be stringified", () => {
    // A circular payload is a caller's mistake. Losing the step it was
    // attached to would make that mistake cost the evidence too.
    const circular: Record<string, unknown> = { name: "loop" };
    circular.self = circular;
    const line = formatTraceLine({ ms: 0, scope: "host", message: "odd", data: { circular, keep: 7 } });
    expect(line).toContain("keep=7");
    expect(line).toContain("odd");
  });

  it("renders the elapsed stamp the way the crash logs are read", () => {
    // One decimal, right-aligned, so a column of steps lines up — the crash log
    // is read as a timeline and a ragged left edge is what made the previous
    // format hard to scan.
    expect(formatTraceLine({ ms: 33_249, scope: "selftest", message: "step" })).toBe("  33.2s  selftest  step");
    expect(formatTraceLine({ ms: 0, scope: "pane", message: "start" })).toBe("     0s  pane  start");
  });
});
