import { afterEach, describe, expect, it } from "vitest";
import { setTracing, trace, traceLog, traceMark, tracing } from "../src/core/trace";

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
    trace("draw", "batch committed", payload);
    payload.shapes = 999;
    expect(traceLog().entries[0].data).toEqual({ shapes: 10 });
  });

  it("hands each reader its own copy of the data, not a shared one", () => {
    setTracing(true);
    trace("draw", "batch committed", { shapes: 10 });
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
