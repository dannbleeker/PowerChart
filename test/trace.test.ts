import { afterEach, describe, expect, it } from "vitest";
import { setTracing, trace, traceLog, tracing } from "../src/core/trace";

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
