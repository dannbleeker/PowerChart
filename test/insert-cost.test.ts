import { describe, it, expect } from "vitest";
import { batchMs, estimateInsertMs, isSlowInsert, describeMs, SLOW_INSERT_MS } from "../src/core/insert-cost";

/**
 * The insert estimate, against the archive it was measured from.
 *
 * The numbers here are the four measured medians in `insert-cost.ts` — 2,917
 * timed batches over 169 rounds — so these tests are the curve's own data
 * asserted back at it. If someone replaces the interpolation with a fitted
 * formula, these say whether the fit still passes through the measurements.
 */
describe("estimating what an insert will cost", () => {
  it("returns the measured medians at the measured points", () => {
    expect(batchMs(0)).toBe(3886);
    expect(batchMs(10)).toBe(5490);
    expect(batchMs(35)).toBe(13995);
    expect(batchMs(75)).toBe(18074);
  });

  it("interpolates between them, rather than stepping", () => {
    // Halfway from 10 to 35 shapes present is halfway from 5490 to 13995.
    expect(batchMs(22.5)).toBeCloseTo((5490 + 13995) / 2, 0);
    // Monotonic: a busier slide is never estimated cheaper.
    let last = -1;
    for (let n = 0; n <= 120; n += 3) {
      const v = batchMs(n);
      expect(v, `batchMs(${n}) went down`).toBeGreaterThanOrEqual(last);
      last = v;
    }
  });

  it("is FLAT past the last reading instead of extrapolating", () => {
    // Past 100 shapes this project has no measurements. A straight line through
    // the last two points would keep climbing on no evidence, and the estimate
    // would grow most confidently exactly where it knows least.
    expect(batchMs(200)).toBe(batchMs(75));
    expect(batchMs(5000)).toBe(batchMs(75));
  });

  it("costs a chart by batches, not by shapes", () => {
    // The renderer draws ten at a time, so 11 shapes is two batches and costs
    // the same as 20. Pretending otherwise would give a false precision.
    expect(estimateInsertMs(10, 0)).toBe(3886);
    expect(estimateInsertMs(11, 0)).toBe(3886 * 2);
    expect(estimateInsertMs(20, 0)).toBe(3886 * 2);
    expect(estimateInsertMs(0, 50)).toBe(0);
  });

  it("stays quiet on an empty slide and speaks up on a loaded one", () => {
    // THE WHOLE POINT OF THE THRESHOLD. A 24-shape chart onto a fresh slide is
    // the everyday case and must not nag; the same chart onto a slide already
    // holding 35 shapes is the 24-second case the user cannot currently see.
    expect(isSlowInsert(24, 0), "warned about a fast insert").toBe(false);
    expect(isSlowInsert(24, 35), "stayed quiet about a slow one").toBe(true);
  });

  it("understates rather than overstates", () => {
    // It uses the STARTING occupancy for every batch, but the slide fills as we
    // draw. A warning that fires early on a fast insert is noise; this project
    // would rather be quiet than cry wolf, so the estimate must not exceed the
    // cost of the same batches priced at the occupancy they will actually meet.
    const shapes = 30;
    const start = 20;
    const optimistic = estimateInsertMs(shapes, start);
    const asItFills = batchMs(start) + batchMs(start + 10) + batchMs(start + 20);
    expect(optimistic).toBeLessThanOrEqual(asItFills);
  });

  it("says a duration in words, roundly", () => {
    // A sentence, not a readout: "about 20 seconds" survives being wrong by two
    // in a way "19.4s" does not.
    expect(describeMs(19_400)).toBe("about 20 seconds");
    expect(describeMs(SLOW_INSERT_MS)).toBe("about 15 seconds");
    expect(describeMs(58_000)).toBe("about a minute");
    expect(describeMs(150_000)).toBe("about 2.5 minutes");
  });
});
