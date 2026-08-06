import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types.
import { countOf, judgeCount } from "../scripts/test-count.mjs";

/**
 * The guard against a suite that silently shrinks.
 *
 * `CLAUDE.md` records the incident: a reorg deleted 43 tests and the suite still
 * went green, because nothing compared before against after. What is asserted
 * here is the decision, not the number — the number lives in
 * `test/fixtures/test-count.json` and rises on its own.
 */
describe("the test-count high-water mark", () => {
  it("fails when the suite has shrunk, and says by how much", () => {
    const v = judgeCount(2035, 2078);
    expect(v.ok).toBe(false);
    expect(v.message).toContain("43 test(s) went missing");
    // And tells the reader what to do about a deliberate drop, in the same
    // breath — a gate that only says no gets switched off.
    expect(v.message).toContain("--update");
  });

  it("passes when the suite grew, and when it stood still", () => {
    // Growth must cost nothing. An exact-count gate taxes every PR that adds a
    // test with a second edit and a merge conflict, which is how a gate earns
    // the resentment that gets it deleted.
    expect(judgeCount(2079, 2078).ok).toBe(true);
    expect(judgeCount(2078, 2078).ok).toBe(true);
  });

  it("refuses a report that is not a vitest run", () => {
    // The failure mode that would make this gate silently useless: point it at
    // the wrong file, read `undefined` as a count, and compare nothing forever.
    expect(() => countOf({ someOtherTool: true })).toThrow(/not a vitest JSON report/);
    expect(() => countOf(null)).toThrow();
    expect(countOf({ numTotalTests: 7 })).toBe(7);
  });
});
