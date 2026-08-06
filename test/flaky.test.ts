import { describe, expect, it } from "vitest";
// @ts-expect-error — a plain .mjs tool with no types.
import { outcomes, disagreements, reportBody } from "../scripts/flaky.mjs";

/**
 * The flake hunt's comparison.
 *
 * A flaky test that passes on the run you are looking at is indistinguishable
 * from a good one, which is why this project has only ever found flakes by
 * accident — and misdiagnosed the one it found as a stale build artifact first.
 * The running-it-three-times half lives in a workflow; this is the half that
 * decides what the runs mean.
 */
const run = (...pairs: [string, string][]) => ({
  testResults: [
    {
      name: "/repo/test/a.test.ts",
      assertionResults: pairs.map(([title, status]) => ({ ancestorTitles: ["group"], title, status })),
    },
  ],
});

describe("comparing repeated runs", () => {
  it("names a test that did not do the same thing twice", () => {
    const found = disagreements([outcomes(run(["flaky", "passed"])), outcomes(run(["flaky", "failed"]))]);
    expect(found).toHaveLength(1);
    expect(found[0].name).toContain("flaky");
    expect(found[0].seen).toEqual(["passed", "failed"]);
  });

  it("says nothing about a suite that is red the same way every time", () => {
    // The distinction the whole tool rests on. A consistently red suite is a
    // broken build, and filing it as a flake sends the next person after a
    // timing bug that does not exist.
    const red = [outcomes(run(["broken", "failed"])), outcomes(run(["broken", "failed"]))];
    expect(disagreements(red)).toHaveLength(0);
  });

  it("treats a test that vanished from a run as a disagreement", () => {
    // Worth more than a status flip: a missing test means the file did not
    // load, which usually takes its whole suite with it — and a summary count
    // is exactly what hides that.
    const found = disagreements([outcomes(run(["gone", "passed"])), outcomes(run())]);
    expect(found[0].seen).toContain("did-not-run");
  });

  it("does not merge two tests that share a title in different files", () => {
    // Same `it()` text in two files is legitimate. Merging them would invent a
    // disagreement that never happened, and an invented flake costs the same
    // hour as a real one.
    const twoFiles = {
      testResults: [
        { name: "/repo/test/a.test.ts", assertionResults: [{ ancestorTitles: [], title: "same", status: "passed" }] },
        { name: "/repo/test/b.test.ts", assertionResults: [{ ancestorTitles: [], title: "same", status: "failed" }] },
      ],
    };
    expect(outcomes(twoFiles).size).toBe(2);
    expect(disagreements([outcomes(twoFiles), outcomes(twoFiles)])).toHaveLength(0);
  });

  it("says so plainly when every run agreed", () => {
    expect(reportBody([], 3)).toMatch(/All 3 runs agreed/);
  });
});
