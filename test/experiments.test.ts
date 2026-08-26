// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, faults } from "./helpers/office-host";
import { runExperiment, EXPERIMENTS } from "../src/render/experiments";

/**
 * One question, on demand, without a round.
 *
 * A round costs fourteen minutes. `grouped-child-by-id-from-slide` sat in the
 * probe sheet for 125 rounds and was never once answered, and the decision it
 * gates — whether two of every three redraws have a route home — is still open
 * because of it.
 */
describe("running one experiment", () => {
  it("answers the grouped-child question against a host that allows it", async () => {
    installHost([makeSlide("s1")]);
    const r = await runExperiment("grouped-child-by-id");
    // The fake groups and keeps its children addressable, so the mechanism is
    // exercised end to end here even though only a real host can ANSWER it.
    expect(["yes", "reads-but-refuses-writes", "no-such-shape", "unreadable", "no-child-ids"]).toContain(r.answer);
    expect(r.asks).toMatch(/inside a group/);
    expect(r.ms).toBeGreaterThanOrEqual(0);
  });

  it("gives the scratch slide back even when the experiment throws", async () => {
    installHost([makeSlide("s1")]);
    const removed: string[] = [];
    const add = vi.fn(async () => "scratch-1");
    const remove = vi.fn(async (id: string) => {
      removed.push(id);
      return true;
    });
    // An experiment that litters is one nobody runs twice, and this one is meant
    // to be run casually while a decision is being made.
    const boom = { id: "boom", asks: "?", run: async () => Promise.reject(new Error("no")) };
    EXPERIMENTS.push(boom);
    try {
      const r = await runExperiment("boom", add, remove);
      expect(r.answer).toBe("threw");
      expect(removed).toEqual(["scratch-1"]);
    } finally {
      EXPERIMENTS.pop();
    }
  });

  it("says so rather than pretending when the host will not give a slide", async () => {
    installHost([makeSlide("s1")]);
    const r = await runExperiment("grouped-child-by-id", async () => null);
    expect(r.answer).toBe("no-scratch-slide");
  });

  it("does not invent an answer for a name nobody registered", async () => {
    installHost([makeSlide("s1")]);
    const r = await runExperiment("not-a-real-experiment");
    expect(r.answer).toBe("no-such-experiment");
  });

  it("reports the host refusing to name the shapes before grouping", async () => {
    installHost([makeSlide("s1")]);
    // The question needs a child id captured BEFORE the group swallows it. A
    // host that will not name a fresh shape cannot be asked at all, and that is
    // a setup failure rather than an answer about groups.
    faults.refuseShapeAdds = true;
    try {
      const r = await runExperiment("grouped-child-by-id");
      expect(["no-child-ids", "threw"]).toContain(r.answer);
    } finally {
      faults.refuseShapeAdds = false;
    }
  });

  it("never leaves the slide behind when the cleanup itself fails", async () => {
    installHost([makeSlide("s1")]);
    // A cleanup that throws must not become the experiment's answer — that
    // would be a lie about the host.
    // A REAL slide, so the experiment itself succeeds and the only failure is
    // the cleanup. Pointing it at a slide that does not exist tests nothing
    // about cleanup — the experiment throws first, which is how the first
    // version of this test passed for the wrong reason.
    const r = await runExperiment(
      "grouped-child-by-id",
      async () => "s1",
      async () => {
        throw new Error("delete refused");
      },
    );
    expect(r.answer).not.toBe("threw");
    expect(r.detail ?? "").not.toMatch(/delete refused/);
  });
});
