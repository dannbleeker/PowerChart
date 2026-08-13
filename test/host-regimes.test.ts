import { describe, expect, it } from "vitest";
// @ts-expect-error - .mjs script, deliberately outside tsconfig's typed surface
import { explainByRegime, flippedTogether, neverPutByRegime, verdictLine } from "../scripts/host-regimes.mjs";

/**
 * `scripts/host-regimes.mjs` — does the host's STATE account for a question that
 * changed its answer mid-round?
 *
 * The probe stamps every sample with the regime the host was in. Nothing read
 * those stamps mechanically before this: they were read by hand into
 * `UNSTABLE_ANSWERS`, which is how that table came to call a question "a coin"
 * that had answered the same way ten rounds running.
 *
 * The load-bearing case is `untested`. With three passes in three different
 * regimes, "every regime maps to one answer" is true by construction — a claim
 * that cannot fail is not evidence, and reporting it as `explained` would
 * manufacture a finding every time a question flipped. That is the same defect
 * as the rasterise control that always blamed the call because its arms ran in a
 * fixed order.
 */

type S = { answer: string; pass: number; atMs: number; regime: string };
const s = (answer: string, pass: number, regime: string): S => ({ answer, pass, atMs: pass * 1000, regime });

describe("does host state account for a mid-round flip", () => {
  it("says steady when the question never moved", () => {
    const r = explainByRegime([s("yes", 1, "healthy"), s("yes", 2, "collection-refused")]);
    expect(r.verdict).toBe("steady");
  });

  it("calls it EXPLAINED only when the mapping could have failed", () => {
    // healthy sampled twice and agreed both times, so "healthy means threw" was
    // a claim with a way to come out wrong. That is what makes it evidence.
    const r = explainByRegime([s("threw", 1, "healthy"), s("threw", 2, "healthy"), s("yes", 3, "collection-refused")]);
    expect(r.verdict).toBe("explained");
    expect(r.repeated).toEqual(["healthy"]);
  });

  it("calls it UNTESTED when every regime was sampled exactly once", () => {
    // THE GUARD THIS TOOL EXISTS FOR. Three passes, three regimes, three
    // answers: "every regime maps to one answer" is unfalsifiable here, so it
    // must NOT be reported as explained.
    const r = explainByRegime([
      s("threw", 1, "healthy"),
      s("yes", 2, "slide-trouble"),
      s("short-0", 3, "collection-refused"),
    ]);
    expect(r.verdict).toBe("untested");
  });

  it("calls it a COIN when one regime gave two different answers", () => {
    const r = explainByRegime([s("threw", 1, "healthy"), s("yes", 2, "healthy"), s("yes", 3, "collection-refused")]);
    expect(r.verdict).toBe("coin");
    expect(r.split[0].regime).toBe("healthy");
  });

  it("ignores passes that never put the question, rather than counting them as a face", () => {
    // A `no-scratch-slide` pass says nothing about the host. Counting it would
    // report a steady question as a flipper on a bad night — the same rule
    // `history()` applies across rounds.
    const r = explainByRegime([
      s("yes", 1, "healthy"),
      s("no-scratch-slide", 2, "slide-trouble"),
      s("yes", 3, "healthy"),
    ]);
    expect(r.verdict).toBe("steady");
    expect(r.real).toBe(2);
  });
});

describe("a trigger and its partner: one mechanism or two", () => {
  const trigger = [s("threw", 1, "healthy"), s("threw", 2, "healthy"), s("yes", 3, "collection-refused")];

  it("reads a shared boundary as lockstep", () => {
    const partner = [s("threw", 1, "healthy"), s("threw", 2, "healthy"), s("yes", 3, "collection-refused")];
    expect(flippedTogether(trigger, partner).verdict).toBe("lockstep");
  });

  it("reads different boundaries as independent", () => {
    // The partner moved between passes 1 and 2; the trigger between 2 and 3.
    const partner = [s("threw", 1, "healthy"), s("yes", 2, "healthy"), s("yes", 3, "collection-refused")];
    expect(flippedTogether(trigger, partner).verdict).toBe("independent");
  });

  it("refuses to compare when they share fewer than two answered passes", () => {
    // Nothing can be concluded from one shared pass, and saying "lockstep"
    // there would be the unfalsifiable shape again.
    const partner = [
      s("no-scratch-slide", 1, "slide-trouble"),
      s("no-scratch-slide", 2, "slide-trouble"),
      s("yes", 3, "collection-refused"),
    ];
    expect(flippedTogether(trigger, partner).verdict).toBe("not-comparable");
  });

  it("says neither moved when both held still across the shared passes", () => {
    const steady = [s("yes", 1, "healthy"), s("yes", 2, "healthy"), s("yes", 3, "collection-refused")];
    expect(flippedTogether(steady, steady).verdict).toBe("neither-moved");
  });
});

describe("what state the host was in for the questions it never put", () => {
  it("tallies the never-put attempts by regime, and names the questions", () => {
    const answers = [
      {
        id: "binding-names-shape-later",
        answer: "no-scratch-slide",
        samples: [s("no-scratch-slide", 1, "slide-trouble"), s("no-scratch-slide", 2, "slide-trouble")],
      },
      {
        id: "tag-on-group-survives",
        answer: "no-scratch-shape",
        samples: [s("no-scratch-shape", 1, "collection-refused")],
      },
      { id: "answered-one", answer: "yes", samples: [s("yes", 1, "healthy")] },
    ];
    const np = neverPutByRegime(answers);
    expect(np.ids).toEqual(["binding-names-shape-later", "tag-on-group-survives"]);
    expect(np.byRegime).toEqual([
      ["slide-trouble", 2],
      ["collection-refused", 1],
    ]);
  });
});

describe("rendering a verdict as prose", () => {
  /**
   * The first version of this built all four strings in one object literal and
   * picked one, so `split.map` ran on an `explained` result — where `split` does
   * not exist. It threw on the FIRST real round it was pointed at, and every
   * unit test above passed, because they call the decision functions and never
   * render a line. A tool that cannot print its own answer is not a tool.
   */
  const cases = [
    [[s("yes", 1, "healthy"), s("yes", 2, "healthy")], "steady"],
    [[s("threw", 1, "healthy"), s("threw", 2, "healthy"), s("yes", 3, "collection-refused")], "EXPLAINED"],
    [[s("threw", 1, "healthy"), s("yes", 2, "slide-trouble")], "UNTESTED"],
    [[s("threw", 1, "healthy"), s("yes", 2, "healthy")], "COIN"],
  ] as const;

  it("renders every verdict without touching a field that verdict does not carry", () => {
    for (const [samples, expected] of cases) {
      const r = explainByRegime(samples as unknown as S[]);
      expect(() => verdictLine(r), `${expected} threw while rendering`).not.toThrow();
      expect(verdictLine(r)).toContain(expected);
    }
  });
});
