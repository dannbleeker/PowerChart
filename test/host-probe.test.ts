// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, applyWebProfile } from "./helpers/office-host";
import { runHostProbes, PROBE_IDS } from "../src/render/host-probe";
// @ts-expect-error — a plain .mjs tool with no types. The baseline lives THERE
// rather than here, so the diff tool and this test cannot drift apart: two
// copies of the same table is how a claim quietly stops matching its check.
import { FAKE_BASELINE, diffAnswers, answersOf } from "../scripts/host-diff.mjs";

/**
 * The fake's own answer sheet, frozen.
 *
 * Two things this holds down, and the second is the point.
 *
 * 1. The sheet comes back COMPLETE. A probe that throws, times out or wedges
 *    contributes its answer and nothing else — the sheet from a misbehaving
 *    host is the one worth having, so surviving misbehaviour is the feature.
 *
 * 2. What the fake CLAIMS about itself is written down. Every assertion in this
 *    repo about Office.js rests on this fake, and nobody has ever checked it
 *    against a real PowerPoint. The baseline below is one half of that check;
 *    the other half arrives the first time someone runs the probe in a real
 *    host and `npm run host-diff` compares the two.
 *
 * So a change to this baseline is never "just update it". It means the fake now
 * claims something different about the host it stands for, and the question is
 * whether the real host agrees.
 */

afterEach(() => vi.unstubAllGlobals());

const sheetOf = async () => {
  const answers = await runHostProbes("fake", "test");
  return Object.fromEntries(answers.answers.map((a) => [a.id, a.answer]));
};

describe("the fake host's answer sheet", () => {
  it("answers every question, and says what it claims", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await sheetOf();
    expect(Object.keys(sheet).sort()).toEqual([...PROBE_IDS].sort());
    // Compared against the table the DIFF tool carries, so the two cannot
    // disagree about what the fake claims. The commentary on each answer lives
    // beside it in `scripts/host-diff.mjs`, together with what would be wrong
    // if a real host answered differently.
    expect(sheet).toEqual(FAKE_BASELINE);
  });

  it("agrees with itself, so a diff of a sheet against the baseline is empty", () => {
    // The diff's own sanity: identical sheets must produce no divergence, or
    // every real comparison is noise.
    const d = diffAnswers(FAKE_BASELINE, FAKE_BASELINE);
    expect(d.differ).toEqual([]);
    expect(d.onlyReal).toEqual([]);
    expect(d.onlyFake).toEqual([]);
    expect(d.agree.length).toBe(PROBE_IDS.length);
  });

  it("reports a disagreement, and a question one side was never asked", () => {
    // A question one side has no answer for is a HOLE in the comparison, not a
    // match. Counting it as agreement is how a diff comes to mean nothing.
    const real = { ...FAKE_BASELINE, "untrack-available": "yes", "extra-question": "hm" };
    delete real["getitemat-past-end"];
    const d = diffAnswers(real, FAKE_BASELINE);
    expect(d.differ.map((x: { id: string }) => x.id)).toEqual(["untrack-available"]);
    expect(d.differ[0].real).toBe("yes");
    expect(d.differ[0].fake).toBe("no");
    expect(d.differ[0].means, "a divergence with nothing said about what rests on it").toBeTruthy();
    expect(d.onlyReal).toEqual(["extra-question"]);
    expect(d.onlyFake).toEqual(["getitemat-past-end"]);
  });

  it("reads a real sheet's wrapper, not just a bare map", () => {
    const sheet = {
      kind: "powerchart-host-answers",
      source: "PowerPoint web",
      build: "abc",
      requirementSets: ["1.1"],
      answers: [{ id: "untrack-available", question: "?", answer: "yes", ms: 1 }],
    };
    expect(answersOf(sheet)).toEqual({ "untrack-available": "yes" });
    expect(answersOf({ nonsense: true })).toEqual({ nonsense: true });
    expect(answersOf(null)).toBeNull();
  });

  it("still comes back complete from a host at its worst", async () => {
    // The sheet's whole value is that a misbehaving host still produces one.
    // A probe run that gave up half way would be worthless exactly when it
    // was needed.
    installHost([makeSlide("s1")]);
    applyWebProfile();
    const sheet = await runHostProbes("fake-web-profile", "test");
    expect(sheet.answers).toHaveLength(PROBE_IDS.length);
    for (const a of sheet.answers) {
      expect(a.answer, `${a.id} gave no answer`).toBeTruthy();
      expect(a.question, `${a.id} lost its question`).toBeTruthy();
      expect(a.ms).toBeGreaterThanOrEqual(0);
    }
  });

  it("carries the requirement sets, without which no answer can be read", async () => {
    // The same verdict means different things on 1.4 and on 1.10. A sheet that
    // did not say which would be uninterpretable a week later.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    expect(sheet.kind).toBe("powerchart-host-answers");
    expect(sheet.requirementSets.length).toBeGreaterThan(0);
    expect(sheet.requirementSets).toContain("1.5");
  });

  it("leaves the deck exactly as it found it", async () => {
    // A diagnostic that litters is one people stop running. Probes add shapes,
    // groups and tags; all of it belongs to a scratch slide that goes back.
    const deck = [makeSlide("s1"), makeSlide("s2")];
    installHost(deck);
    const { slideCount } = await import("../src/render/powerpoint");
    const before = await slideCount();
    await runHostProbes("fake", "test");
    expect(await slideCount(), "the probe run left a slide behind").toBe(before);
  });
});
