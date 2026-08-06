// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, applyWebProfile, faults } from "./helpers/office-host";
import { runHostProbes, PROBE_IDS, describeHostSheet, sheetNeedsAttention } from "../src/render/host-probe";
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

/** The probe's own "this was never put" vocabulary — never a host answer. */
const NOT_ASKED_WORDS = ["no-scratch-slide", "no-scratch-shape"];

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

  /**
   * The property that makes these answers worth reading: no probe holds a
   * proxy across a sync, so a host that refuses stale ones changes nothing.
   *
   * `strictTags` and `strictGroup` are the two faults that model exactly what
   * PowerPoint on the web does to a proxy older than its sync — and four
   * questions on the 2026-08-04 sheet were answering those faults rather than
   * their own questions. `tag-on-group-survives` said NO, which read as "no
   * chart in any deck is re-editable"; `group-reports-its-children` threw
   * PropertyNotLoaded; `tags-add-same-key-twice` and `addgroup-returns-usable`
   * came back undefined. Every one of them wrote or read through a handle it
   * had been given a sync earlier.
   *
   * Deliberately NOT the whole web profile: `refuseGroups` and `hollowReads`
   * change the answers legitimately, and a guard that tolerated those could not
   * tell them from a probe that had gone stale again.
   *
   * What it does NOT prove, said plainly: the fake refuses a proxy older than
   * TWO syncs, and three of those four probes were exactly one sync late, so
   * only `tags-add-same-key-twice` goes red here without the rewrite. The
   * others rest on the same rule and on the real host's own answer to
   * `shape-add-held-slide-proxy` — one sync was enough there. Tightening the
   * fake to one sync for SHAPES would make this guard catch all four, and
   * nobody has asked a host whether that is true; `shape-proxy-survives-one-
   * sync` only establishes it at two. That question is the next sheet's.
   */
  it("answers the same under a host that refuses every stale proxy", async () => {
    installHost([makeSlide("s1")]);
    faults.strictTags = true;
    faults.strictGroup = true;
    try {
      const sheet = await runHostProbes("fake-strict-proxies", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      expect(answers).toEqual(FAKE_BASELINE);
    } finally {
      faults.strictTags = false;
      faults.strictGroup = false;
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

  /**
   * The failure a real host actually produced, and the reason this run now
   * replaces a scratch slide it has lost.
   *
   * PowerPoint on the web resolved the scratch slide's id for the first
   * question and refused it for the other thirteen. Every one of those thirteen
   * was recorded as `"threw"` — a legitimate answer to several of these
   * questions, and one the diff tool compares against the fake's — so a sheet
   * that had asked ONE question came back looking like thirteen host
   * divergences. The questions were never put.
   */
  it("asks every question even when the host keeps losing the scratch slide", async () => {
    installHost([makeSlide("s1")]);
    // Each new slide answers to its id six times and then denies existing.
    //
    // Was two, when every probe shared the one slide handle `withProbeContext`
    // resolved; then four; now six. The number tracks one thing — the most
    // slide lookups a single question can need — and it grows every time a
    // probe stops holding a handle across a sync, because a handle per batch
    // is a lookup per batch. The longest are `tags-add-same-key-twice` and
    // `tag-on-group-survives` at four batches each, plus the liveness check,
    // plus the verify inside `addScratchSlide` for a replacement: six.
    //
    // It has to be the SMALLEST lease that lets a replacement slide carry a
    // whole question, and no larger. Too tight and the sheet comes back
    // incomplete for a reason that has nothing to do with recovery; too loose
    // and the fault never bites — which is how a count-based version of this
    // once passed against the very code it was written to falsify.
    faults.newSlideResolvesTimes = 6;
    try {
      const sheet = await runHostProbes("fake-loses-slides", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      // Not "nothing crashed" — the same answers a healthy host gives, because
      // every question really was asked.
      expect(answers).toEqual(FAKE_BASELINE);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  it("takes every replacement scratch slide back out again", async () => {
    // A run that recovers by adding slides is a run that can litter several. It
    // has to clean up all of them, not the last one it happened to be holding.
    installHost([makeSlide("s1"), makeSlide("s2")]);
    const { slideCount } = await import("../src/render/powerpoint");
    const before = await slideCount();
    faults.newSlideResolvesTimes = 2;
    try {
      await runHostProbes("fake-loses-slides", "test");
      expect(await slideCount(), "left its replacement slides in the deck").toBe(before);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  /**
   * The sheet's whole promise is that it comes back from a misbehaving host.
   *
   * Every QUESTION was bounded for that reason — `PROBE_BUDGET_MS`, "so the
   * sheet from a misbehaving host is the one worth having". The scratch-slide
   * lifecycle around them was not: `addScratchSlide`, the mid-run replacement,
   * and the cleanup all reached `slideIds`/`slideCount`/`deleteSlideById`,
   * whose syncs carried a label and no deadline. One silent deck-listing read
   * and `runHostProbes` never resolved — with `answers` already complete in
   * memory and never returned.
   */
  it("comes back with its sheet even when the host goes silent at cleanup", async () => {
    installHost([makeSlide("s1")]);
    const { _setReadbackTimeoutForTest } = await import("../src/render/powerpoint");
    _setReadbackTimeoutForTest(20);
    faults.wedgeAfterSyncs = 0;
    try {
      const sheet = await Promise.race([
        runHostProbes("fake-goes-silent", "test"),
        new Promise<"never came back">((r) => setTimeout(() => r("never came back"), 800)),
      ]);
      expect(sheet, "the probe run never returned its answers").not.toBe("never came back");
      expect(typeof sheet === "object" && sheet.answers).toHaveLength(PROBE_IDS.length);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
    }
  });

  /**
   * When the host will not keep ANY slide, the sheet has to say that once per
   * question in a word no probe can produce — never `"threw"`, which is a real
   * answer to several of them and would read as a host divergence.
   */
  /**
   * The second half of the same lesson, and it cost a second real sheet.
   *
   * `no-scratch-slide` covers a slide the host will not resolve. It says
   * nothing about a slide that resolves perfectly and then refuses to take a
   * shape — which is what PowerPoint on the web did on 2026-08-04, because
   * `withProbeContext` handed every probe the slide proxy IT had resolved, and
   * a freshly-added slide's by-id handle is good for exactly one sync there.
   * The six questions that resolved a handle of their own were answered; all
   * eight that wrote through the held one failed, and every one of those was
   * recorded as `"threw"` or `"silent"` — both real answers to those questions.
   * `npm run host-diff` duly reported eight host divergences from a sheet that
   * had asked six questions.
   */
  it("keeps asking when a held slide handle is what the host refuses", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
    // Every question answered, despite the fake now refusing a held by-id
    // handle exactly as the host does — because no probe holds one any more.
    // The single `threw` in the baseline is the question that is ABOUT holding
    // one, which is asked on purpose.
    expect(answers).toEqual(FAKE_BASELINE);
    expect(answers["shape-add-held-slide-proxy"], "the question stopped being asked").toBe("threw");
  });

  it("never reports a setup the host refused as an answer to the question", async () => {
    // A probe that could not get its shapes must say so in a word no probe can
    // produce. `"threw"` and `"silent"` are real answers here, and a diff
    // compares them against the fake's — which is how eight questions nobody
    // asked came back looking like eight host divergences.
    installHost([makeSlide("s1")]);
    faults.refuseShapeAdds = true;
    try {
      const sheet = await runHostProbes("fake-refuses-shapes", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      const needShapes = [
        "shape-proxy-survives-one-sync",
        "shape-resolve-held-slide-proxy",
        "shapes-items-count-honest",
        "shapes-items-via-positional-slide",
        "tags-add-same-key-twice",
        "tags-on-fresh-shape",
        "delete-then-lookup",
        "addgroup-returns-usable",
        "group-reports-its-children",
        "tag-on-group-survives",
      ];
      for (const id of needShapes) expect(answers[id], `${id} claimed a host answer`).toBe("no-scratch-shape");
      // And the questions that need no shape are still answered — one refusal
      // must not cost the sheet.
      expect(answers["getcount-populates-same-sync"]).toBe("yes");
      expect(answers["untrack-available"]).toBe("no");
      // A never-asked question is not a divergence, and this is the tool that
      // used to call it one. The shape-add questions ARE divergences — asking
      // whether this host takes a shape is their whole job, and "no" is an
      // answer. Two of the three: the held-handle question already expects a
      // refusal, so a host that refuses every add agrees with the fake there.
      const d = diffAnswers(answers, FAKE_BASELINE);
      expect(d.differ.map((x: { id: string }) => x.id).sort()).toEqual([
        "shape-add-fresh-getitem-slide",
        "shape-add-fresh-slide-proxy",
        "shape-add-positional-slide-proxy",
      ]);
      expect(d.notAsked.map((n: { id: string }) => n.id).sort()).toEqual([...needShapes].sort());
    } finally {
      faults.refuseShapeAdds = false;
    }
  });

  it("says a question was never put, rather than inventing a host answer", async () => {
    installHost([makeSlide("s1")]);
    faults.newSlideResolvesTimes = 1; // one lookup each: addScratchSlide verifies, the probe context fails
    try {
      const sheet = await runHostProbes("fake-loses-slides", "test");
      expect(sheet.answers).toHaveLength(PROBE_IDS.length);
      for (const a of sheet.answers) expect(a.answer, `${a.id} claimed a host answer`).toBe("no-scratch-slide");
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });
});

describe("what the pane says about a probe run", () => {
  it("tells the owner not to bother sending a run that found nothing", async () => {
    // The round trip this removes: download, send, wait for someone to run
    // `host-diff`, hear that the answers are the same as last time. Most probe
    // runs are that.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    expect(sheetNeedsAttention(sheet), "asked for a file that says nothing new").toBe(false);
    expect(describeHostSheet(sheet)).toContain("Nothing new");
  });

  it("names an answer nobody has declared, and asks for the file", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    // A host that answers one question differently from the fake, where the
    // difference is not in KNOWN_DIVERGENCES: the only kind of run worth
    // anyone's attention, and it should not have to be found by eye in a
    // seventeen-row JSON file.
    const changed = {
      ...sheet,
      answers: sheet.answers.map((a) => (a.id === "delete-then-lookup" ? { ...a, answer: "still-there" } : a)),
    };
    expect(sheetNeedsAttention(changed)).toBe(true);
    expect(describeHostSheet(changed)).toContain("NEW: delete-then-lookup");
    expect(describeHostSheet(changed)).toContain("worth sending");
  });

  it("counts a declared divergence as known, not as news", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    // `tag-on-group-survives` is declared: the real host's "no" is withdrawn
    // pending a re-run. A sheet reproducing it is not a finding.
    const known = {
      ...sheet,
      answers: sheet.answers.map((a) => (a.id === "tag-on-group-survives" ? { ...a, answer: "no" } : a)),
    };
    expect(sheetNeedsAttention(known), "treated a declared divergence as news").toBe(false);
    expect(describeHostSheet(known)).toContain("1 known divergence");
  });

  it("asks for the file when a question was never put, however few diverge", async () => {
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const incomplete = {
      ...sheet,
      answers: sheet.answers.map((a) => (a.id === "tags-on-fresh-shape" ? { ...a, answer: "no-scratch-shape" } : a)),
    };
    expect(sheetNeedsAttention(incomplete), "an incomplete sheet looked fine").toBe(true);
    expect(describeHostSheet(incomplete)).toContain("never put");
  });
});

describe("a host that will not name a shape in the batch that made it", () => {
  /**
   * The five questions the 2026-08-05 sheet never put, and why.
   *
   * They are exactly the five that read an id back — `shape-resolve-held-slide-
   * proxy`, `tags-add-same-key-twice`, `addgroup-returns-usable`,
   * `group-reports-its-children`, `tag-on-group-survives` — and no others.
   * Perfect correlation with `idsOf`, and no competing property explains it:
   * batch count does not, since `shapes-items-count-honest` and
   * `delete-then-lookup` span two batches each and were answered while three of
   * the five span two and failed.
   *
   * The id load used to ride in the same sync as the add, to save a round trip.
   * It now costs its own sync, and that is what makes these five answerable on
   * a host like this.
   */
  it("still asks every question, and answers them as a healthy host would", async () => {
    installHost([makeSlide("s1")]);
    faults.noIdInCreatingSync = true;
    // Both, because they are two halves of one behaviour: the load is not
    // answered, AND an unanswered property is unreadable rather than quietly
    // available. Without `strictShapeReads` the fake hands the id over anyway
    // and the first half is inert — which is exactly how a guard comes to pass
    // against the code it was written to falsify.
    faults.strictShapeReads = true;
    try {
      const sheet = await runHostProbes("fake-cannot-name-new-shapes", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      // The property, not the whole sheet: every question was PUT. Demanding
      // the healthy sheet would be wrong here — `strictShapeReads` legitimately
      // changes what some questions answer, and a guard that could not tell a
      // legitimate change from an unasked question is the confusion this file
      // exists to prevent.
      const neverPut = sheet.answers.filter((a) => a.answer === "no-scratch-shape").map((a) => a.id);
      expect(neverPut, "these questions were never put — the ids came back unreadable").toEqual([]);
      // And the five that read an id back reached their own vocabulary.
      for (const id of [
        "shape-resolve-held-slide-proxy",
        "tags-add-same-key-twice",
        "addgroup-returns-usable",
        "group-reports-its-children",
        "tag-on-group-survives",
      ]) {
        expect(NOT_ASKED_WORDS, `${id} did not get asked`).not.toContain(answers[id]);
      }
    } finally {
      faults.noIdInCreatingSync = false;
      faults.strictShapeReads = false;
    }
  });
});

describe("a probe run that has lost its scratch slide", () => {
  /**
   * Losing a slide must cost the questions it happened during, not the sheet.
   *
   * `scratchId` went null when a replacement also failed to resolve, and nothing
   * set it back — so one bad moment recorded `no-scratch-slide` for every
   * question after it, as though the host had been asked each one. That is the
   * same failure this file's whole design is against, one level up: an answer
   * that describes the run's own state rather than the host's.
   */
  it("takes another slide rather than writing off the rest of the sheet", async () => {
    installHost([makeSlide("s1")]);
    // Every new slide answers to its id exactly once: `addScratchSlide`'s own
    // verify spends it, so every probe context then fails its liveness check
    // and every replacement is written off immediately. The run should keep
    // trying, and keep saying honestly that it could not get a slide.
    faults.newSlideResolvesTimes = 1;
    try {
      const sheet = await runHostProbes("fake-loses-every-slide", "test");
      expect(sheet.answers).toHaveLength(PROBE_IDS.length);
      for (const a of sheet.answers) {
        expect(a.answer, `${a.id} claimed a host answer`).toBe("no-scratch-slide");
      }
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  it("comes back at all when the host never answers the slide add", async () => {
    // `addScratchSlide` used to wrap `slides.add()` in a raw `await
    // context.sync()` — the only slide-add in the file without a deadline.
    // office-js#1650 is explicit that the promise can hang while the slide
    // lands, and a probe run that hangs there produces no sheet at all, which
    // is the one thing this diagnostic must never do.
    installHost([makeSlide("s1")]);
    const { _setReadbackTimeoutForTest } = await import("../src/render/powerpoint");
    _setReadbackTimeoutForTest(20);
    faults.wedgeAfterSyncs = 1;
    try {
      const sheet = await Promise.race([
        runHostProbes("fake-hangs-on-add", "test"),
        new Promise<"never came back">((r) => setTimeout(() => r("never came back"), 2000)),
      ]);
      expect(sheet, "the run hung on the slide add and produced no sheet").not.toBe("never came back");
      expect(typeof sheet === "object" && sheet.answers).toHaveLength(PROBE_IDS.length);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
    }
  });
});
