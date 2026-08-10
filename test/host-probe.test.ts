// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, applyWebProfile, faults } from "./helpers/office-host";
import {
  runHostProbes,
  PROBE_IDS,
  ALWAYS_ASKED_IDS,
  SCRATCH_CLEANUP_ID,
  NO_SLIDE_NEEDED_IDS,
  describeHostSheet,
  sheetNeedsAttention,
  summariseHostSheet,
  _setProbeBudgetForTest,
} from "../src/render/host-probe";
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
    // Two-sided, because a follow-up is conditional and a single list cannot
    // say both things. Every question a run ALWAYS puts is here, and nothing
    // here is a question this build does not know how to ask — an id outside
    // `PROBE_IDS` would be an answer the baseline could never account for.
    expect(Object.keys(sheet).sort()).toEqual([...ALWAYS_ASKED_IDS].sort());
    for (const id of Object.keys(sheet)) expect(PROBE_IDS, `${id} is not a question this build asks`).toContain(id);
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
    expect(d.agree.length).toBe(ALWAYS_ASKED_IDS.length);
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
    expect(sheet.answers).toHaveLength(ALWAYS_ASKED_IDS.length);
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
      // COMPLETENESS is the invariant, and it is the one this test is named for:
      // every question got put, on a host that takes the slide away between
      // them. Nothing here may come back in the never-asked vocabulary.
      expect(Object.keys(answers).sort()).toEqual([...ALWAYS_ASKED_IDS].sort());
      for (const [id, answer] of Object.entries(answers))
        expect(NOT_ASKED_WORDS, `${id} was never put — the replacement path did not carry it`).not.toContain(answer);
      // This USED to assert `answers` equalled `FAKE_BASELINE` outright, and
      // that assertion is gone deliberately rather than by accident.
      //
      // The lease is per SLIDE and questions share one until it fails, so how
      // much budget a question inherits depends on what ran before it. Moving
      // `binding-names-shape-later` from position 20 to 6 therefore left
      // `shape-resolve-held-slide-proxy` a partly-spent slide, and it answered
      // `threw` instead of `yes` — a correct answer to "can you resolve a shape
      // through a handle this host has stopped honouring", and a real one, not
      // a never-asked. Value-equality here was quietly pinning probe ORDER.
      //
      // Raising the lease until it went green is the one thing not done: the
      // comment above sets it as the SMALLEST lease that carries the longest
      // question, and tuning it to accommodate a reorder is how a count-based
      // version of this test once passed against the code it was written to
      // falsify. The values are still pinned, on the healthy host, by "answers
      // every question, and says what it claims" above.
      expect(answers["shape-add-fresh-slide-proxy"], "a question early enough to be unaffected drifted").toBe(
        FAKE_BASELINE["shape-add-fresh-slide-proxy"],
      );
      // KNOWN HOLE, and it predates this edit: nothing here proves the
      // REPLACEMENT path ran, though the test is named for it.
      //
      // Measured, not suspected. Disabling the post-not-asked replacement
      // outright — `if (recovered)` never taken, and `addScratchSlide` forced to
      // null — leaves this test green, and left it green against the
      // `toEqual(FAKE_BASELINE)` assertion that used to stand here too. At a
      // lease of six the original slide carries almost every question, so the
      // recovery is never needed and never exercised.
      //
      // A `took > 1` check on the cleanup row was tried and thrown away: it
      // passes under both mutations, because the second pass and the top-of-loop
      // re-acquire take slides of their own. A guard that survives the mutation
      // it is named for is decoration.
      //
      // Closing it needs a fault that forces a not-asked and then relents —
      // `newSlideRefusedForFirst` is the shape of it — which is its own change,
      // not a rider on a probe reorder.
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
    _setProbeBudgetForTest(20);
    try {
      const sheet = await Promise.race([
        runHostProbes("fake-goes-silent", "test"),
        new Promise<"never came back">((r) => setTimeout(() => r("never came back"), 800)),
      ]);
      expect(sheet, "the probe run never returned its answers").not.toBe("never came back");
      expect(typeof sheet === "object" && sheet.answers).toHaveLength(ALWAYS_ASKED_IDS.length);
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
        "group-children-via-getcount",
        "shape-proxy-survives-one-sync",
        "shape-resolve-held-slide-proxy",
        "shapes-items-count-honest",
        "shapes-items-via-positional-slide",
        "tags-add-same-key-twice",
        "tags-on-fresh-shape",
        "delete-then-lookup",
        "addgroup-returns-usable",
        "group-reports-its-children",
        "group-of-existing-shape-readable",
        "picture-then-shape-read",
        "tag-on-group-survives",
        // The binding question needs a shape to bind, so a host that refuses
        // every add has told it nothing. Its first version answered `add-threw`
        // here — a statement about `bindings.add` from a run in which
        // `bindings.add` was never reached — which is why the shape add, the
        // binding call and the commit are now three separate failure points.
        "binding-names-shape-later",
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
      expect(sheet.answers).toHaveLength(ALWAYS_ASKED_IDS.length);
      // Every question that NEEDS a slide. Two rows are not those:
      //
      // - the cleanup row, which reports what became of the slides the run
      //   borrowed — a fact about the host whether or not a single question got
      //   put, and reading it as an answer would make this assertion demand
      //   that the cleanup fail too;
      // - a `noSlideNeeded` question, which is answerable with no slide at all
      //   and must not inherit this failure. Charging it one is how the
      //   2026-08-08 `a546897` sheet lost `untrack-available`, at 43 seconds,
      //   after the host had recovered and answered four questions in a row.
      //
      // The excused set is DERIVED from the probes rather than listed here, so
      // it cannot drift from what the code actually declares — and the loop
      // below is what stops that from being a way to silence this test. An
      // excusal has to be earned: a question marked `noSlideNeeded` must come
      // back with a real answer on a host that has no slide to give, and one
      // that answers `no-scratch-slide` anyway is mismarked and says so here.
      const excused = new Set([SCRATCH_CLEANUP_ID, ...NO_SLIDE_NEEDED_IDS]);
      for (const a of sheet.answers.filter((x) => !excused.has(x.id)))
        expect(a.answer, `${a.id} claimed a host answer`).toBe("no-scratch-slide");
      expect(NO_SLIDE_NEEDED_IDS.length, "no question claims to be answerable without a slide").toBeGreaterThan(0);
      for (const id of NO_SLIDE_NEEDED_IDS) {
        const a = sheet.answers.find((x) => x.id === id);
        expect(a, `${id} is marked noSlideNeeded and produced no row at all`).toBeTruthy();
        expect(NOT_ASKED_WORDS, `${id} is marked noSlideNeeded but was charged for a slide anyway`).not.toContain(
          a!.answer,
        );
      }
      // And it has to be there. A run that could not ask anything is precisely
      // the run that churns through scratch slides, so a sheet silent about
      // them is silent about the only thing that run left behind.
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID);
      expect(cleanup, "the sheet says nothing about the slides the run borrowed").toBeTruthy();
      expect(cleanup!.detail).toMatch(/scratch slide/);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  });

  it("puts the questions a dead window swallowed, once the host is answering again", async () => {
    // Two consecutive real rounds lost roughly HALF their questions this way —
    // 14 of 27, then 13 of 28 — every one of them `no-scratch-slide`, and the
    // immediate retry beside each failure fired 21 times between them and
    // changed nothing. It could not: this host's ability to resolve a freshly
    // added slide comes and goes in windows of about fifteen seconds, so a retry
    // issued straight away lands inside the window that just refused.
    //
    // What both rounds also show is the recovery, in the same run: 2026-08-09
    // answered positions 17, 18, 22, 24 and 26 after losing 10 through 16. So
    // the questions were answerable; the run simply asked them all at the wrong
    // moment and never went back.
    //
    // `newSlideRefusedForFirst` is that window, counted in slides because
    // nothing the probe does consults a clock. Six is comfortably wider than the
    // handful of slides the opening questions spend, so questions genuinely fall
    // into it — and it closes, which is the half the fake could not express
    // before.
    installHost([makeSlide("s1")]);
    faults.newSlideRefusedForFirst = 12;
    try {
      const sheet = await runHostProbes("fake-loses-then-recovers", "test");
      const lost = sheet.answers.filter((a) => NOT_ASKED_WORDS.includes(a.answer));
      const recovered = sheet.answers.filter((a) => a.detail?.includes("second pass"));
      expect(recovered.length, "the second pass rescued nothing — the window never closed for it").toBeGreaterThan(5);
      expect(
        lost.map((a) => a.id),
        "a question the host would have answered was still filed unanswerable",
      ).toEqual([]);
      // A rescued row carries a real answer, not a relabelled failure.
      for (const a of recovered)
        expect(NOT_ASKED_WORDS, `${a.id} "recovered" into ${a.answer}`).not.toContain(a.answer);
    } finally {
      faults.newSlideRefusedForFirst = 0;
    }
  });

  /**
   * The binding question's two ways of being refused, and why they are two.
   *
   * A binding has to be made in the batch that CREATES its shape — a proxy one
   * sync old is refused on this host — so the batch carries two things and its
   * failure is attributable to neither. That is not a hypothetical: the
   * 2026-08-09 evening round came back `UnexpectedError` from the commit in 1.3
   * seconds, and the probe honestly reported "never asked" while
   * `shape-add-fresh-slide-proxy` answered `yes` two rows above it.
   *
   * The control arm is what makes the difference readable. The same batch minus
   * the binding runs first, on the same slide; if it commits and the bound one
   * does not, the binding is the only variable left.
   */
  it("blames the binding only when the same batch without one just worked", async () => {
    installHost([makeSlide("s1")]);
    faults.refuseBindings = "sync";
    try {
      const sheet = await runHostProbes("fake-refuses-bindings-at-sync", "test");
      const row = sheet.answers.find((a) => a.id === "binding-names-shape-later");
      expect(row?.answer, "a refusal the control had already ruled out is an ANSWER, not a missing question").toBe(
        "commit-threw",
      );
      expect(row?.detail).toMatch(/without a binding committed seconds earlier/);
    } finally {
      faults.refuseBindings = null;
    }
  });

  it("tells a binding refused at the call apart from one refused at the commit", async () => {
    // Same API, two different facts about it: `bindings.add` objecting on the
    // spot is not the host rejecting the batch that carried it, and a single
    // word for both would put them in one bucket in the diff.
    installHost([makeSlide("s1")]);
    faults.refuseBindings = "call";
    try {
      const sheet = await runHostProbes("fake-refuses-bindings-at-call", "test");
      const row = sheet.answers.find((a) => a.id === "binding-names-shape-later");
      expect(row?.answer).toBe("add-threw");
      expect(row?.detail).toMatch(/BindingCollection\.add/);
    } finally {
      faults.refuseBindings = null;
    }
  });

  it("takes a fresh slide after a question that wrecks the one it used", async () => {
    // The finding seven rounds took to see, and only because a reorder ran the
    // control by accident.
    //
    // The slot after `shape-add-held-slide-proxy` has never produced an answer.
    // For six rounds that slot held `shape-resolve-held-slide-proxy`, and its
    // 0/6 was put down to its own nature — it reads ids back off setup shapes,
    // which is a known-bad thing to do here. Then `binding-names-shape-later`
    // was moved into the slot for unrelated reasons and failed identically, in
    // 397ms, at the liveness check, before its own code ran:
    //
    //   12-17  #5 shape-add-held-slide-proxy threw | #6 shape-resolve-…  never put
    //   18     #5 shape-add-held-slide-proxy threw | #6 binding-names-…  never put
    //
    // Two questions, one slot, seven for seven: it is the slot. #5 writes
    // through a slide proxy the host has stopped honouring — that IS its
    // question — and the slide does not survive it. Because it ANSWERS, none of
    // the not-asked replacement paths ever noticed.
    //
    // Asserted structurally rather than through a fault, because the fake does
    // not model the poisoning and inventing it would be asserting a mechanism
    // nobody has measured. What IS measured is the rule: a question that
    // declares it wrecks the slide must not hand that slide to the next one.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("fake", "test");
    const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID)!;
    const took = Number(/of (\d+) scratch/.exec(cleanup.detail ?? "")?.[1] ?? 0);
    // A healthy fake loses nothing, so without the rule ONE slide serves the
    // whole run. Any second slide here is the burnt one being given up.
    expect(took, `the run never replaced the wrecked slide — "${cleanup.detail}"`).toBeGreaterThan(1);
  });

  it("keeps sweeping after the host refuses it a slide, instead of bailing on the first", async () => {
    // Round 15 is this bug, in the owner's own deck: `second pass over the
    // questions that were never put {count: 10}`, ONE slide attempt two seconds
    // later, `scratch slide landed but its id will not resolve`, and then
    // nothing. Ten questions decided by a single coin flip on a host whose
    // slide resolution is known to flap — and the `break` was silent, so the
    // log could not tell "tried them all and failed" from "gave up at once".
    //
    // A window of 28 is the discriminator, measured rather than picked: it
    // closes DURING the sweep, so the first attempts are refused and later ones
    // succeed. With the loop continuing past a refusal that rescues 19 of 20;
    // bailing on the first rescues none of them.
    installHost([makeSlide("s1")]);
    faults.newSlideRefusedForFirst = 28;
    try {
      const sheet = await runHostProbes("fake-refuses-then-relents-mid-sweep", "test");
      const rescued = sheet.answers.filter((a) => a.detail?.includes("second pass"));
      expect(rescued.length, "one refused slide ended the whole sweep").toBeGreaterThan(10);
      const lost = sheet.answers.filter((a) => NOT_ASKED_WORDS.includes(a.answer));
      expect(lost.length, `still lost: ${lost.map((a) => a.id).join(", ")}`).toBeLessThan(3);
    } finally {
      faults.newSlideRefusedForFirst = 0;
    }
  });

  it("gives up on a second pass the host is never going to answer", async () => {
    // The other side of the rung, and the one that keeps it from being a way to
    // spend three minutes on a dead host. A window wider than the whole run
    // never closes, so every question stays unasked and the sweep must stop
    // asking rather than work through the list. The sheet still comes back
    // complete — that is the invariant every other case here defends too.
    installHost([makeSlide("s1")]);
    faults.newSlideRefusedForFirst = 500;
    try {
      const sheet = await runHostProbes("fake-never-recovers", "test");
      expect(sheet.answers).toHaveLength(ALWAYS_ASKED_IDS.length);
      expect(sheet.answers.filter((a) => a.detail?.includes("second pass"))).toEqual([]);
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID);
      expect(cleanup, "a run that asked nothing still owes an account of the slides it borrowed").toBeTruthy();
    } finally {
      faults.newSlideRefusedForFirst = 0;
    }
  });

  it("reports the scratch slides it could NOT give back", async () => {
    // The failure this row exists for, and it is not hypothetical: a real round
    // on 2026-08-06 left 21 blank slides in the owner's deck, an earlier one
    // left 14, and both sheets came back saying nothing about it. The cleanup
    // loop already asked `deleteSlideById`, which re-reads the deck rather than
    // trusting a queued delete — it simply threw the boolean away, so the only
    // way to learn the probe litters was to open the deck afterwards.
    installHost([makeSlide("s1")]);
    faults.refuseSlideDelete = true;
    try {
      const sheet = await runHostProbes("fake-refuses-deletes", "test");
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID);
      expect(cleanup, "the sheet says nothing about the slides the run borrowed").toBeTruthy();
      expect(cleanup!.answer, `the host kept every slide and the sheet said "${cleanup!.answer}"`).toBe("none");
      expect(cleanup!.detail).toMatch(/left in the deck/);
    } finally {
      faults.refuseSlideDelete = false;
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
      expect(sheet.answers).toHaveLength(ALWAYS_ASKED_IDS.length);
      // The cleanup row and the slideless question aside — see the sibling case
      // above for both. One is not a question; the other does not need what
      // this fault withholds.
      const excused = new Set([SCRATCH_CLEANUP_ID, ...NO_SLIDE_NEEDED_IDS]);
      for (const a of sheet.answers.filter((x) => !excused.has(x.id))) {
        expect(a.answer, `${a.id} claimed a host answer`).toBe("no-scratch-slide");
      }
      expect(sheet.answers.find((a) => a.id === "untrack-available")?.answer).toBe("no");
      // This is the run that makes one scratch slide per question and throws
      // each away. How many came back is the whole finding, and until this row
      // existed the sheet did not carry it.
      const cleanup = sheet.answers.find((a) => a.id === SCRATCH_CLEANUP_ID)!;
      expect(cleanup.detail, "the sheet counted no scratch slides at all").toMatch(/\d+ of \d+ scratch slide/);
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
      expect(typeof sheet === "object" && sheet.answers).toHaveLength(ALWAYS_ASKED_IDS.length);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
    }
  });
});

/**
 * The partner question, asked in the same run.
 *
 * The repo's rule is "when two explanations fit the evidence, ask — do not
 * reason". Following it has meant writing the partner question, waiting for the
 * owner to run the probe again, and losing a session — twice, at a cost of two
 * full sheets. The reasoning was never the expensive part; the round trip was.
 */
describe("questions that ask their own follow-up", () => {
  const sheetRows = async () => (await runHostProbes("fake", "test")).answers;

  it("asks it when the answer admits two readings, and says which answer caused it", async () => {
    // `getItem` refusing a freshly-added slide has two readings that lead
    // opposite ways: getItem cannot name a NEW slide (so the everyday insert
    // path is fine and only one caller is at risk), or getItem is broken here
    // (so the insert path is broken for everyone). A sheet cannot tell them
    // apart; the pre-existing slide can.
    installHost([makeSlide("s1")]);
    faults.refuseGetItemOnNewSlide = true;
    try {
      const rows = await sheetRows();
      const asked = rows.find((r) => r.id === "shape-add-fresh-getitem-slide")!;
      const partner = rows.find((r) => r.id === "getitem-durable-slide");
      expect(asked.answer, "the fault did not provoke the answer this pair is about").not.toBe("yes");
      expect(partner, "the follow-up was never asked").toBeTruthy();
      // The discriminating result: getItem works perfectly on a slide that was
      // already in the deck, so the refusal is about the slide's newness.
      expect(partner!.answer).toBe("yes");
      expect(partner!.detail, "the sheet does not say the two rows are a pair").toContain(
        'asked because shape-add-fresh-getitem-slide answered "threw"',
      );
    } finally {
      faults.refuseGetItemOnNewSlide = false;
    }
  });

  it("does not ask it when there is nothing to disambiguate", async () => {
    // An unconditional partner is just another probe and belongs in the list.
    // A follow-up earns its place by being worth asking only in the light of a
    // particular answer — so on a host that answers plainly it must cost
    // nothing at all.
    installHost([makeSlide("s1")]);
    const rows = await sheetRows();
    expect(rows.find((r) => r.id === "shape-add-fresh-getitem-slide")!.answer).toBe("yes");
    expect(
      rows.some((r) => r.id === "getitem-durable-slide"),
      "asked a follow-up nobody needed",
    ).toBe(false);
  });

  it("does not follow up a question that was never put", async () => {
    // `no-scratch-slide` is the probe's own vocabulary for "this never reached
    // its question". A partner to that would be a second question about the
    // probe's own setup, dressed as a fact about the host — precisely the
    // confusion `NOT_ASKED` exists to prevent, and worse here because the
    // follow-up's answer looks like a real finding in its own right.
    installHost([makeSlide("s1")]);
    faults.swallowAdds = 500;
    try {
      const rows = await sheetRows();
      expect(rows.find((r) => r.id === "shape-add-fresh-getitem-slide")!.answer).toBe("no-scratch-slide");
      expect(
        rows.some((r) => r.id === "getitem-durable-slide"),
        "followed up a question the host was never asked",
      ).toBe(false);
    } finally {
      faults.swallowAdds = 0;
    }
  });

  it("says it had no control rather than measuring a slide this run added", async () => {
    // On a deck the run has built entirely there is no pre-existing slide, and
    // the honest answer is that the control was unavailable. Falling back to a
    // scratch slide would be the control measuring the very thing it exists to
    // be compared against — the exact confusion the pair removes.
    installHost([]);
    faults.refuseGetItemOnNewSlide = true;
    try {
      const rows = await sheetRows();
      expect(rows.find((r) => r.id === "getitem-durable-slide")?.answer).toBe("no-durable-slide");
    } finally {
      faults.refuseGetItemOnNewSlide = false;
    }
  });

  it("still produces a sheet when the deck will not list itself at all", async () => {
    // The control read happens before the first question, so an unbounded one
    // took the whole run down with it — no sheet at all, which is the single
    // failure mode this file exists to prevent. A control nobody can name costs
    // the follow-up and nothing else.
    installHost([makeSlide("s1")]);
    const { _setReadbackTimeoutForTest } = await import("../src/render/powerpoint");
    faults.wedgeAfterSyncs = 0;
    _setReadbackTimeoutForTest(20);
    try {
      const ids = new Set((await sheetRows()).map((r) => r.id));
      for (const id of ALWAYS_ASKED_IDS) {
        expect(ids, `lost ${id} because the deck would not list itself`).toContain(id);
      }
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
    }
  });
});

describe("the questions added from the office-js tracker", () => {
  const sheetOfRows = async () => (await runHostProbes("fake", "test")).answers;

  it("asks about layouts on a slide the run did not add", async () => {
    // A freshly-added slide's handle is good for exactly one sync on this host,
    // so a load queued on the scratch slide and read after its own sync answers
    // "unreadable" whatever the layout API does. That is an answer about the
    // probe's plumbing wearing the clothes of a fact about layouts — the
    // failure this whole file exists to prevent, and the first draft of this
    // question had it.
    installHost([]);
    const rows = await sheetOfRows();
    expect(rows.find((r) => r.id === "slide-layout-readable")?.answer).toBe("no-durable-slide");
  });

  it("still answers the layout question when the deck has a settled slide", async () => {
    installHost([makeSlide("s1")]);
    const rows = await sheetOfRows();
    expect(rows.find((r) => r.id === "slide-layout-readable")?.answer).toBe("yes");
  });

  it("names a group's children through a handle resolved later, not the one that made it", async () => {
    // office-js#5849 is about `Shape.group` on a group found afterwards, which
    // is how `countGroupChildrenPage` finds one. `group-reports-its-children`
    // already asks in the batch that MADE the group, and the two answers are
    // different claims — one about the API, one about proxy age.
    installHost([makeSlide("s1")]);
    const rows = await sheetOfRows();
    expect(rows.find((r) => r.id === "group-of-existing-shape-readable")?.answer).toBe("2");
  });
});

describe("a host that answers every call and reads back nothing", () => {
  /**
   * The third kind of bad host, and the one with no test until now.
   *
   * Two are covered already: a host that THROWS (`refuseShapeAdds`) and one that
   * goes SILENT (`wedgeAfterSyncs`). This is the one in between — every call is
   * taken, every sync resolves, and then the properties those syncs were
   * supposed to populate are unreadable. It is the shape PowerPoint on the web
   * shows most often, and the shape a probe is most likely to mis-report:
   * "the value is not there" and "the value is false" are one line apart in
   * every one of these questions.
   *
   * The property, not the sheet. What each question answers here is allowed to
   * differ — an unreadable host is a legitimate finding — but a question that
   * silently vanishes, or an answer that is not a word, is the probe failing
   * rather than the host.
   */
  it("still puts every question, and answers each from its own vocabulary", async () => {
    installHost([makeSlide("s1")]);
    faults.strictShapeReads = true;
    faults.unansweredTagLoads = 500;
    faults.hollowNameReads = 500;
    try {
      const sheet = await runHostProbes("fake-unreadable", "test");
      const ids = sheet.answers.map((a) => a.id);
      for (const id of ALWAYS_ASKED_IDS) {
        expect(ids, `${id} vanished from the sheet on an unreadable host`).toContain(id);
      }
      for (const a of sheet.answers) {
        expect(typeof a.answer, `${a.id} answered with a ${typeof a.answer}`).toBe("string");
        expect(a.answer.length, `${a.id} answered with an empty string`).toBeGreaterThan(0);
        // A JS `undefined` reaching the sheet as a word is the specific bug this
        // guards. Every one of these questions reads a property that may not be
        // there, and `String(undefined)` is a perfectly good-looking answer that
        // the diff would compare against a real host forever.
        expect(a.answer, `${a.id} put a JS undefined in the sheet`).not.toMatch(/^undefined$/);
      }
    } finally {
      faults.strictShapeReads = false;
      faults.unansweredTagLoads = 0;
      faults.hollowNameReads = 0;
    }
  });
});

/**
 * The probe stops asking a host that has stopped answering.
 *
 * The self-test has had this rung for weeks; the probe had none, and
 * PowerPoint's own "Sorry, we ran into a problem" dialog is what exposed the
 * gap — the host was gone, the pane's elapsed timer kept counting, and the run
 * went on putting questions to a dead document, each one spending its full
 * budget to record an answer about nothing.
 *
 * Only the NO-FALSE-POSITIVE half is guarded here, and that is worth stating
 * rather than leaving to be discovered. Reaching the positive case needs a host
 * that takes a question and never answers, and the fake's wedge hangs
 * `addScratchSlide` first — the run stops before any question can miss its
 * deadline, so the test times out instead of exercising the breaker. A fault
 * that wedges question syncs while leaving slide adds alone is what it would
 * take, and that is a change to the fake with its own justification owed.
 */
describe("a host that has stopped answering", () => {
  /**
   * Why the breaker counts deadline misses and not `no-scratch-slide`.
   *
   * A round on 2026-08-08 answered `no-scratch-slide` FOUR times running — the
   * host had stopped resolving fresh slide ids for about fifteen seconds — then
   * recovered and answered five more questions, including two of the three this
   * project cares most about. Those questions were still being answered, in two
   * to six seconds; they just could not be set up. A breaker keyed on that
   * signal throws the recovery away.
   */
  it("keeps going through a host that refuses slides but still answers promptly", async () => {
    installHost([makeSlide("s1")]);
    faults.newSlideResolvesTimes = 1;
    try {
      const sheet = await runHostProbes("flaky-slides", "web");
      const notReached = sheet.answers.filter((a) => a.answer === "not-asked");
      expect(
        notReached,
        `gave up on a host that was still answering — ${notReached.length} question(s) abandoned`,
      ).toHaveLength(0);
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  }, 60_000);
});

/**
 * The three group questions, on the host that could never answer them.
 *
 * They were unanswerable on the one host that matters, and `scratchShapes`' own
 * comment had already named the lead: the six questions this host never answers
 * are exactly the six that pass a `load`. They add two shapes, sync, read the
 * ids back — and it is that read the host refuses, the same refusal behind
 * `shape-proxy-survives-one-sync: unreadable` and the empty collection reads
 * everywhere else. The questions were never reached. They died in their own
 * setup and reported `no-scratch-slide`, which describes the probe rather than
 * the host.
 *
 * They matter more than their position in the list suggests. `contentShapes`
 * returns UNKNOWN_CONTENT for every grouped slide, which is what makes the
 * reconcile report a slide complete without counting it — so whether the web
 * host can count a group's children decides whether those verdicts can ever be
 * measurements. office-js#5849 is closed for inactivity and is a DESKTOP report;
 * nobody has established the web answer.
 */
describe("the group questions, when the host will not read an id back", () => {
  const GROUP_QUESTIONS = ["addgroup-returns-usable", "group-reports-its-children", "group-of-existing-shape-readable"];

  it("still asks them, and says which route the members came by", async () => {
    installHost([makeSlide("s1")]);
    // The host from the run logs: it will not read an id back off a shape this
    // run just added, however young the proxy — which is exactly what the strict
    // setup depends on. A large count, because every group question tries it.
    faults.refuseShapeIdLoads = 500;
    try {
      const sheet = await runHostProbes("no-id-readback", "web");
      const byId = Object.fromEntries(sheet.answers.map((a) => [a.id, a]));
      for (const id of GROUP_QUESTIONS) {
        const a = byId[id];
        expect(a, `${id} is missing from the sheet entirely`).toBeTruthy();
        expect(
          ["no-scratch-slide", "no-scratch-shape", "not-asked"],
          `${id} still died in its own setup: ${a.answer} — ${a.detail ?? "no detail"}`,
        ).not.toContain(a.answer);
        // A fallback that does not SAY it is a fallback is the trap this whole
        // file is about: grouping same-batch proxies is the friendliest case a
        // host can be given, and an answer from it says nothing about grouping
        // by id. The sheet has to keep them apart.
        expect(a.detail ?? "", `${id} did not record how its members were obtained`).toMatch(
          /members via (ids|same-batch)/,
        );
      }
    } finally {
      faults.refuseShapeIdLoads = 0;
    }
  }, 60_000);

  it("prefers the strict route when the host allows it", async () => {
    // The negative control. If the fallback ran unconditionally these questions
    // would quietly stop testing what production does — grouping by id — and
    // nothing would say so.
    installHost([makeSlide("s1")]);
    const sheet = await runHostProbes("healthy", "web");
    const byId = Object.fromEntries(sheet.answers.map((a) => [a.id, a]));
    for (const id of GROUP_QUESTIONS) {
      expect(byId[id]?.detail ?? "", `${id} took the fallback on a host that never needed it`).toContain(
        "members via ids",
      );
    }
  }, 60_000);
});

describe("a question that needs no slide", () => {
  /**
   * The liveness check `withProbeContext` runs before every question is, on
   * PowerPoint web, the single most likely thing to fail: a fresh slide's id
   * resolves once and then stops. Charging it to a question that never touches
   * the slide turns a real answer into `no-scratch-slide`, which is a statement
   * about the probe wearing the clothes of a statement about the host.
   *
   * That is not hypothetical. The 2026-08-08 `a546897` round lost
   * `untrack-available` this way at 43 seconds — AFTER the host had recovered
   * and answered four questions in a row. `untrack` is a `typeof` on a proxy
   * that `getItemOrNullObject` builds without a round trip, so the answer was
   * there for the taking.
   */
  it("answers even when the scratch slide has stopped resolving", async () => {
    installHost([makeSlide("s1")]);
    // Every slide this run adds is dead on arrival, so the liveness check can
    // only fail. Anything that still answers did so without one.
    faults.newSlideResolvesTimes = 0;
    try {
      const sheet = await runHostProbes("fake-dead-scratch", "test");
      const answers = Object.fromEntries(sheet.answers.map((a) => [a.id, a.answer]));
      expect(answers["untrack-available"], "a question that needs no slide was charged for one").toBe("no");
      // Non-vacuity: the scratch slide really is unusable, so this is not a run
      // where the check simply passed.
      expect(answers["shape-add-fresh-slide-proxy"]).toBe("no-scratch-slide");
    } finally {
      faults.newSlideResolvesTimes = null;
    }
  }, 60_000);
});

describe("the pane's own read of a sheet", () => {
  /**
   * A declared answer must not be announced as news, wherever it is declared.
   *
   * `shape-add-positional-slide-proxy` is a coin — `yes, yes, threw, yes, yes,
   * threw` across six rounds — and the fake says `yes`. The summary counted
   * only `KNOWN_DIVERGENCES`, so every round the coin landed `threw` the pane
   * said "NEW: shape-add-positional-slide-proxy", about the single question
   * this repo documents at greatest length as varying run to run. A gate that
   * cries wolf on a schedule is one people stop reading.
   */
  it("does not call a documented coin flip a new divergence", () => {
    const sheet = {
      kind: "powerchart-host-answers" as const,
      source: "test",
      build: "test",
      requirementSets: [],
      answers: [
        // The coin, landed on the side the fake does not take.
        { id: "shape-add-positional-slide-proxy", question: "?", answer: "threw", ms: 1 },
        // Non-vacuity: something genuinely undeclared, so this is not a sheet
        // in which nothing could have been reported either way.
        { id: "getitemat-past-end", question: "?", answer: "surprise", ms: 1 },
      ],
    };
    const s = summariseHostSheet(sheet);
    expect(s.fresh, "a documented unstable answer was announced as news").toEqual(["getitemat-past-end"]);
    expect(s.known).toContain("shape-add-positional-slide-proxy");
  });
});
