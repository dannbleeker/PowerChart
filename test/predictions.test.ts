import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
// @ts-expect-error — plain .mjs tool, no types.
import { judgePrediction, roundToJudgeOn } from "../scripts/triage.mjs";
// Its own line — see the grouped-import trap documented in `triage.test.ts`.
// @ts-expect-error — as above.
import { scenarioHistory } from "../scripts/triage.mjs";

/**
 * A prediction that cannot fail is not worth staking, and a judge that cannot
 * say "undetermined" turns every honest unknown into a false verdict.
 */
describe("judging a prediction against a round", () => {
  const sheet = (answers: { id: string; answer: string; detail?: string }[]) => ({ hostAnswers: { answers } });

  it("holds when the named questions answered", () => {
    const r = judgePrediction(
      { claim: { kind: "probe-answers", ids: ["a", "b"] } },
      sheet([
        { id: "a", answer: "yes" },
        { id: "b", answer: "threw" },
      ]),
    );
    expect(r.verdict).toBe("held");
  });

  it("FAILS when a question the prediction promised would answer did not", () => {
    // `threw` is an answer; `no-scratch-shape` is the question never being put.
    // Reading the second as an answer is the mistake this whole vocabulary
    // exists to prevent, so the judge must not make it either.
    const r = judgePrediction(
      { claim: { kind: "probe-answers", ids: ["a", "b"] } },
      sheet([
        { id: "a", answer: "yes" },
        { id: "b", answer: "no-scratch-shape" },
      ]),
    );
    expect(r.verdict).toBe("FAILED");
    expect(r.why).toContain("b=no-scratch-shape");
  });

  it("holds a prediction that nothing would change, which is how a change is shown to be scoped", () => {
    const r = judgePrediction(
      { claim: { kind: "probe-starves", ids: ["a"] } },
      sheet([{ id: "a", answer: "no-scratch-slide" }]),
    );
    expect(r.verdict).toBe("held");
  });

  it("says undetermined rather than FAILED when the question was never put", () => {
    // The detail of a never-put question describes the refusal, not the answer.
    // Matching a pattern against it would blame the prediction for the host
    // declining to be asked.
    const r = judgePrediction(
      { claim: { kind: "probe-detail-matches", id: "a", pattern: "slide stable" } },
      sheet([{ id: "a", answer: "no-scratch-shape", detail: "the host says the scratch slide is gone" }]),
    );
    expect(r.verdict).toBe("undetermined");
  });

  it("says undetermined when the round's sheet does not carry the question at all", () => {
    const r = judgePrediction(
      { claim: { kind: "probe-answers", ids: ["missing"] } },
      sheet([{ id: "a", answer: "yes" }]),
    );
    expect(r.verdict).toBe("undetermined");
  });

  it("judges a detail pattern when the question WAS answered", () => {
    const answered = sheet([{ id: "a", answer: "other", detail: "value=undefined; slide stable (259#3258414649)" }]);
    expect(
      judgePrediction({ claim: { kind: "probe-detail-matches", id: "a", pattern: "slide stable" } }, answered).verdict,
    ).toBe("held");
    const changed = sheet([
      { id: "a", answer: "other", detail: "value=undefined; slide CHANGED under the probe: 1 -> 2 -> 3" },
    ]);
    expect(
      judgePrediction({ claim: { kind: "probe-detail-matches", id: "a", pattern: "slide stable" } }, changed).verdict,
    ).toBe("FAILED");
  });

  it("has NO round to judge on when the build it was staked on has never been rounded", () => {
    // The false-HELD case, and it happened the first time a prediction was
    // staked the moment a change landed. The rule is "the newest round taken
    // AFTER the prompting build"; when that build is not in the archive at all
    // the old code fell back to judging against EVERY round, i.e. the newest —
    // which is older than the change. `same scale across the deck` passes in the
    // recent archive, so a prediction about code that had never been run came
    // out `held` on evidence recorded before it existed.
    const logs = [{ build: "aaaaaaa" }, { build: "bbbbbbb" }];
    const buildOf = (l: { build: string }) => l.build;
    expect(roundToJudgeOn(logs, "ccccccc", buildOf), "judged on a round older than the change").toBeUndefined();
    // And the ordinary cases still work: after the prompting round, never at or
    // before it.
    expect(roundToJudgeOn(logs, "aaaaaaa", buildOf)).toEqual({ build: "bbbbbbb" });
    expect(roundToJudgeOn(logs, "bbbbbbb", buildOf)).toBeUndefined();
  });

  it("judges a prediction whose build nobody rounded on the rounds that came after it", () => {
    // The normal case, not the exotic one: a claim is written the moment a
    // change lands, and the next merge supersedes that commit before any round
    // runs. Matching the build exactly would answer `no round yet` forever — an
    // entry that can never be judged, which is the same as no entry at all.
    //
    // A round's build stamp carries its date, so a dated entry can be judged on
    // any round taken after that day.
    const logs = [{ build: "aaaaaaa · 2026-08-17 09:03Z" }, { build: "bbbbbbb · 2026-08-20 07:11Z" }];
    const buildOf = (l: { build: string }) => l.build.split(" ")[0];
    expect(roundToJudgeOn(logs, "ccccccc", buildOf, "2026-08-19")).toEqual({ build: "bbbbbbb · 2026-08-20 07:11Z" });
    // And a round taken BEFORE the claim was made is still no evidence about it.
    expect(roundToJudgeOn(logs, "ccccccc", buildOf, "2026-08-21")).toBeUndefined();
    // The build match wins where it exists — it is exact, and a date is only as
    // good as the stamp.
    expect(roundToJudgeOn(logs, "aaaaaaa", buildOf, "2026-01-01")).toEqual({ build: "bbbbbbb · 2026-08-20 07:11Z" });
  });

  it("judges a claim about what the TRACE says, and knows silence from absence", () => {
    // The ledger's four kinds were all about scenarios and probes, so a question
    // whose whole answer is "does this line appear" had to live in prose — which
    // is what this ledger exists to replace.
    const claim = {
      claim: {
        kind: "trace-line-present",
        message: "the namespace IS reachable — the fault is further in",
        insteadOf: "the namespace is unreachable too",
      },
    };
    const round = (...msgs: string[]) => ({ trace: { entries: msgs.map((m) => ({ message: m })) } });
    expect(judgePrediction(claim, round("the namespace IS reachable — the fault is further in")).verdict).toBe("held");
    expect(judgePrediction(claim, round("the namespace is unreachable too")).verdict).toBe("FAILED");
    // NEITHER LINE IS NOT A REFUTATION. A round that never reached the code did
    // not ask the question, and without `insteadOf` this kind could not tell
    // that apart from an answer — every round that skipped the path would read
    // as evidence against the claim.
    expect(judgePrediction(claim, round("something else entirely")).verdict).toBe("undetermined");
    expect(judgePrediction(claim, {}).verdict, "no trace at all is not evidence either").toBe("undetermined");
  });

  it("can claim a symptom is GONE, and knows that from a round that never ran", () => {
    // What a fix predicts is an absence, and this kind could only claim a
    // presence. `absent: true` inverts it; `insteadOf` — a line every round
    // carries — is what separates "the symptom is gone" from "nothing happened".
    const claim = {
      claim: {
        kind: "trace-line-present",
        message: "the namespace IS reachable — the fault is further in",
        insteadOf: "round starting",
        absent: true,
      },
    };
    const round = (...msgs: string[]) => ({ trace: { entries: msgs.map((m) => ({ message: m })) } });
    expect(judgePrediction(claim, round("round starting")).verdict, "a cured round").toBe("held");
    expect(
      judgePrediction(claim, round("round starting", "the namespace IS reachable — the fault is further in")).verdict,
    ).toBe("FAILED");
    // THE HALF THAT MATTERS: a round that never ran is silent about everything,
    // and silence must not read as a cure.
    expect(judgePrediction(claim, round("something else")).verdict, "silence read as a cure").toBe("undetermined");
  });

  it("judges a prediction on a round taken the SAME DAY it was staked", () => {
    // ROUND 088 IS THIS TEST. The #586 entry was staked on 2026-08-19 and round
    // 088 was taken at 13:58Z the same day — and the report still said `no round
    // yet`, because the stamp was read down to the DAY and compared with a
    // strict `>`, so `"2026-08-19" > "2026-08-19"` is false and the only round
    // that could settle the claim was the one round it refused.
    //
    // Same-day is the NORMAL case, not the edge: a claim is staked when the
    // change lands and the round is run within the hour.
    const logs = [{ build: "95170cf · 2026-08-17 09:03Z" }, { build: "3056f91 · 2026-08-19 13:58Z" }];
    const buildOf = (l: { build: string }) => l.build.split(" ")[0];
    expect(roundToJudgeOn(logs, "b9fef69", buildOf, "2026-08-19"), "refused the round it was waiting for").toEqual({
      build: "3056f91 · 2026-08-19 13:58Z",
    });
    // A day the archive has not reached is still no round, rather than the
    // newest one standing in for it.
    expect(roundToJudgeOn(logs, "b9fef69", buildOf, "2026-08-20")).toBeUndefined();
    // And a round from the day BEFORE is still not evidence.
    expect(roundToJudgeOn([logs[0]], "b9fef69", buildOf, "2026-08-19")).toBeUndefined();
  });

  it("does not judge a prediction on a sibling round of the very build it was staked on", () => {
    // The cycle archives two rounds at 16:9 and one at 4:3 on ONE build, and
    // "run the same build twice" is the discipline this loop is built on. Taking
    // the FIRST index of that build left its own siblings in the slice, so the
    // newest of them could be handed back as the round that judges a change made
    // after all three were taken — the judge ruling on its own control.
    // THE SIBLINGS HAVE TO BE THE NEWEST ROUNDS or this test proves nothing —
    // the first draft put a later build after them, and then BOTH the first-index
    // and last-index readings return that later build and the test passes against
    // the bug. Mutation caught it; reading did not. With the pair at the end of
    // the archive there is nothing legitimate to return, so the first-index
    // reading hands back a sibling and the last-index reading correctly has
    // nothing.
    const logs = [
      { build: "aaaaaaa · 2026-08-18 01:00Z" },
      { build: "bbbbbbb · 2026-08-18 02:00Z" },
      { build: "bbbbbbb · 2026-08-18 03:00Z" },
      { build: "bbbbbbb · 2026-08-18 04:00Z" },
    ];
    const buildOf = (l: { build: string }) => l.build.split(" ")[0];
    expect(roundToJudgeOn(logs, "bbbbbbb", buildOf), "judged on a sibling of its own build").toBeUndefined();
    // And a build with rounds after it still judges on the newest of THOSE.
    const later = [...logs, { build: "ccccccc · 2026-08-19 05:00Z" }];
    expect(roundToJudgeOn(later, "bbbbbbb", buildOf)).toEqual({ build: "ccccccc · 2026-08-19 05:00Z" });
  });

  it("calls a scenario the host never ran undetermined, not FAILED", () => {
    // "A skip is not a flip" and "a miss is not a failure" are this repo's own
    // doctrine, and `probe-answers` and `probe-detail-matches` both honour it —
    // `scenario-passes` did not, and recorded the ledger's strongest verdict on
    // a round where the host declined to be asked.
    //
    // Round 088 made it live: `insert onto a slide that already has content`
    // came back `skipped: true, ok: false` because PowerPoint stopped answering
    // while drawing shapes 1-10 of 16.
    const claim = { claim: { kind: "scenario-passes", names: ["insert onto a slide that already has content"] } };
    const skipped = judgePrediction(claim, {
      selftest: [{ name: "insert onto a slide that already has content", ok: false, skipped: true }],
    });
    expect(skipped.verdict, "a host that never ran the scenario refuted the claim").toBe("undetermined");
    expect(skipped.why).toMatch(/never ran/);
    // A scenario absent from the sheet is the same kind of silence.
    expect(judgePrediction(claim, { selftest: [] }).verdict).toBe("undetermined");
    // AND THE GUARD STILL FIRES. A scenario the host DID run and that failed is
    // still FAILED — the failure mode of this fix is a judge that never refutes
    // anything, which would be worse than the bug.
    expect(
      judgePrediction(claim, { selftest: [{ name: "insert onto a slide that already has content", ok: false }] })
        .verdict,
    ).toBe("FAILED");
    expect(
      judgePrediction(claim, { selftest: [{ name: "insert onto a slide that already has content", ok: true }] })
        .verdict,
    ).toBe("held");
  });

  it("keeps the ledger honest: every entry is judgeable and carries its reasoning", () => {
    const ledger = JSON.parse(readFileSync("rounds/predictions.json", "utf8"));
    expect(ledger.length, "the ledger is empty").toBeGreaterThan(0);
    const kinds = new Set([
      "probe-answers",
      "probe-starves",
      "scenario-passes",
      "probe-detail-matches",
      "trace-line-present",
    ]);
    for (const p of ledger) {
      expect(p.id, "a prediction with no id").toBeTruthy();
      expect(p.because, `${p.id} states no reasoning — the prose is the half worth keeping`).toBeTruthy();
      expect(p.afterBuild, `${p.id} does not say which build it was made on`).toMatch(/^[0-9a-f]{7}$/);
      expect(kinds.has(p.claim?.kind), `${p.id} has an unjudgeable claim kind`).toBe(true);
      // `undetermined` is a real outcome and the judge has always been able to
      // reach it — `judgePrediction` returns it for a question that was never
      // put. The stored ledger could not say it, so a prediction the round
      // could not settle had to be filed as open forever or, worse, as held.
      // `slide-identity-explains-the-undefined-tag` was the case that showed
      // it: seven rounds of `slide stable (?)`, which is three unreadable ids
      // agreeing with each other, and it read as a pass.
      expect(["open", "held", "failed", "undetermined"], `${p.id} has an odd outcome`).toContain(p.outcome);
      if (p.outcome !== "open")
        expect(p.judgedOn, `${p.id} is settled but does not say which round settled it`).toBeTruthy();
    }
  });
});

describe("scenario verdicts across rounds", () => {
  const round = (verdicts: Record<string, boolean | "skip">) => ({
    selftest: Object.entries(verdicts).map(([name, v]) => ({ name, ok: v === true, skipped: v === "skip" })),
  });

  it("flags a scenario that has said both pass and fail about unchanged code", () => {
    const h = scenarioHistory([round({ a: true }), round({ a: false }), round({ a: true })]);
    expect(h[0].flips).toBe(true);
    expect(h[0].verdicts).toEqual(["pass", "FAIL", "pass"]);
  });

  it("does NOT call a skip a flip", () => {
    // `the chart is actually visible` reads pass pass pass skip. It has never
    // disagreed with itself; the host stopped answering. Calling that unstable
    // would send someone hunting a bug in a scenario that never failed.
    const h = scenarioHistory([round({ a: true }), round({ a: true }), round({ a: "skip" })]);
    expect(h[0].flips, "a skip was counted as a disagreement").toBe(false);
    expect(h[0].skips).toBe(1);
  });

  it("leaves a consistently failing scenario alone — that is a finding, not noise", () => {
    const h = scenarioHistory([round({ a: false }), round({ a: false })]);
    expect(h[0].flips).toBe(false);
  });
});
