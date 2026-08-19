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

  it("keeps the ledger honest: every entry is judgeable and carries its reasoning", () => {
    const ledger = JSON.parse(readFileSync("rounds/predictions.json", "utf8"));
    expect(ledger.length, "the ledger is empty").toBeGreaterThan(0);
    const kinds = new Set(["probe-answers", "probe-starves", "scenario-passes", "probe-detail-matches"]);
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
