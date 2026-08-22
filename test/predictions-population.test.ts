import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types.
import { judgePrediction } from "../scripts/triage.mjs";
// Its own line — see the grouped-import trap documented in `triage.test.ts`.
// @ts-expect-error — as above.
import { judgeAcross } from "../scripts/triage.mjs";
// @ts-expect-error — as above.
import { roundsToJudgeOn } from "../scripts/triage.mjs";

/**
 * WHAT #520 COST: a prediction staked on 2026-08-18 sat OPEN for 130 rounds
 * while the archive answered it twenty times. Two independent defects, both the
 * house shape — an unmeasured thing reported as a negative measurement.
 */
describe("a conditional probe's question is not always put", () => {
  const sheet = (answers: { id: string; answer: string; detail?: string }[]) => ({ hostAnswers: { answers } });
  const claim = { claim: { kind: "probe-detail-matches", id: "grp", pattern: "tags-gone|tags was undefined" } };

  it("reads a precondition that did not occur as undetermined, not as a refutation", () => {
    // `does-a-failed-group-poison-the-tag` only has a question when the host
    // REFUSES the group. In 94 of 114 archived rounds it grouped happily and
    // the probe said so in as many words — and the judge scored that FAILED.
    const r = judgePrediction(
      claim,
      sheet([
        {
          id: "grp",
          answer: "no-refusal",
          detail: "the host grouped through a slide handle two syncs old, so the refusal was never provoked",
        },
      ]),
    );
    expect(r.verdict).toBe("undetermined");
    expect(r.why).toContain("no-refusal");
  });

  it("matches the claim against the probe's answer KEY, not only its prose", () => {
    // #520 was staked on `InvalidParam|5010|GeneralException` — error codes this
    // probe has never emitted — so even its twenty firings read as FAILED. An
    // answer key is an enum the probe picks; a detail is a sentence someone
    // rewords.
    const r = judgePrediction(
      { claim: { kind: "probe-detail-matches", id: "grp", pattern: "tags-gone" } },
      sheet([{ id: "grp", answer: "tags-gone", detail: "a sentence a later commit reworded" }]),
    );
    expect(r.verdict).toBe("held");
  });

  it("still FAILS on a round that measured the thing and got the other answer", () => {
    const r = judgePrediction(claim, sheet([{ id: "grp", answer: "tags-kept", detail: "the tag survived" }]));
    expect(r.verdict).toBe("FAILED");
  });
});

describe("judging across the population instead of on whichever round was newest", () => {
  const buildOf = (l: { build: string }) => l.build;
  const round = (build: string, answer: string, detail = "") => ({
    build,
    hostAnswers: { answers: [{ id: "grp", answer, detail }] },
  });
  const claim = { claim: { kind: "probe-detail-matches", id: "grp", pattern: "tags-gone" } };

  it("counts every round that measured it and ignores the ones that could not", () => {
    const t = judgeAcross(
      claim,
      [round("a", "tags-gone"), round("b", "no-refusal"), round("c", "no-refusal"), round("d", "tags-gone")],
      buildOf,
    );
    expect(t.verdict).toBe("held");
    expect([t.held, t.failed, t.undecided]).toEqual([2, 0, 2]);
  });

  it("does not let the newest round speak for a probe that rarely fires", () => {
    // The whole defect in one assertion: the newest round says nothing, and the
    // single-round judge takes its verdict from exactly that round.
    const rounds = [round("a", "tags-gone"), round("b", "no-refusal")];
    expect(judgePrediction(claim, rounds[rounds.length - 1]).verdict).toBe("undetermined");
    expect(judgeAcross(claim, rounds, buildOf).verdict).toBe("held");
  });

  it("says BOTH when the decided rounds disagree, rather than picking one", () => {
    const t = judgeAcross(claim, [round("a", "tags-gone"), round("b", "tags-kept")], buildOf);
    expect(t.verdict).toBe("BOTH");
    expect([t.held, t.failed]).toEqual([1, 1]);
  });

  it("stays undetermined when no eligible round measured it", () => {
    const t = judgeAcross(claim, [round("a", "no-refusal"), round("b", "no-refusal")], buildOf);
    expect(t.verdict).toBe("undetermined");
    expect(t.decided).toBe(0);
  });

  it("quotes the LATEST reading, and prefers a refutation when both exist", () => {
    const t = judgeAcross(claim, [round("a", "tags-gone"), round("b", "tags-kept", "the tag survived")], buildOf);
    expect(t.last.build).toBe("b");
    expect(t.last.why).toContain("tags-kept");
  });

  it("hands back every round after the staking build, newest last", () => {
    const logs = [{ build: "a" }, { build: "b" }, { build: "c" }];
    expect(roundsToJudgeOn(logs, "a", buildOf).map(buildOf)).toEqual(["b", "c"]);
    expect(roundsToJudgeOn(logs, "c", buildOf)).toEqual([]);
  });
});
