import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
// @ts-expect-error — a plain .mjs tool with no types. The tables live THERE so
// the diff tool and this gate cannot drift apart.
// prettier-ignore
import { FAKE_BASELINE, KNOWN_DIVERGENCES, PENDING_QUESTIONS, answersOf, diffAnswers, RENAMED_ANSWERS, pendingAlreadyAnswered } from "../scripts/host-diff.mjs";
import { isHostAnswersKind } from "../src/render/host-probe";

/**
 * The fake, checked against a real PowerPoint — in CI, on every commit.
 *
 * `test/host-probe.test.ts` freezes what the fake CLAIMS. This freezes what a
 * real host actually SAID, and puts the two in the same room. Until now that
 * comparison was a command someone had to remember to run, against a sheet that
 * lived in a chat upload, which is the same as not having it: every host bug
 * this project has fixed was found by a human running PowerPoint and noticing.
 *
 * The gate is not "the two must agree" — they do not, and several of the
 * disagreements are deliberate. It is "every disagreement is DECLARED", in
 * `KNOWN_DIVERGENCES`, with the reason. A new one fails here, in seconds,
 * instead of surviving until someone next has a real host in front of them.
 *
 * When a new real sheet arrives: replace the fixture, run this, and deal with
 * what goes red. That is the whole ritual.
 */

const REAL_SHEET_PATH = fileURLToPath(new URL("./fixtures/host-answers-web.json", import.meta.url));
const realSheet = JSON.parse(readFileSync(REAL_SHEET_PATH, "utf8"));
const real = answersOf(realSheet);

describe("the fake, against the real host it stands for", () => {
  it("carries a sheet that says which host and which build answered", () => {
    // A sheet with no provenance cannot be read a week later: the same verdict
    // means different things on 1.4 and 1.10, and on a build before a probe was
    // rewritten it may mean nothing at all.
    // RECOGNISED, not equal to today's spelling. This fixture is a real sheet
    // captured from the live host before the 2026-08-27 rename, so it carries
    // `powerchart-host-answers` — and it should keep carrying it. Rewriting a
    // captured artefact to match a rename falsifies the evidence, and pinning
    // the new name here would force exactly that. The archive holds 257 rounds
    // in the same position.
    expect(isHostAnswersKind(realSheet.kind), `unrecognised sheet kind: ${realSheet.kind}`).toBe(true);
    expect(realSheet.source, "the fixture does not say which host answered").toBeTruthy();
    expect(realSheet.build, "the fixture does not say which build asked").toBeTruthy();
    expect(realSheet.requirementSets?.length, "the fixture does not say what the host supports").toBeGreaterThan(0);
  });

  it("diverges only where the divergence is declared", () => {
    const { differ } = diffAnswers(real, FAKE_BASELINE);
    const undeclared = differ.filter((d: { id: string }) => !(d.id in KNOWN_DIVERGENCES));
    expect(
      undeclared.map((d: { id: string; real: string; fake: string }) => `${d.id}: host=${d.real} fake=${d.fake}`),
      "the fake now claims something a real PowerPoint contradicts, and nothing says why",
    ).toEqual([]);
  });

  it("declares nothing that has stopped diverging", () => {
    // The other direction, and the one that rots quietly. An entry left behind
    // after the fake was fixed reads as a known problem forever, and the next
    // person to look decides the gate is noise.
    //
    // A question the run could not SET UP is not evidence that the divergence
    // went away. `diffAnswers` sorts `no-scratch-slide` and `no-scratch-shape`
    // into `notAsked` precisely because they say nothing about the host, and
    // reading their absence from `differ` as agreement would have this gate
    // demand the deletion of a declaration nobody has contradicted. The
    // 2026-08-08 sheet did exactly that to `tags-add-same-key-twice` and
    // `tag-on-group-survives`, on the strength of a scratch slide that would
    // not resolve.
    const { differ, notAsked } = diffAnswers(real, FAKE_BASELINE);
    const unresolved = new Set([
      ...differ.map((d: { id: string }) => d.id),
      ...notAsked.map((n: { id: string }) => n.id),
    ]);
    const stale = Object.keys(KNOWN_DIVERGENCES).filter((id) => !unresolved.has(id));
    expect(stale, "declared as divergent, but the fake and the host now agree").toEqual([]);
  });

  it("gives a reason for each one, not a placeholder", () => {
    // "We have not looked into it" is not a reason; an entry like that is a
    // to-do wearing a passing test's clothes. Three shapes are allowed, and all
    // three say something a reader can act on: a different host on purpose, an
    // answer known to be about the probe and re-asked, or NO CALLER — the fake
    // is knowingly the optimistic side and nothing in the repo depends on it.
    //
    // The third was added for `binding-names-shape-later`, where the host
    // rejects the batch that carries a binding and the fake says it works.
    // Teaching the fake to poison a batch would be fiction with nothing to
    // protect, since no code here makes a binding — but the divergence still
    // has to be declared, because the fake being the optimistic one is the
    // direction that misleads, and the entry is what a future reader reaching
    // for bindings finds first.
    for (const [id, why] of Object.entries(KNOWN_DIVERGENCES) as [string, string][]) {
      expect(why.length, `${id} is declared with no reason`).toBeGreaterThan(40);
      expect(
        /WITHDRAWN|models a DIFFERENT host|models the host|happy path|no caller/i.test(why),
        `${id}'s reason does not say WHY the divergence is allowed to stand`,
      ).toBe(true);
    }
  });

  it("has an answer from the real host for every question, or says why not", () => {
    // A question the sheet has no answer for is a hole in the comparison, and
    // this gate is worth exactly as much as its coverage. A probe added without
    // a re-run leaves one — allowed, but only in writing, so nobody can shrink
    // what this covers without saying so.
    const { onlyFake } = diffAnswers(real, FAKE_BASELINE);
    const undeclared = onlyFake.filter((id: string) => !(id in PENDING_QUESTIONS));
    expect(
      undeclared,
      `the committed sheet predates these questions — re-run the probe (fixture build: ${realSheet.build})`,
    ).toEqual([]);
  });

  it("declares nothing as pending that the sheet already answers", () => {
    // The list has to shrink on its own when a newer sheet lands, or it becomes
    // a place where questions go to be forgotten.
    //
    // But a sheet that CARRIES a question and says `no-scratch-slide` has not
    // answered it — the run never got as far as putting it. Counting that as
    // answered would retire the question from the pending list on the strength
    // of a setup failure, which is the forgetting this list exists to prevent,
    // arriving by the door marked "shrinks on its own".
    const { onlyFake, notAsked } = diffAnswers(real, FAKE_BASELINE);
    const unknown = new Set([...onlyFake, ...notAsked.map((n: { id: string }) => n.id)]);
    const answered = Object.keys(PENDING_QUESTIONS).filter((id) => !unknown.has(id));
    expect(answered, "declared as unanswered, but the committed sheet answers it").toEqual([]);
  });

  it("reads a pre-rename capture as the same observation, without rewriting it", async () => {
    /**
     * `all` was `which-end-a-short-read-drops`'s way of saying its question did
     * not arise. It is `not-a-short-read` now, because as `all` it ranked as a
     * real answer, locked the row, and MATCHED the fake's `all` — so `host-diff`
     * recorded 87 rounds of the host agreeing with the fake about which end a
     * short read drops, on a question neither had ever answered.
     *
     * The committed sheet still says `all`, and must. The rule is one screen up
     * in this file, about the 2026-08-27 `kind` rename: "Rewriting a captured
     * artefact to match a rename falsifies the evidence." So the old word is
     * RECOGNISED on the way in and the artefact is left alone — otherwise a
     * rename manufactures a host divergence that never happened, and
     * `KNOWN_DIVERGENCES` gains an entry calling a vocabulary change a
     * behavioural one.
     */
    const read = answersOf;
    expect(RENAMED_ANSWERS["which-end-a-short-read-drops"].all).toBe("not-a-short-read");
    const seen = read({
      kind: realSheet.kind,
      answers: [
        { id: "which-end-a-short-read-drops", answer: "all" },
        { id: "scratch-slides-returned", answer: "all" },
      ],
    });
    expect(seen["which-end-a-short-read-drops"], "the pre-rename word was not recognised").toBe("not-a-short-read");
    // THE SAME WORD, AND IT MUST NOT MOVE HERE. `scratch-slides-returned`
    // answers `all` and means it — every scratch slide came back. The table is
    // keyed by question for exactly this reason.
    expect(seen["scratch-slides-returned"], "translated another probe's real answer").toBe("all");
  });

  it("catches a pending question the ARCHIVE has already answered", async () => {
    /**
     * The register says to delete an entry once the host answers, and calls an
     * id left behind "the fixture going stale in writing". The gate above cannot
     * see that: it compares against the committed sheet, where a pending id is
     * absent precisely BECAUSE the fixture predates it — green whether the
     * question is unanswered or answered fifty times.
     *
     * Measured 2026-09-01: `named-preset-resolves` had answered `draws` in 11 of
     * 11 rounds, 33 of 33 samples, across five builds, while every gate stayed
     * green. The archive can see what the fixture cannot.
     */
    const round = (id: string, answer: string) => ({
      hostAnswers: { kind: realSheet.kind, answers: [{ id, answer }] },
    });
    const pending = { "a-real-question": "why it is pending" };

    const found = pendingAlreadyAnswered([round("a-real-question", "draws")], pending);
    expect(
      found.map((f: { id: string }) => f.id),
      "the archive answered it and nothing said so",
    ).toEqual(["a-real-question"]);
    expect(found[0].answers).toEqual([{ answer: "draws", n: 1 }]);

    // A WEAK WORD IS NOT AN ANSWER. Retiring a question because one round said
    // `unreadable` or `silent` would be the forgetting this register exists to
    // prevent, arriving through the door marked "the archive says so". It is
    // also exactly the state `rotation-keeps-the-unrotated-box` sits in until a
    // round runs on a build that stopped letting `unreadable` lock its row.
    for (const quiet of ["unreadable", "silent", "no-scratch-slide", "other"])
      expect(
        pendingAlreadyAnswered([round("a-real-question", quiet)], pending),
        `"${quiet}" retired a question`,
      ).toEqual([]);
  });
});
