import { describe, expect, it } from "vitest";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
// @ts-expect-error — a plain .mjs tool with no types. The tables live THERE so
// the diff tool and this gate cannot drift apart.
import { FAKE_BASELINE, KNOWN_DIVERGENCES, PENDING_QUESTIONS, answersOf, diffAnswers } from "../scripts/host-diff.mjs";

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
    expect(realSheet.kind).toBe("powerchart-host-answers");
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
    // to-do wearing a passing test's clothes. Two shapes are allowed, and both
    // say something a reader can act on: a different host on purpose, or an
    // answer known to be about the probe and re-asked.
    for (const [id, why] of Object.entries(KNOWN_DIVERGENCES) as [string, string][]) {
      expect(why.length, `${id} is declared with no reason`).toBeGreaterThan(40);
      expect(
        /WITHDRAWN|models a DIFFERENT host|models the host|happy path/i.test(why),
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
});
