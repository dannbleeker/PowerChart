import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types. One directive per import, one
// import per line: Prettier merges a multi-name import onto its own lines and
// the directive then lands on the wrong one.
import { roundFromCrash } from "../scripts/salvage-crashed.mjs";
// @ts-expect-error — as above.
import { slideSizeFrom } from "../scripts/salvage-crashed.mjs";
// @ts-expect-error — as above.
import { alreadySalvaged } from "../scripts/salvage-crashed.mjs";
// @ts-expect-error — as above.
import { trustedSlideSize } from "../scripts/salvage-crashed.mjs";
// @ts-expect-error — as above.
import { verdictsFrom } from "../scripts/salvage-crashed.mjs";
// @ts-expect-error — as above.
import { expectedByBuild } from "../scripts/salvage-crashed.mjs";

/**
 * A ROUND THAT RAN IS EVIDENCE EVEN WHEN NOTHING FILED IT.
 *
 * The host dies in `collectDeckEvidence`, which runs after every verdict is in —
 * 9 builds against 4 on the per-phase traces, in a band of 441-572s. Until
 * 2026-08-29 the pane assembled its run log AFTER that scan, so a crash there
 * turned a complete round into a crash record and nothing else. 33 of 49 records
 * hold all their verdicts; 22 of those can also say, trustworthily, what slide
 * size they ran at.
 *
 * The whole risk of this tool is filing something that is NOT what it claims, so
 * most of these tests are about what it REFUSES.
 */
describe("salvaging a round from the crash record that kept it", () => {
  const verdict = (name: string, ok = true) => ({ key: `selftest:${name}`, value: { name, ok, detail: "…" } });
  const NAMES = ["one", "two", "three"];
  const crash = (over: Record<string, unknown> = {}) => ({
    build: "abc1234 · 2026-08-29 12:55Z",
    host: "PowerPoint · OfficeOnline · 0.0.0.0",
    steps: ["  10.0s  host  slide size read  width=960 height=540 source=pageSetup ms=350"],
    findings: [
      { key: "hostAnswers", value: { answers: [{ id: "q", answer: "yes" }] } },
      ...NAMES.map((n) => verdict(n)),
    ],
    ...over,
  });

  it("files a round when the verdicts cover what that build's archived round ran", () => {
    const out = roundFromCrash(crash(), NAMES, "2026-08-29T16-49-22-crashed-run.json");
    expect(out.ok, out.why).toBe(true);
    expect(out.round.selftest.map((s: { name: string }) => s.name)).toEqual(NAMES);
    expect(out.round.slideSize).toEqual({ width: 960, height: 540, source: "pageSetup" });
    expect(out.round.hostAnswers.answers).toHaveLength(1);
    expect(out.round.salvagedFrom, "a salvaged round must say so on its face").toBe(
      "2026-08-29T16-49-22-crashed-run.json",
    );
  });

  it("invents neither a deck nor a trace", () => {
    // `crashes/README.md` is right that a partial record must never be pooled,
    // and an INVENTED inventory is exactly the fabricated evidence it warns
    // about — everything downstream reads that field as measured. The crash
    // keeps `steps` as formatted strings; a round's trace is structured entries.
    // Reshaping prose into structure would be the same lie in a different shape.
    const out = roundFromCrash(crash(), NAMES, "x.json");
    expect(out.round.deck, "invented deck evidence the scan never returned").toBeUndefined();
    expect(out.round.trace, "reshaped the crash's prose steps into a structured trace").toBeUndefined();
  });

  it("files a round that never reached every verdict, and records how far it got", () => {
    /**
     * THIS USED TO BE A REFUSAL, and it was right when it was written. What it
     * cost: between this bar and the slide-size one, 47 crash records holding
     * real verdicts from a real host were being discarded — while this script's
     * own header says what such rounds are good for, "any question that does not
     * care about order: what a build's verdicts were, how often a scenario has
     * ever failed."
     *
     * A partial set is evidence about the scenarios that RAN. What made it
     * dangerous was pooling it against a global denominator: a round that
     * reached 2 of 3 has not PASSED the third, and a naive rate reads that
     * silence as a success. So the reach is recorded, and anything pooling
     * these has to condition on it.
     *
     * The names are still the test rather than the count — a missing scenario
     * and a renamed one are different things, and only names tell them apart.
     */
    const short = crash({ findings: [verdict("one"), verdict("three")] });
    const out = roundFromCrash(short, NAMES, "x.json");
    expect(out.ok, "a partial round is still evidence about what it did run").toBe(true);
    expect(out.round.salvagedPartial, "filed a partial round without saying it was partial").toBeTruthy();
    expect(out.round.salvagedPartial.reached).toBe(2);
    expect(out.round.salvagedPartial.of).toBe(NAMES.length);
    expect(out.round.salvagedPartial.missing, "did not name what never ran").toContain("two");
    expect(out.round.selftest.map((s: { name: string }) => s.name)).toEqual(["one", "three"]);
  });

  it("marks a complete round as complete, so `partial` means something", () => {
    // The other half: a flag every round carried would say nothing.
    const out = roundFromCrash(crash(), NAMES, "x.json");
    expect(out.round.salvagedPartial, "a complete round was labelled partial").toBeUndefined();
  });

  it("refuses a slide size that came from the rung the host falls to after giving up", () => {
    /**
     * THE ONE THAT WOULD HAVE DONE REAL DAMAGE. `exportedSlide` is the last
     * rung, reached only when the host stopped answering, and it reads a saved
     * file PowerPoint may not have updated. `archive`'s own docstring records
     * rounds 115 and 116 filing 720x540 while running on a 960x540 deck exactly
     * that way — caught only because `driverSlideSize` recorded the driver's
     * independent reading beside the pane's.
     *
     * A salvaged round has no driver reading: nobody was holding the other end,
     * because the driver filed a crash instead of a round. 18 of 46 records rest
     * on this rung, and this repo's rule is that filing a round under the wrong
     * profile is worse than not running it.
     */
    const fallback = crash({
      steps: ["  10.0s  host  slide size read  width=720 height=540 source=exportedSlide ms=1831"],
    });
    const out = roundFromCrash(fallback, NAMES, "x.json");
    /**
     * THE PROFILE IS REFUSED; THE ROUND IS NOT. Changed 2026-09-03, and the
     * distinction is the whole point.
     *
     * `slideSize` stays ABSENT — a present field is read as a measurement by
     * everything downstream, and no flag beside it survives being globbed. So
     * nothing can file this round under a profile, which is the rule the old
     * refusal was protecting. The reading is kept under a name that says what
     * it is, so nothing is thrown away and nothing is claimed.
     *
     * 27 of the 69 salvageable records rest on this rung, and every one is a
     * complete answer to every question that is not about profile.
     */
    expect(out.ok, "threw away a round because one field could not be trusted").toBe(true);
    expect(out.round.slideSize, "filed a profile read off the rung that mis-filed rounds 115 and 116").toBeUndefined();
    expect(out.round.slideSizeUnverified, "discarded the reading instead of recording it").toEqual({
      width: 720,
      height: 540,
      source: "exportedSlide",
    });
    // `trustedSlideSize` itself is unchanged: null still means "do not file a
    // profile from this". Only what the caller does with null moved.
    expect(trustedSlideSize({ width: 960, height: 540, source: "pageSetup" })).toBeTruthy();
    expect(trustedSlideSize({ width: 960, height: 540, source: "exportedSlide" })).toBeNull();
  });

  it("reads both spellings of the slide-size trace line", () => {
    /**
     * The pane traces this two ways — `slide size read` when a rung answers and
     * `slide size` for the value it settled on. The matcher named only the
     * first, so every record carrying only the second was refused as "nothing
     * in it says what slide size it ran at" while saying it plainly. Across the
     * corpus that is 64 lines against 60.
     *
     * Same family as CHAR(11) and CHAR(10) in this project's notes: a matcher
     * that names one spelling covers exactly that spelling.
     */
    const settled = crash({
      steps: ["  10.0s  host  slide size  width=960 height=540 source=pageSetup"],
    });
    const out = roundFromCrash(settled, NAMES, "x.json");
    expect(out.ok, "the settled spelling still reads as no slide size at all").toBe(true);
    expect(out.round.slideSize).toEqual({ width: 960, height: 540, source: "pageSetup" });
  });

  it("refuses when nothing in the record says what size it ran at", () => {
    expect(roundFromCrash(crash({ steps: [] }), NAMES, "x.json").ok).toBe(false);
    expect(roundFromCrash(crash({ steps: [] }), NAMES, "x.json").why).toMatch(/slide size/);
  });

  it("refuses when no archived round of that build says what complete means", () => {
    // Better than a hard-coded 14: that would refuse the genuinely complete
    // twelve-scenario rounds earlier in this archive and wave through a partial
    // one the day a fifteenth scenario lands.
    const out = roundFromCrash(crash(), [], "x.json");
    expect(out.ok).toBe(false);
    expect(out.why).toMatch(/no archived round/);
  });

  it("prefers the header the pane now banks, and otherwise takes the LAST reading", () => {
    const withHead = crash({
      findings: [{ key: "runLogHead", value: { slideSize: { width: 720, height: 540, source: "pageSetup" } } }],
    });
    expect(slideSizeFrom(withHead)).toEqual({ width: 720, height: 540, source: "pageSetup" });
    // The pane reads the size at the END of a round, so the last reading is the
    // one a filed round would have carried.
    const twice = crash({
      steps: [
        "  10.0s  host  slide size read  width=960 height=540 source=pageSetup ms=350",
        " 200.0s  host  slide size read  width=720 height=540 source=pageSetup ms=350",
      ],
    });
    expect(slideSizeFrom(twice)!.width).toBe(720);
  });

  it("never takes its idea of a complete round from another salvage", () => {
    // The bar is the LONGEST set seen, so a short salvage could never lower it —
    // the reachable danger is the other direction, a salvage whose verdict set
    // is longer than any round actually filed for that build, which would then
    // define "complete" for every record after it. Reference data must come from
    // rounds the driver filed, never from this tool's own output: otherwise the
    // second salvage is judged against the first and nothing outside the loop
    // ever checks it.
    const files = ["001-a.json", "002-a.json"];
    const read = (p: string) =>
      p.endsWith("001-a.json")
        ? JSON.stringify({ build: "a · x", selftest: [{ name: "one" }] })
        : JSON.stringify({ build: "a · x", selftest: [{ name: "one" }, { name: "two" }], salvagedFrom: "c.json" });
    const got = expectedByBuild("rounds", read as never, (() => files) as never);
    expect(got.get("a"), "a salvaged round was allowed to define what complete means").toEqual(["one"]);
  });

  it("reads the verdicts and nothing else out of the findings", () => {
    const mixed = crash({ findings: [{ key: "hostAnswers", value: { answers: [] } }, verdict("one")] });
    expect(verdictsFrom(mixed).map((v: { name: string }) => v.name)).toEqual(["one"]);
  });
});

describe("running the salvage twice", () => {
  /**
   * IT RAN ONCE BY HAND ON 2026-08-29 AND NOTHING CALLED IT AGAIN. 47 more
   * crash records accumulated, and a second run would have re-filed the
   * original 22 under fresh numbers beside the copies already there —
   * `archive`'s twin check catches the byte-identical ones and nothing catches
   * the rest. An archive that double-counts its own evidence is worse than one
   * that is merely incomplete.
   */
  const files = (names: string[]) => () => names;
  const holding = (map: Record<string, string>) => (p: string) => {
    const key = p.split("/").pop()!;
    if (!(key in map)) throw new Error(`no such file: ${p}`);
    return map[key];
  };

  it("knows which crash records it has already filed", () => {
    const seen = alreadySalvaged(
      "out",
      () => true,
      files(["001-aaa.json", "002-bbb.json"]),
      holding({
        "001-aaa.json": JSON.stringify({ salvagedFrom: "2026-08-26T23-04-11-crashed-run.json" }),
        "002-bbb.json": JSON.stringify({ salvagedFrom: "2026-08-27T05-48-38-crashed-run.json" }),
      }),
    );
    expect(seen.has("2026-08-26T23-04-11-crashed-run.json")).toBe(true);
    expect(seen.has("2026-08-27T05-48-38-crashed-run.json")).toBe(true);
    expect(seen.has("2026-09-01T00-00-00-crashed-run.json"), "claimed a record it has never seen").toBe(false);
  });

  it("treats a missing target directory as nothing salvaged, not as an error", () => {
    // The first run, and the only state this keeps is the directory itself.
    expect(alreadySalvaged("out", () => false, files([]), holding({})).size).toBe(0);
  });

  it("reads past a salvaged round that will not parse", () => {
    /**
     * The same rule the round loader follows: one bad file must not refuse the
     * other sixty-eight. The cost of skipping it is that its source record may
     * be filed a second time; the cost of throwing would be salvaging nothing
     * at all, for ever, until someone noticed.
     */
    const seen = alreadySalvaged(
      "out",
      () => true,
      files(["001-aaa.json", "002-broken.json", "003-ccc.json"]),
      holding({
        "001-aaa.json": JSON.stringify({ salvagedFrom: "a.json" }),
        "002-broken.json": "{ this is not json",
        "003-ccc.json": JSON.stringify({ salvagedFrom: "c.json" }),
      }),
    );
    expect(seen.has("a.json")).toBe(true);
    expect(seen.has("c.json"), "one unparseable file stopped it reading the rest").toBe(true);
    expect(seen.size).toBe(2);
  });

  it("ignores a salvaged round that names no source", () => {
    // Nothing to key on. Counting it would be inventing a match.
    const seen = alreadySalvaged(
      "out",
      () => true,
      files(["001-aaa.json"]),
      holding({ "001-aaa.json": JSON.stringify({ build: "x" }) }),
    );
    expect(seen.size).toBe(0);
  });
});
