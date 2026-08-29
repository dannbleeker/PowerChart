import { describe, expect, it } from "vitest";
// @ts-expect-error — plain .mjs tool, no types. One directive per import, one
// import per line: Prettier merges a multi-name import onto its own lines and
// the directive then lands on the wrong one.
import { roundFromCrash } from "../scripts/salvage-crashed.mjs";
// @ts-expect-error — as above.
import { slideSizeFrom } from "../scripts/salvage-crashed.mjs";
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

  it("refuses a round that never reached every verdict, and names which", () => {
    // The count is not the test — a missing scenario and a renamed one are
    // different, and only names can tell them apart.
    const short = crash({ findings: [verdict("one"), verdict("three")] });
    const out = roundFromCrash(short, NAMES, "x.json");
    expect(out.ok).toBe(false);
    expect(out.why).toMatch(/two/);
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
    expect(out.ok, "filed a profile read off the rung that has already mis-filed two rounds").toBe(false);
    expect(out.why).toMatch(/exportedSlide/);
    expect(trustedSlideSize({ width: 960, height: 540, source: "pageSetup" })).toBeTruthy();
    expect(trustedSlideSize({ width: 960, height: 540, source: "exportedSlide" })).toBeNull();
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
