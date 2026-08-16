import { describe, expect, it } from "vitest";
import { spawnSync } from "child_process";
import { mkdtempSync, writeFileSync, rmSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
// @ts-expect-error — plain .mjs tools, no types; both are deliberately
// independent of src/ so they cannot inherit a bug from the code they audit.
import { readDeckBytes } from "../scripts/verify-deck.mjs";
// @ts-expect-error — as above.
import { triage, runsIn, selfTestIn, knownBug, deckEvidence, poolRasteriseArms } from "../scripts/triage.mjs";
// Its own line, and this file has now predicted its own bug three times: adding
// a name to the grouped import above pushes it over the print width, prettier
// reflows it across lines, and `@ts-expect-error` covers only the NEXT line — so
// the directive stops reaching the `from` clause. Suite green, `tsc` red.
// @ts-expect-error — as above.
import { poolTagFaults } from "../scripts/triage.mjs";
// Its own line, same reason as every other single import in this file.
// @ts-expect-error — as above.
import { poolGroupVsTag } from "../scripts/triage.mjs";
// A NAMESPACE import, destructured below, and the fourth time this trap has been
// paid for. `@ts-expect-error` covers exactly one LINE — the one carrying `from`
// — so a named list long enough for prettier to wrap moves `from` four lines
// down, uncovers the import, and turns the directive itself into an "unused
// directive" error. Adding `scenarioRegressions` to the list did precisely that.
// A namespace import is one line whatever is destructured off it.
// @ts-expect-error — as above.
import * as pools from "../scripts/triage.mjs";
const {
  poolFreshVsEstablished,
  poolStarvedQuestions,
  poolBatchSpanVsGroup,
  scenarioRegressions,
  poolOriginTagLosses,
  roundProfile,
  profileDivergence,
  traceNovelty,
  traceSignature,
} = pools;
// Its own line, for the reason spelled out below: adding it to the grouped
// import above reflowed that statement across lines, and `@ts-expect-error`
// covers only the NEXT line — so the directive stopped reaching the `from`
// clause. Suite stayed green, `tsc` went red, exactly as predicted here.
// @ts-expect-error — as above.
import { poolEveryDraw } from "../scripts/triage.mjs";
// Its own line: adding it above pushes that import over the print width, and a
// reflowed import moves this directive off the statement it is annotating.
// @ts-expect-error — as above.
import { describeFinding } from "../scripts/triage.mjs";
// Also its own line, and for the reason the comment above gives: adding it to
// the grouped import pushed that statement over the print width, prettier
// reflowed it across lines, and `@ts-expect-error` covers only the NEXT line —
// so the directive stopped reaching the `from` clause where the error is
// reported. The suite stayed green and `tsc` went red.
// @ts-expect-error — as above.
import { batchPopulations } from "../scripts/triage.mjs";
import { buildDeckBase64 } from "../src/render/pptx-deck";
import { buildChart } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";

/**
 * The run/deck joiner.
 *
 * Its whole job is to disagree — to say where the add-in's account of a run
 * and the file that run produced do not match. So the cases that matter are
 * the ones where they differ, and every one of them below is paired with the
 * same slot in the clean deck to prove the verdict came from the difference
 * and not from the tool's default mood.
 */

const cfg = (title: string) => ({ ...sampleConfig("line"), title });

interface Piece {
  title: string;
  slot: number;
  run?: string;
  tagged?: boolean;
}

/** Build a real .pptx from slot descriptions and read it back as the auditor sees it. */
async function deckOf(pieces: Piece[]) {
  const built = await buildDeckBase64(
    pieces.map((p) => ({
      scene: buildChart(cfg(p.title)),
      title: p.title,
      configJson: p.tagged === false ? undefined : JSON.stringify(cfg(p.title)),
      slot: p.slot,
      run: p.run ?? "run-x",
    })),
  );
  return readDeckBytes(Uint8Array.from(atob(built.base64), (c) => c.charCodeAt(0)));
}

/** A run log claiming every named item rendered and was tagged. */
function logOf(titles: string[], over: Record<string, unknown> = {}) {
  return {
    build: "test-build",
    run: "run-x",
    host: "test host",
    totalMs: 1234,
    path: "file",
    items: titles.map((title) => ({
      title,
      status: "rendered",
      shapes: 3,
      ms: 10,
      grouped: true,
      tagged: true,
      chart: true,
      lateOutcome: "",
    })),
    deck: { slidesAdded: titles.length, addsIssued: titles.length, lost: 0, blank: [] },
    ...over,
  };
}

const verdicts = (t: { slots: { title: string; verdict: string }[] }) =>
  Object.fromEntries(t.slots.map((s) => [s.title, s.verdict]));

describe("triage — joining a run log to the deck it produced", () => {
  it("finds no disagreement when the deck holds exactly what the run claimed", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "B", slot: 1 },
      ]),
      logOf(["A", "B"]),
    );
    expect(verdicts(t)).toEqual({ A: "ok", B: "ok" });
    expect(t.disagreements).toBe(0);
    expect(t.run).toBe("run-x");
    expect(t.inferredRun).toBe(false);
  });

  it("calls a slot the deck does not hold lost", async () => {
    // Negative control: B is `ok` above, from the same log, with its slide.
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), logOf(["A", "B"]));
    expect(verdicts(t)).toEqual({ A: "ok", B: "lost" });
    expect(t.disagreements).toBe(1);
  });

  it("does not call a skipped item lost — the run never claimed to draw it", async () => {
    const log = logOf(["A", "B"]);
    log.items[1].status = "skipped";
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), log);
    expect(verdicts(t)).toEqual({ A: "ok", B: "skipped" });
    expect(t.disagreements).toBe(0);
  });

  it("calls a tag the run believed it wrote and the file does not carry tag-lost", async () => {
    // The file path cannot produce this by construction, which is exactly why
    // it is worth naming when it appears: it means the deck lost a tag the
    // add-in had already committed to the bytes.
    const t = triage(await deckOf([{ title: "A", slot: 0, tagged: false }]), logOf(["A"]));
    expect(verdicts(t)).toEqual({ A: "tag-lost" });
    expect(t.disagreements).toBe(1);
  });

  it("calls a tag the run gave up on and the file carries repaired, and not a fault", async () => {
    const log = logOf(["A"]);
    log.items[0].tagged = false;
    const t = triage(await deckOf([{ title: "A", slot: 0 }]), log);
    expect(verdicts(t)).toEqual({ A: "repaired" });
    // The repair pass working is not a disagreement to chase.
    expect(t.disagreements).toBe(0);
  });

  it("calls a chart object with no config not-editable when the run never claimed the tag", async () => {
    const log = logOf(["A"]);
    log.items[0].tagged = false;
    const t = triage(await deckOf([{ title: "A", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ A: "not-editable" });
    expect(t.disagreements).toBe(1);
  });

  it("calls two slides claiming one slot duplicated", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "A again", slot: 0 },
      ]),
      logOf(["A"]),
    );
    expect(verdicts(t)).toEqual({ A: "duplicated" });
    expect(t.slots[0].rows).toBe(2);
    expect(t.disagreements).toBe(1);
  });

  it("counts another run's slides as foreign, never as this run's losses", async () => {
    // The case the run token exists to survive. Before it existed, a second
    // insert into the same deck made every item match two slides and the
    // repair pass deleted a healthy run as the other one's duplicate.
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "A", slot: 0, run: "run-earlier" },
      ]),
      logOf(["A"]),
    );
    expect(verdicts(t)).toEqual({ A: "ok" });
    expect(t.foreign).toHaveLength(1);
    expect(t.foreign[0].run).toBe("run-earlier");
    expect(t.disagreements).toBe(0);
  });

  it("calls a slide carrying a slot this run never issued an orphan", async () => {
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "stray", slot: 7 },
      ]),
      logOf(["A"]),
    );
    expect(t.orphans).toEqual([{ slot: 7, verdict: "orphan", indexes: [2] }]);
    expect(t.disagreements).toBe(1);
  });

  it("infers the run from the deck when the log predates run ids, and says it guessed", async () => {
    const log = logOf(["A", "B"]) as Record<string, unknown>;
    delete log.run;
    const t = triage(
      await deckOf([
        { title: "A", slot: 0 },
        { title: "B", slot: 1 },
      ]),
      log,
    );
    expect(t.run).toBe("run-x");
    expect(t.inferredRun).toBe(true);
    expect(t.disagreements).toBe(0);
  });

  it("leaves an item that was never a chart alone, config tag or not", async () => {
    // The bug that produced this test: a clean 38-of-38 run triaged as seven
    // broken slides. The title page, both contents pages and four elements are
    // drawn as `PowerChart` objects with NO config by design, and the rule
    // "a chart object without a config is not re-editable" flagged every one.
    const log = logOf(["Title"]);
    log.items[0].tagged = false;
    log.items[0].chart = false;
    const t = triage(await deckOf([{ title: "Title", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ Title: "ok" });
    expect(t.disagreements).toBe(0);
  });

  it("will not guess at intent a log is too old to state", async () => {
    // Same deck and the same missing config, with `chart` absent: the tool
    // cannot tell a title page from a chart that lost its tag, so it says so
    // instead of picking one. `tagged` cannot stand in — on the shape path
    // false means both "never had a config" and "the write did not land".
    const log = logOf(["Mystery"]) as { items: Record<string, unknown>[] };
    log.items[0].tagged = false;
    delete log.items[0].chart;
    const t = triage(await deckOf([{ title: "Mystery", slot: 0, tagged: false }]), log);
    expect(verdicts(t)).toEqual({ Mystery: "no-config" });
    expect(t.disagreements).toBe(1);
  });

  it("prefers the log's own run id over the deck's commonest", async () => {
    // A deck where the run under investigation is the MINORITY. Inferring from
    // the deck would pick the other one and report every slot lost.
    const t = triage(
      await deckOf([
        { title: "A", slot: 0, run: "run-mine" },
        { title: "old", slot: 0, run: "run-older" },
        { title: "old", slot: 1, run: "run-older" },
      ]),
      logOf(["A"], { run: "run-mine" }),
    );
    expect(t.run).toBe("run-mine");
    expect(verdicts(t)).toEqual({ A: "ok" });
    expect(t.foreign).toHaveLength(2);
  });
});

describe("triage — a log holding more than one run", () => {
  it("reads each run in a both-paths log separately", async () => {
    // One click can take both insert paths, and they fail in completely
    // different ways. Merging them would produce a report that is true of
    // neither: the file run's slides are present while the shape run's are
    // still missing, in the same deck, at the same moment.
    const deck = await deckOf([
      { title: "A", slot: 0, run: "run-file" },
      { title: "B", slot: 1, run: "run-file" },
      { title: "A", slot: 0, run: "run-shapes" },
    ]);
    const file = { ...logOf(["A", "B"], { run: "run-file", path: "file" }) };
    const shapes = { ...logOf(["A", "B"], { run: "run-shapes", path: "shapes" }) };
    const runs = runsIn({ build: "b", host: "h", runs: [file, shapes] });
    expect(runs).toHaveLength(2);
    expect(verdicts(triage(deck, runs[0]))).toEqual({ A: "ok", B: "ok" });
    // The shape run only landed its first slide, and says so — while the file
    // run beside it in the same deck is clean.
    expect(verdicts(triage(deck, runs[1]))).toEqual({ A: "ok", B: "lost" });
  });

  it("reads a single-run log as a list of one", async () => {
    // Every artifact captured so far is shaped this way. Refusing them would
    // make the tool useless on exactly the runs it was written to explain.
    const flat = logOf(["A"]);
    expect(runsIn(flat)).toEqual([flat]);
  });
});

describe("triage — logs that are not inserts", () => {
  it("finds the host self-test's verdicts in a log that holds no runs", async () => {
    // The self-test writes a log with an empty `runs` list: its scenarios are
    // not inserts and own no slots. Reading only `runs` made the whole report
    // come out blank and exit 0 — indistinguishable from a clean deck, on a
    // file someone handed this tool expecting an answer.
    const log = {
      build: "b",
      host: "h",
      runs: [],
      selftest: [
        { name: "insert twice", ok: true, detail: "deck grew by 4", ms: 10 },
        { name: "explode a picture", ok: false, detail: "config did not survive", ms: 20 },
      ],
    };
    expect(runsIn(log)).toEqual([]);
    expect(selfTestIn(log).map((s: { name: string }) => s.name)).toEqual(["insert twice", "explode a picture"]);
  });

  it("finds nothing to report in a log that is neither", async () => {
    expect(selfTestIn({ build: "b", host: "h", runs: [] })).toEqual([]);
    expect(selfTestIn(undefined)).toEqual([]);
  });

  /**
   * The round's own deck evidence — the half that replaced saving a .pptx and
   * taking a screenshot, and that nothing read until a round had to be worked
   * out by hand.
   *
   * The distinction each case pins is between an empty slide and an empty
   * ANSWER. This host is known to report a shape collection short without
   * throwing (`shapes-items-count-honest`), so a readback of zero is one
   * witness; the host's own picture of the same slide is a second, and only
   * the two together say a slide is really blank.
   */
  const evidenceDeck = (slides: { id: string; count: number; png?: number }[], added?: string[]) => ({
    inventory: slides.map((s, i) => ({ slideId: s.id, index: i, shapes: [], count: s.count })),
    newSlides: added ?? slides.map((s) => s.id),
    // Base64 is 4 characters per 3 bytes, which is the ratio the reader undoes.
    shots: slides.filter((s) => s.png !== undefined).map((s) => ({ slideId: s.id, png: "A".repeat(s.png! * 4) })),
  });

  it("only calls a slide blank when its picture agrees with its readback", () => {
    const e = deckEvidence(
      evidenceDeck([
        { id: "a", count: 12, png: 900 }, // has shapes — not a blank at all
        { id: "b", count: 0, png: 300 }, // empty, and the smallest picture here
        { id: "c", count: 0 }, // empty, no picture — the rasterise cap
        { id: "d", count: 0, png: 900 }, // empty, but the picture has content
      ]),
    );
    expect(e.withShapes).toBe(1);
    expect(e.confirmed, "only the slide whose picture is blank too").toBe(1);
    expect(e.unseen, "no picture is not evidence of an empty slide").toBe(1);
    expect(e.lying, "a blank readback over a picture with content is the host lying").toBe(1);
  });

  it("counts only the slides this round added, never the deck it landed in", () => {
    // A round drops its slides into whatever the user already had. Counting the
    // pre-existing ones would report someone's own empty slides as our losses.
    const e = deckEvidence(
      evidenceDeck(
        [
          { id: "theirs", count: 0 },
          { id: "ours", count: 0 },
        ],
        ["ours"],
      ),
    );
    expect(e.added).toBe(1);
    expect(e.unseen).toBe(1);
  });

  it("names the added slides whose id never finished, and joins them to the blanks", () => {
    // ROUND 041 put both halves of this in the log and nothing put them side by
    // side: seven slides added, two of them named `256#0` and `257#0` against
    // `288#3603562595` for the rest, one added slide blank, and a
    // `delete-by-id left slides behind` line in the same round. It took a hand
    // query to see, which is how a finding nearly went unnoticed.
    const e = deckEvidence(
      evidenceDeck(
        [
          { id: "288#3603562595", count: 24 },
          { id: "256#0", count: 0, png: 300 },
          { id: "257#0", count: 12 },
        ],
        ["288#3603562595", "256#0", "257#0"],
      ),
    );
    expect(e.oddIds, "an id ending #0 is not the shape a finished slide has").toEqual(["256#0", "257#0"]);
    // THE JOIN, which is the whole point: one of the two odd ids is also the
    // blank, so the two observations are one event rather than two.
    expect(e.oddAndBlank).toBe(1);
    expect(e.confirmed).toBe(1);
  });

  it("pools tag faults per build, so two rounds of one build can be compared with themselves", () => {
    // WHY THIS EXISTS: on 2026-08-15 a tag-anchor fix was nearly reported as
    // working because the round carrying it looked like the round before. It
    // did — and so did two rounds with NO renderer change between them, where
    // `tags-undefined` went 1 → 5. Grouping by build is what lets the same build
    // run twice pool instead of compete, and that pair is the only noise
    // measurement that owes nothing to an argument.
    const round = (build: string, faults: { undef?: number; grouped?: number }) => ({
      build,
      trace: {
        entries: [
          ...Array.from({ length: faults.undef ?? 0 }, () => ({
            message: "tagging failed — charts are not re-editable",
            data: { error: "Cannot read properties of undefined (reading 'add')" },
          })),
          ...Array.from({ length: faults.grouped ?? 0 }, () => ({
            message: "grouping refused",
            data: { error: "at=grouping the chart's shapes | code=5010" },
          })),
        ],
      },
    });
    const byBuild = poolTagFaults([
      round("aaaaaaa · 2026-08-15", { undef: 1 }),
      round("aaaaaaa · 2026-08-15", { undef: 5 }),
      round("bbbbbbb · 2026-08-15", { undef: 5, grouped: 2 }),
    ]);
    expect([...byBuild.keys()], "the timestamp is not part of the build").toEqual(["aaaaaaa", "bbbbbbb"]);
    // Two rounds of one build pooled, not averaged away — the SPREAD is the point.
    expect(byBuild.get("aaaaaaa")!.map((r: Record<string, number>) => r["tags-undefined"])).toEqual([1, 5]);
    expect(byBuild.get("aaaaaaa")!.map((r: Record<string, number>) => r["tagging-failed"])).toEqual([1, 5]);
    expect(byBuild.get("bbbbbbb")![0]["group-5010"]).toBe(2);
    // A fault line that is not a tagging failure must not inflate the count that
    // a fix would be judged on.
    expect(byBuild.get("bbbbbbb")![0]["tagging-failed"]).toBe(5);
  });

  it("flags a scenario that WAS passing and has stopped, and stays quiet about the host's mood", () => {
    // THE ONLY AUTOMATIC CHECK ON A ROUND'S OWN RESULT. Rounds 070-072 took
    // `same scale across the deck` from 35 consecutive failures to three
    // consecutive passes, and nothing guarded that: every other number a round
    // produces is read by a person and filed, so a later build could take it
    // back and no gate would fail.
    const round = (pairs: [string, boolean][]) => ({ selftest: pairs.map(([name, ok]) => ({ name, ok })) });
    const passing: [string, boolean][] = [
      ["same scale", true],
      ["visible", true],
    ];

    // Established (three straight passes) and now failing — the case it is for.
    const gone = scenarioRegressions([
      round(passing),
      round(passing),
      round(passing),
      round([
        ["same scale", false],
        ["visible", true],
      ]),
    ]);
    expect(
      (gone as { name: string }[]).map((g) => g.name),
      "a scenario that had passed three rounds running failed unnoticed",
    ).toEqual(["same scale"]);

    // MOOD IS NOT A REGRESSION, and this is what keeps the gate usable. This
    // host swings 4-of-5 to 1-of-5 on the same build with nothing changed; a
    // scenario that fails half the time was never established, so its next
    // failure is not news. A gate that cries wolf gets switched off.
    const flaky = scenarioRegressions([
      round([["same scale", true]]),
      round([["same scale", false]]),
      round([["same scale", true]]),
      round([["same scale", false]]),
    ]);
    expect(flaky, "reported a chronically flaky scenario as a regression").toEqual([]);

    // A NEW scenario cannot have been established, so its first bad round is not
    // a fault — otherwise every scenario added would arrive pre-broken.
    const fresh = scenarioRegressions([
      round([["same scale", true]]),
      round([["same scale", true]]),
      round([["same scale", true]]),
      round([
        ["same scale", true],
        ["brand new", false],
      ]),
    ]);
    expect(fresh, "a scenario on its first outing was called a regression").toEqual([]);

    // And too little history is no history: it must not judge from two rounds.
    expect(scenarioRegressions([round(passing), round([["same scale", false]])])).toEqual([]);
  });

  it("counts every chart that cannot follow a drag, which a passing scenario hides", () => {
    // A GREEN VERDICT OVER A SAMPLE. `an update follows a moved chart` passes and
    // tests ONE chart; rounds 073 and 074 lost the origin tag on 9 of 19 and 8 of
    // 17 charts in those same rounds. Every one of those would snap back to where
    // it was inserted instead of following the user's drag, while the scenario
    // reported the round trip holding.
    //
    // Same shape as `does a rasterise poison the next draw` counting only its own
    // four draws, and the fresh-slide split sitting unqueried for eleven rounds.
    // A scenario samples; a pooled count does not.
    const lost = (n: number) => ({ message: "origin tag lost — the chart is re-editable", data: { charts: n } });
    const o = poolOriginTagLosses([
      { trace: { entries: [lost(9), lost(1)] } },
      { trace: { entries: [] } },
      { trace: { entries: [lost(8)] } },
    ]);
    expect(o.charts, "did not total the charts across rounds").toBe(18);
    expect(o.rounds, "counted a round that lost nothing").toBe(2);
    // TEN, not nine: the first round lost 9 charts on one line and 1 on another,
    // and a round's damage is the SUM of its lines. Getting this wrong first time
    // is why it is asserted — a per-line maximum would under-report every round
    // that failed in more than one batch.
    expect(o.worst, "took the worst LINE rather than the worst ROUND").toBe(10);

    // A round with no losses contributes nothing at all, so the report stays
    // silent on an archive that never had the problem.
    expect(poolOriginTagLosses([{ trace: { entries: [] } }])).toEqual({ rounds: 0, charts: 0, worst: 0 });
  });

  it("keeps 4:3 and 16:9 apart, and never judges one against the other", () => {
    // A NIGHTLY CYCLE RUNS 16:9 TWICE AND 4:3 ONCE. Judged against three 16:9
    // rounds, a 4:3 round would be flagged for scoring differently — which it
    // does, by design: round 077 scored 10 of 13 where 16:9 scored 13 of 13
    // twice. The gate would fire every night, and a gate that cries wolf gets
    // switched off. This file has already watched that happen twice.
    const at = (w: number, h: number, pairs: [string, boolean][]) => ({
      build: "aaaaaaa · 2026-08-16",
      slideSize: { width: w, height: h, source: "pageSetup" },
      selftest: pairs.map(([name, ok]) => ({ name, ok })),
    });
    const wide = (pairs: [string, boolean][]) => at(960, 540, pairs);
    const std = (pairs: [string, boolean][]) => at(720, 540, pairs);

    expect(roundProfile(wide([]))).toBe("16:9");
    expect(roundProfile(std([]))).toBe("4:3");
    // A round with no size is one of the 53 filed before the field existed, and
    // every one of those was 16:9 — a documented fact, not a guess.
    expect(roundProfile({}), "an old round must not become its own profile").toBe("16:9");
    // Anything else is visibly itself rather than folded into the nearest named.
    expect(roundProfile(at(1000, 500, []))).toBe("1000x500");

    // THE GATE: a 4:3 round after three passing 16:9 rounds is not a regression.
    const ok: [string, boolean][] = [["same scale", true]];
    expect(
      scenarioRegressions([wide(ok), wide(ok), wide(ok), std([["same scale", false]])]),
      "judged a 4:3 round against 16:9 history",
    ).toEqual([]);
  });

  it("names a scenario that passes at one slide size and fails at another", () => {
    // THE OTHER QUESTION, and it needs its own answer. The gate asks whether a
    // scenario fell against its OWN history; this asks whether one profile
    // failed what another passed on the SAME BUILD. Round 077 was exactly that
    // — 10 of 13 at 4:3 against 13 of 13 at 16:9 — and nothing said so.
    const at = (build: string, w: number, h: number, pairs: [string, boolean][]) => ({
      build: `${build} · 2026-08-16`,
      slideSize: { width: w, height: h, source: "pageSetup" },
      selftest: pairs.map(([name, ok]) => ({ name, ok })),
    });
    const d = profileDivergence([
      at("aaaaaaa", 960, 540, [
        ["same scale", true],
        ["visible", true],
      ]),
      at("aaaaaaa", 720, 540, [
        ["same scale", false],
        ["visible", true],
      ]),
    ]);
    expect(d.map((x: { name: string }) => x.name)).toEqual(["same scale"]);
    expect(d[0].passedIn).toEqual(["16:9"]);
    expect(d[0].failedIn).toEqual(["4:3"]);
    expect(d[0].flaky, "a clean difference between profiles must not be called flaky").toBeFalsy();

    // FAILING IN BOTH is an ordinary bug, not divergence — reporting it here
    // would bury the one signal this exists for.
    expect(
      profileDivergence([
        at("aaaaaaa", 960, 540, [["same scale", false]]),
        at("aaaaaaa", 720, 540, [["same scale", false]]),
      ]),
      "reported a plain bug as a slide-size difference",
    ).toEqual([]);

    // A PROFILE THAT DISAGREES WITH ITSELF IS FLAKY, NOT DIFFERENT — and the
    // check's first live outing got this wrong. `explode a degraded picture`
    // passed at 4:3, then passed once and failed once at 16:9 on one build.
    // Collapsing a profile to its worst outcome reported "diverged between
    // slide sizes", which was true of the worst reading and wrong about the
    // cause — and it would send someone to investigate an aspect ratio for a
    // scenario that is simply unreliable.
    const unstable = profileDivergence([
      at("aaaaaaa", 720, 540, [["explode", true]]),
      at("aaaaaaa", 960, 540, [["explode", true]]),
      at("aaaaaaa", 960, 540, [["explode", false]]),
    ]);
    expect(unstable.map((x: { name: string }) => x.name)).toEqual(["explode"]);
    expect(unstable[0].flaky, "reported a profile disagreeing with itself as a slide-size difference").toBe(true);
    expect(unstable[0].unstableIn).toEqual(["16:9"]);
    // And it must NOT claim the stable profile failed anything.
    expect(unstable[0].failedIn).toEqual([]);

    // ACROSS BUILDS IS NOT A COMPARISON. Two rounds on different builds differ
    // for reasons that have nothing to do with slide size.
    expect(
      profileDivergence([
        at("aaaaaaa", 960, 540, [["same scale", true]]),
        at("bbbbbbb", 720, 540, [["same scale", false]]),
      ]),
      "compared two builds and blamed the slide size",
    ).toEqual([]);
  });

  it("does not call a scenario that DID NOT MEASURE a regression", () => {
    // THE GATE'S FIRST LIVE OUTING GOT THIS WRONG. Round 073 flagged `explode a
    // degraded picture` as having stopped passing, for a result whose own words
    // were "the slide not naming it proves nothing either way" — the host had
    // refused ids mid-scenario, so nothing was measured either way.
    //
    // An absence of evidence is not a fall, and a gate that fires on a scenario
    // declining to conclude is a gate that gets switched off. This repo has
    // already watched that happen to one.
    const skipped = (name: string) => ({ name, ok: false, skipped: true });
    const passed = (name: string) => ({ name, ok: true });
    const failed = (name: string) => ({ name, ok: false });

    expect(
      scenarioRegressions([
        { selftest: [passed("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [skipped("explode")] },
      ]),
      "a scenario that measured nothing was reported as having fallen",
    ).toEqual([]);

    // A GENUINE failure after the same history still fires, or the fix would
    // have bought quiet at the price of the gate.
    expect(
      scenarioRegressions([
        { selftest: [passed("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [failed("explode")] },
      ]).map((g: { name: string }) => g.name),
    ).toEqual(["explode"]);

    // And a scenario that SKIPPED earlier was never established, so its later
    // failure is not a regression either — it has no passing run to fall from.
    expect(
      scenarioRegressions([
        { selftest: [passed("explode")] },
        { selftest: [skipped("explode")] },
        { selftest: [passed("explode")] },
        { selftest: [failed("explode")] },
      ]),
      "treated a skipped round as a passing one when establishing",
    ).toEqual([]);
  });

  it("splits grouping by whether the chart spanned batches, and counts a draw once", () => {
    // THE SHARPEST SEPARATION IN THE ARCHIVE: 353 of 452 multi-batch draws
    // grouped, against 49 of 214 single-batch. It is OURS — `refreshShapes` is
    // set from `spansBatches()`, so only a multi-batch chart gets the
    // pre-grouping re-read that resolves its shapes by id; a single-batch chart
    // hands addGroup the raw created proxies and this host refuses them.
    //
    // Found while chasing what looked like a RASTERISE effect (22% vs 93%).
    // Splitting by batch count collapsed that to 22% vs 27% — nothing. The
    // scenarios that rasterise just draw small charts. Counting only the LAST
    // batch of a draw is what makes the two arms comparable at all.
    const batch = (upTo: number, total: number) => ({ message: "batch issued", data: { upTo, total } });
    const b = poolBatchSpanVsGroup([
      {
        trace: {
          entries: [
            // A 24-shape chart: three batches, ONE draw. Counting per batch
            // would credit this arm three times for a single outcome.
            batch(10, 24),
            batch(20, 24),
            batch(24, 24),
            { message: "grouped the chart's shapes", data: {} },
            // A single-batch chart whose group was refused.
            batch(7, 7),
            { message: "grouping the chart's shapes", data: { error: "InvalidParam ... code=5010" } },
            // A single-batch chart that DID group — the arm is not all failure.
            batch(9, 9),
            { message: "grouped the chart's shapes", data: {} },
            // A draw the round never resolved either way belongs to neither arm.
            batch(6, 6),
          ],
        },
      },
    ]);
    expect(b.multi, "a multi-batch draw was counted once per batch, not once per draw").toBe(1);
    expect(b.multiGrouped).toBe(1);
    expect(b.single).toBe(2);
    expect(b.singleGrouped).toBe(1);
  });

  it("names the questions that have never once answered, and splits them by whose fault it is", () => {
    // SIX QUESTIONS WERE 0-FOR-41 and nothing said so. The per-round report
    // names what was "never put" in THAT round, so a permanently dead question
    // read as bad luck, forty-one times running — while costing a scratch slide
    // and host time on every one of them.
    //
    // Four of the six were the group cluster. The probe has been blind on groups
    // for the whole archive, and rounds 064/065 then answered the most important
    // of them (`tag-on-group-survives`) from PRODUCTION, twice, in one evening.
    //
    // The split is the point. "Never asked" is a harness problem and ours to
    // fix; "unreadable" is the host declining to answer, which is a finding.
    // Pooling them would hide exactly the distinction that decides what to do.
    const round = (answers: { id: string; answer: string }[]) => ({ hostAnswers: { answers } });
    const dead = poolStarvedQuestions([
      round([
        { id: "tag-on-group-survives", answer: "no-scratch-slide" },
        { id: "addgroup-returns-usable", answer: "unreadable" },
        { id: "sometimes-answers", answer: "no-scratch-shape" },
        { id: "always-answers", answer: "yes" },
      ]),
      round([
        { id: "tag-on-group-survives", answer: "no-scratch-slide" },
        { id: "addgroup-returns-usable", answer: "unreadable" },
        { id: "sometimes-answers", answer: "survives-8" },
        { id: "always-answers", answer: "yes" },
      ]),
    ]);
    // Typed at the boundary: `triage.mjs` is plain JS behind a @ts-expect-error
    // import, so everything out of it is `any` and an untyped callback param is
    // an implicit-any typecheck error — green locally under vitest, red in CI.
    type Dead = { id: string; rounds: number; never: number; unanswerable: number; answered: number };
    const ids = (dead as Dead[]).map((d) => d.id);
    expect(ids, "a question that answers SOMETIMES is doing its job, not waste").not.toContain("sometimes-answers");
    expect(ids).not.toContain("always-answers");
    expect([...ids].sort()).toEqual(["addgroup-returns-usable", "tag-on-group-survives"]);

    const byId = Object.fromEntries((dead as Dead[]).map((d) => [d.id, d]));
    // Never asked — the harness could not set it up.
    expect(byId["tag-on-group-survives"]).toMatchObject({ never: 2, unanswerable: 0, answered: 0, rounds: 2 });
    // Asked, and the host declined — a different problem with a different fix.
    expect(byId["addgroup-returns-usable"]).toMatchObject({ never: 0, unanswerable: 2, answered: 0, rounds: 2 });

    // A question seen only once is not yet a pattern, and calling it dead would
    // put every newly-added probe into a report about waste on its first round.
    expect(poolStarvedQuestions([round([{ id: "brand-new", answer: "no-scratch-slide" }])])).toEqual([]);
  });

  it("separates a chart on a fresh slide from one on a slide that already had shapes", () => {
    // THE ROOT, and the cleanest separation this project has: 82 charts on a
    // slide that already had shapes, 81 grouped; 74 on a freshly added empty
    // slide, 1 grouped. A chart that is not grouped loses its config, so the
    // slide's newness decides the config — one level below everything the tag
    // work was aimed at.
    const chart = (name: string, onSlide: number, msgs: string[]) =>
      msgs.map((m, i) => ({ message: m, data: i === 0 ? { chart: name, onSlide } : { chart: name } }));
    const f = poolFreshVsEstablished([
      {
        trace: {
          entries: [
            ...chart("1/4", 32, ["batch issued", "grouped the chart's shapes"]),
            ...chart("2/4", 0, ["batch issued", "not grouping: no member handle this host will accept"]),
            ...chart("3/4", 0, ["batch issued", "grouped the chart's shapes"]),
            // Never decided — the round did not reach it, so it belongs to
            // neither column rather than being counted as a failure.
            ...chart("4/4", 0, ["batch issued"]),
          ],
        },
      },
    ]);
    expect(f.established).toBe(1);
    expect(f.establishedGrouped).toBe(1);
    expect(f.fresh, "a chart the round never decided must not be counted").toBe(2);
    expect(f.freshGrouped, "a fresh slide CAN group — the claim is a rate, not a law").toBe(1);
  });

  it("separates the charts that got grouped from the ones that did not", () => {
    // THE QUESTION NOBODY ASKED FOR ELEVEN ROUNDS, and the archive held the
    // answer throughout: pooled over every round, 64 grouped charts lost 1 tag
    // and 62 ungrouped charts lost 41. A chart that groups keeps its config; one
    // that cannot loses it two times in three.
    //
    // It reframes the tag work. Grouping puts the tag on the GROUP — a handle
    // made in that batch and never resolved — and it lands. No group means the
    // tag falls back to a `created` handle, which is the path four rounds and a
    // renderer change went into. That is a question about the losing side.
    const chart = (name: string, msgs: string[]) => msgs.map((m) => ({ message: m, data: { chart: name } }));
    const g = poolGroupVsTag([
      {
        trace: {
          entries: [
            ...chart("1/3", ["grouped the chart's shapes"]),
            ...chart("2/3", ["grouped the chart's shapes", "tagging failed — charts are not re-editable"]),
            ...chart("3/3", ["not grouping: no member handle this host will accept", "tagging failed — nope"]),
            // No chart key at all: batch-level noise that must not be counted as
            // a chart in either column.
            { message: "tagging failed — charts are not re-editable", data: {} },
          ],
        },
      },
    ]);
    expect(g.grouped).toBe(2);
    expect(g.groupedLost, "a grouped chart CAN still lose its tag — the claim is a rate, not a law").toBe(1);
    expect(g.ungrouped).toBe(1);
    expect(g.ungroupedLost).toBe(1);
  });

  it("says nothing about ids when every one of them finished", () => {
    // A field that reports on a healthy round is a field a reader learns to skip.
    const e = deckEvidence(evidenceDeck([{ id: "288#3603562595", count: 24 }], ["288#3603562595"]));
    expect(e.oddIds).toEqual([]);
    expect(e.oddAndBlank).toBe(0);
  });

  it("reads a slide the scan could only partly list as carrying shapes", () => {
    // `count` is the host's own number, `shapes` is what the scan managed to
    // list. Taking the smaller would call a partial listing empty — the exact
    // mistake of trusting one witness that this whole block exists to avoid.
    const e = deckEvidence({
      inventory: [{ slideId: "a", index: 0, shapes: [{ id: "1" }], count: 24 }],
      newSlides: ["a"],
      shots: [],
    });
    expect(e.withShapes).toBe(1);
    expect(e.unseen).toBe(0);
  });

  /**
   * Pooling the counterbalanced arms, and the trap under it.
   *
   * The per-batch line is logged BEFORE the sync, deliberately — the sync is
   * where a bad host goes quiet, so the number has to be on screen while you
   * wait. That means every stall has one of its own immediately before it, and
   * reading those lines as successes is wrong twice over: it counts a failure
   * as a success, and it inflates the denominator with the same event.
   *
   * Both mistakes were made analysing this data by hand. One pass produced 0
   * stalls in 32 draws; another produced a 6x rasterise effect that was not
   * there. The outcome of a draw is whether a `gave up waiting` lands between
   * that draw's step and the next one — nothing else.
   *
   * The line was CALLED `batch committed` while both of those went wrong, which
   * is why the writer does not call it that any more (`batch issued`). These
   * fixtures keep the old name on purpose: the owner's saved rounds carry it,
   * and a pooling function that reads either name is a pooling function that
   * reads neither — which is the property being pinned here.
   */
  const drawStep = (ms: number, arm: string, i: number) => ({
    ms,
    scope: "selftest",
    message: "rasterise-draw step",
    data: { what: `after a ${arm} #${i}: drawing` },
  });
  const committed = (ms: number) => ({ ms, scope: "draw", message: "batch committed", data: { upTo: 7, total: 7 } });
  const gaveUp = (ms: number) => ({
    ms,
    scope: "host",
    message: "gave up waiting",
    data: { what: "drawing shapes 1-7 of 7", afterMs: 45000 },
  });

  it("counts a draw that was SENT and never landed as a stall, not a success", () => {
    const log = {
      trace: {
        entries: [
          drawStep(10, "rasterise", 0),
          committed(11),
          gaveUp(60),
          drawStep(70, "cheap read", 1),
          committed(71),
        ],
      },
    };
    const { arms } = poolRasteriseArms([log]);
    expect(arms.rasterise, "a committed-then-abandoned draw was read as a success").toEqual({ ok: 0, stall: 1 });
    expect(arms["cheap read"]).toEqual({ ok: 1, stall: 0 });
  });

  it("adds the arms up across rounds, which is the whole point", () => {
    // One round gives two draws an arm — enough to catch a call that fails
    // every time, hopeless for anything smaller. Eight rounds of that sat in
    // separate files, each independently reporting "no pattern".
    const round = (stallArm: string) => ({
      trace: {
        entries: [
          drawStep(10, "rasterise", 0),
          committed(11),
          ...(stallArm === "rasterise" ? [gaveUp(60)] : []),
          drawStep(70, "cheap read", 1),
          committed(71),
          ...(stallArm === "cheap read" ? [gaveUp(120)] : []),
        ],
      },
    });
    const { rounds, arms } = poolRasteriseArms([round("rasterise"), round("rasterise"), round("cheap read")]);
    expect(rounds).toBe(3);
    expect(arms.rasterise).toEqual({ ok: 1, stall: 2 });
    expect(arms["cheap read"]).toEqual({ ok: 2, stall: 1 });
  });

  it("says nothing about a log that never ran the scenario", () => {
    expect(poolRasteriseArms([{ trace: { entries: [committed(1)] } }, {}]).arms.rasterise).toEqual({ ok: 0, stall: 0 });
  });

  it("counts the draws the scenario's four arms never see", () => {
    // Round 28 in miniature, and the reason `poolEveryDraw` exists. The
    // visibility scenario rasterises and then draws; that draw stalls. It
    // carries no arm tag, so the arms report a clean round while the round
    // contains the very thing they are looking for.
    const rasterised = {
      ms: 500,
      scope: "selftest",
      message: "visibility step",
      data: { what: "rasterising a slide" },
    };
    // `batch issued`, NOT the `committed` helper above: that one still emits the
    // pre-2026-08-11 name `batch committed`, which this file's own comment says
    // is read by nothing here, by design. Using it made the population come back
    // empty and looked like a classifier bug.
    const issued = (ms: number) => ({ ms, scope: "draw", message: "batch issued", data: {} });
    const log = {
      trace: {
        entries: [
          // The scenario's own arms: both land, so the arms see nothing.
          drawStep(10, "rasterise", 0),
          committed(11),
          drawStep(70, "cheap read", 1),
          committed(71),
          // An untagged draw, straight after a rasterise, that stalls.
          rasterised,
          issued(501),
          gaveUp(560),
          // An untagged draw with no rasterise before it, that lands.
          issued(900),
        ],
      },
    };
    expect(poolRasteriseArms([log]).arms.rasterise, "the arms should still see a clean round").toEqual({
      ok: 1,
      stall: 0,
    });
    const { after } = poolEveryDraw([log]);
    expect(after.rasterise, "the untagged stall after a rasterise was not counted").toEqual({ ok: 0, stall: 1 });
    expect(after["anything else"].ok).toBeGreaterThanOrEqual(2);
    expect(after["anything else"].stall).toBe(0);
  });

  it("blames the rasterise nearest the draw, not one from earlier in the round", () => {
    // Without resetting at each draw, one rasterise early on would tar every
    // draw after it and the population would be meaningless.
    const issued = (ms: number) => ({ ms, scope: "draw", message: "batch issued", data: {} });
    const log = {
      trace: {
        entries: [
          { ms: 10, scope: "selftest", message: "visibility step", data: { what: "rasterising a slide" } },
          issued(20), // after the rasterise
          issued(30), // NOT after it — a draw intervened
          gaveUp(40),
        ],
      },
    };
    const { after } = poolEveryDraw([log]);
    expect(after.rasterise, "the first draw follows the rasterise and landed").toEqual({ ok: 1, stall: 0 });
    expect(after["anything else"], "the second draw does not follow a rasterise, and stalled").toEqual({
      ok: 0,
      stall: 1,
    });
  });

  it("has nothing to say about a log that carries no deck evidence", () => {
    expect(deckEvidence(undefined)).toBeNull();
    expect(deckEvidence({ inventory: [], newSlides: [], shots: [] })).toBeNull();
  });

  it("calls a slide tagged with a slot before the run's range an orphan", async () => {
    // Only slots PAST the end were checked, so a negative slot — a tag written
    // wrong, or carried in from something that was not this run — matched no
    // item, fell outside the orphan test, and was reported by nothing.
    const deck = await deckOf([
      { title: "A", slot: 0 },
      { title: "stray", slot: -3 },
    ]);
    const t = triage(deck, logOf(["A"]));
    expect(t.orphans).toEqual([{ slot: -3, verdict: "orphan", indexes: [2] }]);
    expect(t.disagreements).toBe(1);
  });
});

describe("triage — naming the host bug behind a failure", () => {
  it("recognises the signatures that belong to PowerPoint, not to PowerChart", () => {
    // The first hour of the last diagnosis went into establishing that
    // `InvalidParam passed to GetItem(id)` was a host bug rather than ours. A
    // reader who sees an issue number stops looking for the mistake in this
    // repo — which is the whole saving.
    expect(knownBug("InvalidParam passed to GetItem(id) | code=5010")).toMatch(/office-js#2903/);
    expect(knownBug("The property 'items' is not available.")).toMatch(/office-js#6363/);
    expect(knownBug("did not respond while drawing shapes 1-10 of 39")).toMatch(/office-js#5022/);
  });

  it("sends a wedged selection to the selection bugs, not to the generic sync hang", () => {
    // A wedged selection subsystem does not throw, it goes quiet — so what the
    // trace records is a TIMEOUT, and a timeout's text is "did not respond
    // while <phase>". The generic sync-hang entry matches that too, and on a
    // first-match-wins table it would claim every one of them and point the
    // reader at #5022: a different bug, with a different cause and no fix that
    // applies. The phase name is all that separates them, so the phases that
    // belong to selection must be matched BEFORE the generic signature.
    for (const raw of [
      "PowerPoint did not respond while reading the selected chart (90.0s)",
      "PowerPoint did not respond while selecting a shape (90.0s)",
      "PowerPoint did not respond while clearing the shape selection (10.0s)",
    ]) {
      expect(knownBug(raw), raw).toMatch(/#3083/);
      expect(knownBug(raw), raw).not.toMatch(/5022/);
    }
    // The scenario's own verdict line lands on the same note.
    expect(
      knownBug(
        "the host stopped answering selection calls after a programmatic select — known web-host limitation, " +
          "same family as office-js#3083 / #3698; the pane's own Edit-it path is unaffected",
      ),
    ).toMatch(/#3083/);
    // And a hang somewhere else still goes where it always did.
    expect(knownBug("PowerPoint did not respond while reading slides 20-38 for charts (90.0s)")).toMatch(/5022/);
  });

  it("says nothing about a failure that is ours", () => {
    // The mapping must not turn every problem into somebody else's. A reason
    // with no known host bug behind it gets no note, and stays ours to fix.
    expect(knownBug("chart is one object but carries no config tag")).toBeNull();
    expect(knownBug("empty slide left behind by a lost add")).toBeNull();
    expect(knownBug("")).toBeNull();
  });
});

/**
 * The tool's own documented invocation must not print LESS than the degraded one.
 *
 * `reportTrace` was reachable only from the no-deck branch, so
 * `triage.mjs <deck.pptx> <run-log.json>` — the form in the usage line and in
 * CLAUDE.md, the one you use when you actually have a deck — dropped the entry
 * histogram, "phases an error escaped", the problems tally and every
 * `known host bug: office-js#…` annotation, and said nothing about it. A round
 * with a trace and no self-test went further and reported "this log holds no
 * runs and no self-test" over 186 entries, exit 0.
 *
 * Driven through the CLI rather than the exported functions on purpose: what
 * was wrong was the WIRING, and every function involved was already correct.
 */
describe("triage's two invocations", () => {
  const run = (args: string[]) =>
    spawnSync(process.execPath, ["scripts/triage.mjs", ...args], { encoding: "utf8", timeout: 60_000 });

  it("reports the trace whether or not a deck was passed", async () => {
    const log = {
      build: "test-build",
      host: "test-host",
      runs: [],
      trace: {
        summary: {
          steps: [{ scope: "draw", message: "batch committed", n: 2 }],
          problems: [
            {
              text: "PowerPoint did not respond while drawing shapes 1-10 of 24 (45s) | at=drawing the chart's shapes",
              n: 1,
            },
          ],
        },
        entries: [
          { ms: 1, scope: "draw", message: "batch committed" },
          { ms: 2, scope: "draw", message: "batch committed" },
        ],
      },
      selftest: [],
    };
    const dir = mkdtempSync(join(tmpdir(), "pc-triage-"));
    const logPath = join(dir, "round.json");
    writeFileSync(logPath, JSON.stringify(log));

    const withoutDeck = run([logPath]);
    const withDeck = run(["examples/showcase.pptx", logPath]);

    // The trace is a property of the FILE, like the structural faults, so both
    // forms must show it — and in particular the problem line, which is the
    // most locating thing in any round.
    expect(withoutDeck.stdout).toMatch(/TRACE 2 entries/);
    expect(withDeck.stdout, "the deck path dropped the whole trace section").toMatch(/TRACE 2 entries/);
    expect(withDeck.stdout).toMatch(/did not respond while drawing/);
    // And a round that carries a trace is never described as holding nothing.
    expect(withDeck.stdout).not.toMatch(/holds no runs and no self-test/);
    rmSync(dir, { recursive: true, force: true });
  }, 120_000);
});

/**
 * A picture can be called blank only if it looks like nothing.
 *
 * The baseline used to be `Math.min` over the very shots being classified, so
 * the smallest picture satisfied `<=` against itself and at least one added
 * slide was ALWAYS reported under "read back empty AND rasterise blank" — the
 * two-witness line a maintainer reads as proven data loss.
 */
describe("deckEvidence's blank test", () => {
  const png = (bytes: number) => "A".repeat(Math.ceil((bytes * 4) / 3));
  const deckOf = (sizes: number[]) => ({
    newSlides: sizes.map((_, i) => `s${i}`),
    inventory: sizes.map((_, i) => ({ slideId: `s${i}`, index: i, shapes: [], count: 0 })),
    shots: sizes.map((b, i) => ({ slideId: `s${i}`, png: png(b) })),
  });

  it("does not call the smallest CHART on the slide a blank", () => {
    // Every slide reads back empty (the documented short-collection answer) and
    // every picture is chart-sized. The honest verdict is that the readback is
    // lying — the old code confirmed one of them blank and sent the reader
    // after a drawing bug that is not there.
    const e = deckEvidence(deckOf([5142, 5307, 5307, 5319, 5322, 5322]))!;
    expect(e.confirmed, "a 5kB picture of a chart was called blank").toBe(0);
    expect(e.lying).toBe(6);
  });

  it("still confirms a real blank, and does so with no chart in the round to contrast against", () => {
    // Rounds 12 and 13 ran with the picture cap at 12, so every shot they took
    // is a blank. A test that judges by contrast within the round calls those
    // 23 genuinely blank slides content — which is why the ceiling is absolute.
    const mixed = deckEvidence(deckOf([1146, 1146, 5322, 5319]))!;
    expect(mixed.confirmed).toBe(2);
    expect(mixed.lying).toBe(2);
    const allBlank = deckEvidence(deckOf([1146, 1146, 1146]))!;
    expect(allBlank.confirmed, "a round with nothing to contrast against lost its blanks").toBe(3);
  });
});

/**
 * A crashed round's verdicts are banked as they land now, so the file carries
 * them — and this is the reader that has to show them, or they are invisible in
 * a different place.
 */
describe("triage shows what a crashed run had already concluded", () => {
  it("reads a scenario verdict as its verdict", () => {
    expect(describeFinding({ name: "insert a chart", ok: true, ms: 2400 })).toContain("passed");
    expect(describeFinding({ name: "same scale", ok: false, detail: "3 of 8 charts", ms: 34000 })).toContain(
      "FAILED — 3 of 8 charts",
    );
    expect(describeFinding({ name: "picked", ok: false, skipped: true, detail: "not reached" })).toContain("SKIPPED");
  });

  it("reads an answer sheet as how much of it got answered", () => {
    const sheet = { answers: [{ answer: "yes" }, { answer: "threw" }, { answer: null }] };
    expect(describeFinding(sheet)).toBe("answer sheet, 2 of 3 question(s) answered");
  });

  it("falls back to the value rather than dropping it", () => {
    expect(describeFinding("[dropped: 200000 bytes, over the cap]")).toContain("dropped");
    expect(describeFinding({ odd: 1 })).toBe('{"odd":1}');
  });
});

/**
 * A config the pane cannot reach is not a chart the user can edit.
 *
 * `verify-deck.mjs` draws that distinction on the field itself: `config` only
 * reports that the tag part EXISTS, which is equally true of a tag nothing in
 * the shape tree points at. The pane loads a chart by walking from the
 * PowerChart object to its tag rid, so a config anchored on the SLIDE is
 * unreachable and the chart is not re-editable.
 *
 * triage was never migrated when that field was added. For the exact case
 * `verify-deck` was extended to catch, its slot table printed `ok`, its deck
 * column printed `config`, and `disagreements` — the number the final line and
 * the `--json` payload both carry — was 0, while verify-deck's own report on
 * the same deck said `orphan`.
 */
describe("triage reads the config the pane can actually reach", () => {
  const log = { run: "r1", items: [{ title: "chart", status: "ok", tagged: true, chart: true }] };
  const row = (extra: Record<string, unknown>) => ({
    index: 0,
    slot: 0,
    run: "r1",
    shapes: 3,
    groups: 1,
    chartObject: true,
    ...extra,
  });

  it("calls an orphaned config what it is, not ok", () => {
    const t = triage({ rows: [row({ config: true, configOnChart: false, configOrphaned: true })] }, log);
    expect(t.slots[0].verdict, "a config the pane cannot load was reported as fine").not.toBe("ok");
    expect(t.disagreements, "an unreachable config counted as agreement").toBeGreaterThan(0);
    expect(t.slots[0].deck).toContain("ORPHANED");
  });

  it("still calls a properly anchored config ok", () => {
    const t = triage({ rows: [row({ config: true, configOnChart: true, configOrphaned: false })] }, log);
    expect(t.slots[0].verdict).toBe("ok");
    expect(t.disagreements).toBe(0);
  });

  it("falls back to `config` for a deck read before the anchor was reported", () => {
    // Unknown must not become a fault: an older verify-deck did not say where
    // the tag was anchored, and the two agree wherever it does.
    const t = triage({ rows: [row({ config: true })] }, log);
    expect(t.slots[0].verdict).toBe("ok");
  });
});

describe("the two populations a draw batch falls into", () => {
  /**
   * "FAST IS THE BROKEN MODE" rests on the batch times being bimodal with an
   * empty band, and every number behind it was read off a round by hand — which
   * is how the band's edge stayed quoted at 29.2s after a later round put
   * surviving batches at 29.3, 30.8 and 31.1.
   *
   * The load-bearing case is the NEGATIVE one. A split taken at the largest gap
   * always exists, so "bimodal" would be true of any data at all; it is only
   * reported when the gap is bigger than the fast group's whole spread, which is
   * what "an empty band" means. Same shape as `untested` in `host-regimes.mjs`:
   * a verdict that cannot come out the other way is not a measurement.
   */
  const log = (ms: number[]) => ({
    trace: { entries: ms.map((prevBatchMs) => ({ message: "batch issued", data: { prevBatchMs } })) },
  });

  it("finds the empty band in a genuinely bimodal round", () => {
    // Round 17's own twenty-two timed batches, not invented ones. A first
    // attempt at synthetic data put three slow points far apart, which moved
    // the largest gap INSIDE the slow group and correctly answered `false` —
    // the guard working, and a reminder that this shape needs a real
    // distribution rather than a sketch of one.
    const p = batchPopulations(
      log([
        3326, 3344, 3540, 3637, 3717, 5049, 5127, 5220, 5307, 5409, 11057, 13118, 18016, 18240, 18491, 18875, 18927,
        19114, 20175, 21963, 25572, 26360,
      ]),
    )!;
    expect(p.bimodal).toBe(true);
    expect(p.fast[p.fast.length - 1]).toBe(5409);
    expect(p.slow[0]).toBe(11057);
    expect(p.gap).toBe(5648);
    expect(p.max).toBe(26360);
  });

  it("refuses to call an evenly spread round bimodal", () => {
    // Every set of numbers has a largest gap, so a split always exists. On
    // uniform times the argmax picks the first, leaving a one-member fast group
    // whose spread is zero — which any gap beats.
    expect(batchPopulations(log([3000, 4000, 5000, 6000, 7000, 8000]))!.bimodal).toBe(false);
  });

  it("refuses a gap no wider than the cluster it is supposed to separate", () => {
    // THE CASE THAT PINS THE SPREAD COMPARISON, and the test above does not:
    // there the length check already answers, so dropping `gap > spread` left
    // the suite green. Here both groups have two or more members and the gap is
    // exactly the fast group's own spread — an interval, not an empty band.
    const p = batchPopulations(log([1000, 2000, 3000, 4000, 5000, 9000, 10000]))!;
    expect(p.fast.length).toBeGreaterThanOrEqual(2);
    expect(p.slow.length).toBeGreaterThanOrEqual(2);
    expect(p.bimodal).toBe(false);
  });

  it("says nothing at all on too few timed batches", () => {
    // A round whose draws were mostly first batches carries no `prevBatchMs`,
    // and three points cannot show a population.
    expect(batchPopulations(log([3000, 9000, 20000]))).toBeNull();
    expect(batchPopulations({ trace: { entries: [] } })).toBeNull();
    expect(batchPopulations(undefined)).toBeNull();
  });
});

/**
 * The trace is 95K characters a person cannot count, so the thing that counts it
 * has to be right about WHICH KIND of difference it found.
 */
describe("what the newest round said that the archive has not", () => {
  const entry = (scope: string, message: string, error?: string) => ({
    scope,
    message,
    ...(error ? { data: { error } } : {}),
  });
  const round = (build: string, es: ReturnType<typeof entry>[]) => ({ build, trace: { entries: es } });
  const times = (n: number, e: ReturnType<typeof entry>) => Array.from({ length: n }, () => e);
  /** Enough prior rounds to earn a baseline, all saying the same quiet thing. */
  const quietPriors = (n: number, per = 3) =>
    Array.from({ length: n }, (_, i) => round(`old${i}`, times(per, entry("draw", "batch issued"))));

  it("counts one event happening twice as one signature, not two", () => {
    // Ids and counts are the noise. Without this the archive's vocabulary grows
    // by one every time a slide gets a different id, and nothing ever has a
    // history to be measured against.
    expect(traceSignature(entry("group", "repaired 3 tags on slide 7f2a91c4"))).toBe(
      traceSignature(entry("group", "repaired 5 tags on slide 0b41ee02")),
    );
  });

  it("keeps the error class, because a failure is not its healthy twin", () => {
    // The reading this was added for: round 077's 52 `UnexpectedError`s share a
    // message with the successful writes and lived in `data.error`, so a
    // signature over scope+message alone reported the loudest fact in that round
    // as ordinary traffic.
    expect(traceSignature(entry("error", "writing the chart's config tag", "UnexpectedError: nope"))).not.toBe(
      traceSignature(entry("error", "writing the chart's config tag")),
    );
    // The class only. Error text carries ids and offsets, and hashing those
    // would give every single failure its own signature.
    expect(traceSignature(entry("error", "writing", "UnexpectedError: id 7f2a91c4 at 33"))).toBe(
      traceSignature(entry("error", "writing", "UnexpectedError: id 0b41ee02 at 91")),
    );
  });

  it("calls a shape the archive has never produced NOVEL", () => {
    const rounds = [...quietPriors(6), round("new", [entry("group", "something nobody has traced before")])];
    const out = traceNovelty(rounds);
    expect(out.novel).toHaveLength(1);
    expect(out.novel[0].sig).toContain("something nobody has traced before");
    expect(out.spikes).toHaveLength(0);
  });

  // THE REGRESSION GUARD FOR THE COLLAPSE BUG. `profileDivergence` shipped
  // reporting a flaky scenario as a slide-size difference because it kept one
  // worst-case reading of two different causes. This is the same split, and
  // getting it wrong here means every fix this project lands is reported as a
  // fault on the night it starts working.
  it("separates a mechanism that just shipped from a count that left its baseline", () => {
    const reread = entry("group", "re-reading the slide's shapes again after a settle delay");
    const tagWrite = entry("error", "writing the chart's config tag", "UnexpectedError: x");
    const priors = Array.from({ length: 8 }, (_, i) =>
      round(`old${i}`, [
        ...times(3, tagWrite),
        ...times(2, entry("draw", "batch issued")),
        // RARE, NOT ABSENT — and modelling it as absent is what made the first
        // version of this test fail against a correct implementation. The
        // re-read existed before build 17a8204; `needsPreGroupRefresh` widened
        // which charts reach it, so the count went from occasional to eleven.
        // A signature nothing has EVER produced is a different report (NOVEL),
        // and the archive proves the difference: at round 077 the
        // `UnexpectedError` signatures were novel, while the settle-pass repair
        // beside them was this bucket.
        ...(i === 7 ? [reread] : []),
      ]),
    );
    const newest = round("17a8204", [
      // Present in the vocabulary, median 0 across the priors — the shape a
      // mechanism makes on the night it starts working.
      ...times(11, reread),
      // Seen in every prior round, three at a time, now eighteen.
      ...times(18, tagWrite),
    ]);
    const out = traceNovelty([...priors, newest]);
    expect(out.sinceBuild.map((s: { sig: string }) => s.sig).join()).toContain("re-reading the slide's shapes");
    expect(out.spikes.map((s: { sig: string }) => s.sig).join()).toContain("writing the chart's config tag");
    // Neither may appear in the other's bucket — that is the whole point.
    expect(out.spikes.map((s: { sig: string }) => s.sig).join()).not.toContain("re-reading");
    expect(out.sinceBuild.map((s: { sig: string }) => s.sig).join()).not.toContain("config tag");
    expect(out.spikes[0].median).toBe(3);
  });

  it("says nothing at all about a round that repeated the archive", () => {
    // The quiet case has to be assertable, or "nothing new" is just the report
    // failing to run and nobody can tell the difference.
    const out = traceNovelty([...quietPriors(8), round("same", times(3, entry("draw", "batch issued")))]);
    expect(out.novel).toHaveLength(0);
    expect(out.sinceBuild).toHaveLength(0);
    expect(out.spikes).toHaveLength(0);
    expect(out.vocabulary).toBe(1);
  });

  it("will not build a baseline out of too few rounds", () => {
    // With three priors a median is whichever one happened to be in the middle,
    // and every count looks like a spike against it. A fresh clone should get
    // silence, not an alarm resting on nothing.
    const out = traceNovelty([...quietPriors(3), round("new", times(40, entry("draw", "batch issued")))]);
    expect(out.spikes).toHaveLength(0);
    expect(out.sinceBuild).toHaveLength(0);
  });

  it("ignores a difference too small to be a difference", () => {
    const out = traceNovelty([...quietPriors(8), round("new", times(9, entry("draw", "batch issued")))]);
    expect(out.spikes).toHaveLength(0);
  });

  it("survives a round with no trace at all", () => {
    expect(() => traceNovelty([...quietPriors(6), { build: "x" }])).not.toThrow();
    expect(traceNovelty([]).novel).toEqual([]);
  });
});
