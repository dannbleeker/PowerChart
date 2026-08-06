// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  installHost,
  makeSlide,
  makeShape,
  applyWebProfile,
  faults,
  trips,
  userClicksShape,
  selectionHandlerCount,
  type FakeSlide,
} from "./helpers/office-host";
import {
  CHART_TAG,
  requestStop,
  resetStop,
  isStopRequested,
  _setReadbackTimeoutForTest,
  listChartsInDeck,
  timeShapeRounds,
} from "../src/render/powerpoint";
import { sampleConfig } from "../src/core/samples";
import { setTracing, traceLog } from "../src/core/trace";
import {
  runSelfTest,
  describeSelfTest,
  selfTestNeedsAttention,
  setSelfTestRasterizer,
  SCENARIO_NAMES,
  ROUTINE_SCENARIO_NAMES,
  hostSeemsSick,
  readDegradation,
  _setClickWaitForTest,
  _setDegradeSizeForTest,
  type ScenarioResult,
} from "../src/taskpane/selftest";

/**
 * The host self-test — nine paths the demo deck never touches, plus one
 * experiment that a full run deliberately leaves out.
 *
 * The battery's own value is that it runs against a REAL PowerPoint, which
 * nothing here can do. What these cases pin is the property that makes it
 * worth clicking: that it comes back with a verdict for every scenario, that
 * the verdicts say what was observed, and above all that a scenario which
 * blows up does not take the other eight with it. A battery that stops at the
 * first error spends a whole real-host session to learn one thing — and the
 * scenarios after the failure are precisely the ones nobody has data for.
 */

afterEach(() => vi.unstubAllGlobals());

const byName = (rs: ScenarioResult[]) => Object.fromEntries(rs.map((r) => [r.name, r]));

describe("the host self-test battery", () => {
  it("returns a verdict for every scenario, in order", async () => {
    installHost([makeSlide("s1")]);
    const results = await runSelfTest("probe");
    expect(results).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
    expect(results.map((r) => r.name)).toEqual([
      "insert on top of an earlier run",
      "two slides claiming one slot",
      "edit a chart on the visible slide",
      "insert onto a slide that already has content",
      "same scale across the deck",
      "explode a degraded picture",
      // The three newest last, heaviest dead last — a crash in unproven code
      // must not cost the verdicts of scenarios that already work. Pinned,
      // because the ordering is a diagnostic property, not a detail.
      "edit the chart the user selected",
      "stop a run part-way",
      "the chart is actually visible",
    ]);
    for (const r of results) {
      expect(typeof r.ok).toBe("boolean");
      expect(r.detail, `${r.name} reported no detail`).toBeTruthy();
      expect(r.ms).toBeGreaterThanOrEqual(0);
    }
  });

  it("keeps going when a scenario throws, and records the throw", async () => {
    // The property the whole battery rests on. A host that wedges on the first
    // scenario used to mean five unknowns instead of one known and four
    // answers — on a session someone had to sit through to produce.
    installHost([makeSlide("s1")]);
    const boom = new Error("the host wedged");
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw boom;
      },
      SlideLayoutType: { blank: "blank" },
    });
    const results = await runSelfTest("probe");
    expect(results).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
    // Every one of them ran and reported; none is missing.
    for (const r of results) {
      expect(r.name).toBeTruthy();
      expect(r.detail).toBeTruthy();
    }
    const threw = results.filter((r) => r.detail.startsWith("threw:"));
    expect(threw.length).toBeGreaterThan(0);
    expect(threw[0].detail).toContain("the host wedged");
  });

  it("separates 'we did not check' from 'we checked and it is broken'", async () => {
    // A host below PowerPointApi 1.5 can insert nothing from base64 and can
    // hold no pictures. Reporting that as a failure would send a diagnosis
    // after a bug that is really a requirement-set gap.
    installHost([makeSlide("s1")], [], undefined, () => false);
    const results = await runSelfTest("probe");
    const skipped = results.filter((r) => r.skipped);
    expect(skipped.length).toBeGreaterThan(0);
    // A skip always says why. Two reasons are possible — the host cannot do
    // it, or an earlier scenario left nothing to work with — and at least one
    // here is the first, since this host advertises no requirement set at all.
    for (const s of skipped) expect(s.detail).toBeTruthy();
    expect(skipped.some((s) => /host/i.test(s.detail))).toBe(true);
    // A skip is never counted as a pass in the headline either.
    expect(describeSelfTest(results)).toContain("skipped");
  });

  /**
   * "Found nothing, therefore nothing was left behind" is an assertion a blind
   * scan satisfies for free.
   *
   * On a real host this scenario reported PASS off a scan that had just read
   * 1 of 8 slides — `unread=7 slides=8` is the line immediately before its own
   * verdict — while two other scenarios reported SKIPPED under exactly the same
   * blindness. An assertion that cannot fail is not a guard, and this one is
   * guarding the promise that Stop is non-destructive, which nothing else
   * checks.
   */
  it("does not pass 'nothing was left behind' on a scan that read nothing", async () => {
    installHost([makeSlide("s1")]);
    // Every deck scan comes back empty and SAYS so. `hollowReads` is per-read
    // and this scenario's scan is the last of many, so a large count is the
    // honest way to hold the whole run blind.
    faults.hollowReads = 500;
    try {
      // Targeted, so the host-sickness breaker cannot trip first and report
      // this scenario as "not reached" — which is also correct behaviour, and
      // would test the breaker instead of the assertion this case is about.
      const r = byName(await runSelfTest("probe", "stop a run part-way"))["stop a run part-way"];
      expect(r.ok, "passed on a scan that read nothing").toBe(false);
      expect(r.skipped, "reported a blind scan as a real verdict").toBe(true);
      expect(r.blind, "attributed the blind scan to something else").toBe(true);
      expect(r.detail).toMatch(/could not see the whole deck/i);
    } finally {
      faults.hollowReads = 0;
    }
  });

  /**
   * Stop asking a host that has stopped answering.
   *
   * Every real-host artefact this project owns ends the same way. The last one:
   * at 261.6s a scenario failed, at 261.8s the deck scan read 0 of 8 slides —
   * and the battery then ran Same Scale for 95 seconds (six charts, six
   * consecutive 5010 failures, two 90-second waits), a picture round-trip, and
   * two more scenarios, until PowerPoint killed the tab at ~396s. Roughly 130
   * seconds and 150 shapes were issued AFTER three distinct signals that the
   * host was already gone, and the tab took the remaining verdicts with it.
   */
  it("stops the run once the host has failed three scenarios in a row", async () => {
    installHost([makeSlide("s1")]);
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("the host wedged");
      },
      SlideLayoutType: { blank: "blank" },
    });
    const results = await runSelfTest("probe");
    // Still a verdict for every scenario — abandoning the run must not lose the
    // report, which is the only thing a run this bad produces.
    expect(results).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
    const abandoned = results.filter((r) => r.detail.includes("in a row"));
    expect(abandoned.length, "kept asking a host that had stopped answering").toBeGreaterThan(0);
    // And it gave the host a fair number of tries first: three real attempts
    // before it gives up, or a single flaky scenario would end every run.
    const attempted = results.filter((r) => !r.detail.includes("in a row"));
    expect(attempted.length, "gave up too early to have found anything").toBeGreaterThanOrEqual(3);
    // Never green, and never filed as a capability gap.
    expect(selfTestNeedsAttention(results)).toBe(true);
  });

  /**
   * The three verdicts of a real run, in the order it produced them.
   *
   * PowerPoint on the web, 2026-08-05. Two scenarios hit the 45-second draw
   * deadline and said "did not respond"; the third hit the SAME deadline after
   * 49.8 seconds and reported "the host stopped answering selection calls after
   * a programmatic select — known web-host limitation". The breaker matched
   * prose, so the third reset the counter to zero one short of tripping. The
   * battery then ran two more scenarios on a host that had been dead for three
   * and the tab died, taking their verdicts with it.
   *
   * Every one of these is a verbatim `detail` from that run.
   */
  it("counts a scenario that waited out its deadline, however it words the verdict", () => {
    const timedOutAndSaidSo: ScenarioResult = {
      name: "same scale across the deck",
      ok: false,
      detail:
        "threw: PowerPoint did not respond while drawing shapes 1-10 of 24 (45s) | at=redrawing the chart's shapes",
      ms: 48954,
    };
    const timedOutAndSaidSomethingElse: ScenarioResult = {
      name: "edit the chart the user selected",
      ok: false,
      skipped: true,
      detail:
        "the host stopped answering selection calls after a programmatic select — known web-host limitation, " +
        "same family as office-js#3083 / #3698; the pane's own Edit-it path is unaffected",
      ms: 49840,
    };
    // The rule this replaced, verbatim, so the guard proves the DIFFERENCE
    // rather than merely agreeing with itself. Without this the test would pass
    // against any rule that happens to return true, including the broken one.
    const oldRule = (r: ScenarioResult) =>
      r.blind === true || r.detail.startsWith("threw:") || /did not respond|gave up/i.test(r.detail);
    expect(oldRule(timedOutAndSaidSo), "the old rule caught this one — that part was never wrong").toBe(true);
    expect(
      oldRule(timedOutAndSaidSomethingElse),
      "the old rule already handled this verdict, so this guard proves nothing",
    ).toBe(false);

    // The one the old rule caught, and it still counts.
    expect(hostSeemsSick(timedOutAndSaidSo, true)).toBe(true);
    // The one it did not. A skip, no "threw:", and its wording matches neither
    // phrase the pattern looked for — but a deadline fired inside it, which is
    // the only thing being asked.
    expect(
      hostSeemsSick(timedOutAndSaidSomethingElse, true),
      "reset the breaker on a scenario that had just waited out the full budget",
    ).toBe(true);
    // And the rule stays narrow: without a deadline that same verdict is an
    // ordinary capability skip and says nothing about the host's health.
    expect(hostSeemsSick(timedOutAndSaidSomethingElse, false)).toBe(false);
    // A plain failure is the battery working, not a sick host.
    expect(hostSeemsSick({ name: "x", ok: false, detail: "the chart was gone", ms: 1 }, false)).toBe(false);
    // A blind deck scan still counts on its own.
    expect(hostSeemsSick({ name: "x", ok: false, blind: true, detail: "could not see the deck", ms: 1 }, false)).toBe(
      true,
    );
  });

  it("counts a bounded wait that hit its deadline, as a fact rather than a phrase", async () => {
    // The fact the rule above rests on. One deadline fired, one counted — if
    // this stops being true the breaker goes quiet again and nothing else says
    // so.
    const { deadlinesFired, _resetDeadlinesFiredForTest, slideCount, _setReadbackTimeoutForTest } =
      await import("../src/render/powerpoint");
    installHost([makeSlide("s1")]);
    _resetDeadlinesFiredForTest();
    _setReadbackTimeoutForTest(20);
    faults.wedgeAfterSyncs = 0;
    try {
      expect(deadlinesFired).toBe(0);
      // The rejection is the point; what is being counted is that it happened.
      await slideCount().catch(() => undefined);
      const after = (await import("../src/render/powerpoint")).deadlinesFired;
      expect(after, "a deadline fired and nothing counted it").toBeGreaterThan(0);
    } finally {
      faults.wedgeAfterSyncs = null;
      _setReadbackTimeoutForTest(90_000);
      _resetDeadlinesFiredForTest();
    }
  });

  it("counts a blind scan apart from a host that lacks the API", async () => {
    // Both are skips and they are not the same finding. The summary called
    // every skip "host cannot run them", so a run in which the host refused
    // every deck read reported, in green, a list of capability gaps.
    const results: ScenarioResult[] = [
      { name: "alpha", ok: true, detail: "fine", ms: 1 },
      { name: "beta", ok: false, skipped: true, detail: "host has no picture fill", ms: 1 },
      { name: "gamma", ok: false, skipped: true, blind: true, detail: "the deck scan could not see", ms: 1 },
    ];
    const text = describeSelfTest(results);
    expect(text).toContain("1 skipped (host cannot run them)");
    expect(text).toMatch(/1 could not run — the deck scan went blind/);
    // And the run is not green: a deck nobody could read is something to look at.
    expect(selfTestNeedsAttention(results)).toBe(true);
    expect(selfTestNeedsAttention(results.slice(0, 2)), "a plain capability skip is not a problem").toBe(false);
  });

  it("says which scenarios failed, not just how many", async () => {
    const results: ScenarioResult[] = [
      { name: "alpha", ok: true, detail: "fine", ms: 1 },
      { name: "beta", ok: false, detail: "broke", ms: 1 },
      { name: "gamma", ok: false, skipped: true, detail: "host cannot", ms: 1 },
    ];
    const text = describeSelfTest(results);
    expect(text).toContain("1 of 2 scenarios passed");
    expect(text).toContain("beta");
    // The skipped one is neither a pass nor a named failure.
    expect(text).not.toContain("gamma");
    expect(text).toContain("1 skipped");
  });

  it("runs to completion against a host at its worst", async () => {
    // The point of pinning it to the hostile profile: whatever the verdicts
    // are, the battery must survive stale proxies, hollow reads and a refused
    // addGroup well enough to report them. A battery that throws here would
    // be unusable on the one host it exists for.
    installHost([makeSlide("s1")]);
    applyWebProfile();
    const results = await runSelfTest("probe");
    expect(results).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
    const named = byName(results);
    expect(named["insert on top of an earlier run"].detail).toBeTruthy();
  });
});

/**
 * The three scenarios added when `setSelectedShapes` turned out to exist.
 *
 * Each is asserted twice: that it PASSES on a well-behaved host, and that it
 * FAILS when the thing it exists to catch is actually broken. Only the second
 * half makes it a test. A scenario that cannot fail is a line in a report that
 * always says "ok", which is worse than no scenario at all — it is evidence,
 * and it is false.
 */
describe("the scenarios the selection API unlocked", () => {
  it("edits through the selection, and notices when the host selects the wrong shape", async () => {
    installHost([makeSlide("s1")]);
    const good = byName(await runSelfTest("probe"))["edit the chart the user selected"];
    expect(good.skipped, good.detail).toBeFalsy();
    expect(good.ok, good.detail).toBe(true);

    // A host that takes the call and selects something else. The pane would
    // then edit a chart the user did not click — the failure this scenario is
    // for, and one no shape count or tag read can see.
    //
    // It lands on the "did not read back as a PowerChart" branch rather than
    // the "read back a different chart" one, because the shape it wrongly
    // selects is a chart PART and parts carry no config tag. Both branches are
    // the same defect — the selection went to the wrong shape — and the message
    // is pinned so a later change cannot quietly turn this into a pass for an
    // unrelated reason.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    faults.selectionIgnoresIds = true;
    const bad = byName(await runSelfTest("probe"))["edit the chart the user selected"];
    faults.selectionIgnoresIds = false;
    expect(bad.ok, `reported ok against a host that selects the wrong shape: ${bad.detail}`).toBe(false);
    expect(bad.skipped, "reported as skipped rather than broken").toBeFalsy();
    expect(bad.detail).toMatch(/did not read back as a PowerChart|read back a different chart/);
  });

  it("calls a host that stops answering after a select a known limitation, not a failure", async () => {
    // The third of PowerPoint-on-the-web's selection bugs, and the one that
    // cost a whole real-host round: `setSelectedShapes` is GA at PowerPointApi
    // 1.5 and TAKES the call, then leaves the selection subsystem unable to
    // answer anything. Measured on build 55011a3 — `getSelectedShapes` ran out
    // a 90-second budget, and the `setSelectedSlides` behind it did too.
    //
    // Two properties are pinned here, and the second is the one that matters:
    //
    //   1. It reports SKIPPED, not FAILED. Nothing in the add-in is broken by
    //      this — the pane never selects a shape from code, it reads the
    //      selection the user made with a click — so a red line here sends the
    //      next diagnosis after our own code, which is what happened.
    //   2. It COMES BACK, in about a budget rather than in minutes. The
    //      scenario asks for ten seconds instead of the ninety it would
    //      otherwise inherit, which is the difference between a battery a
    //      person will run again and one they will not.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    faults.selectionWedgesHost = true;
    // The scenario's budget is capped by this one, so shortening it shortens
    // both — that cap is why a bounded wait is testable at all here.
    _setReadbackTimeoutForTest(200);
    const started = Date.now();
    let wedged: ScenarioResult;
    try {
      wedged = byName(await runSelfTest("probe", "edit the chart the user selected"))[
        "edit the chart the user selected"
      ];
    } finally {
      faults.selectionWedgesHost = false;
      _setReadbackTimeoutForTest(90_000);
    }
    expect(wedged.skipped, `reported a host limitation as our failure: ${wedged.detail}`).toBe(true);
    expect(wedged.detail).toContain("stopped answering selection calls");
    expect(wedged.detail).toMatch(/office-js#3083/);
    // Not the point of the test, but it is the reason the gate exists: at the
    // inherited budget this is three minutes of a person's evening.
    expect(Date.now() - started, "waited far past its own budget").toBeLessThan(5_000);
  });

  it("names the exact selection call the host stopped answering", async () => {
    // Two accounts, different culprits. This project measured
    // `setSelectedShapes([id])` being taken and everything after it going
    // silent; office-js#3698 says it is `setSelectedShapes([])` whose promise
    // never resolves. Another blind round would produce the same ambiguity at
    // the same cost, so the ladder climbs from the least invasive call to the
    // most and STOPS at the first silence — because after a wedge every call is
    // silent, and four timeouts name nothing.
    //
    // The fake wedges on the non-empty select, so that is the answer the report
    // must reach. Against a fake that wedged on the empty one it would have to
    // reach the other, which is the whole point of asking rather than assuming.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    setTracing(true);
    faults.selectionWedgesHost = true;
    _setReadbackTimeoutForTest(200);
    let r: ScenarioResult;
    try {
      r = byName(await runSelfTest("probe", "which selection call wedges the host"))[
        "which selection call wedges the host"
      ];
    } finally {
      faults.selectionWedgesHost = false;
      _setReadbackTimeoutForTest(90_000);
    }
    // An experiment, not an assertion: it ran, so it reports ok whatever it found.
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail, "the ladder did not name the call that went silent").toContain("SILENT at");
    // The rung AFTER the wedging call is where silence shows up — the write
    // itself is taken. Naming the read alone would be the old, useless answer;
    // naming the rung before it is what makes this an attribution.
    expect(r.detail).toContain("after selecting a shape");
    expect(r.detail, "did not say what the host last answered").toContain("the rung before it");

    // And it STOPPED there — asserted against the rungs the ladder actually
    // traced, not against the sentence it wrote. The first version of this
    // check looked for the absence of a later rung's NAME in the summary, and
    // that summary only ever mentions two rungs, so it passed just as happily
    // against a ladder that climbed all the way to the top. It was decoration.
    const rungs = traceLog()
      .entries.filter((e) => e.message === "ladder rung")
      .map((e) => e.data as { step: string; outcome: string });
    setTracing(false);
    expect(rungs.length, "the ladder recorded nothing").toBeGreaterThan(0);
    expect(
      rungs.filter((x) => x.outcome === "silent"),
      "climbed past the first silence",
    ).toHaveLength(1);
    expect(rungs.at(-1)?.outcome, "kept going after the host went quiet").toBe("silent");
  });

  it("says so plainly when nothing wedges at all", async () => {
    // The other half. A host that answers every rung is a real result — the
    // wedge is gone, or is not on this host — and a report that could only ever
    // say "SILENT at …" would be evidence that cannot be contradicted, which is
    // not evidence.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    const r = byName(await runSelfTest("probe", "which selection call wedges the host"))[
      "which selection call wedges the host"
    ];
    expect(r.ok, r.detail).toBe(true);
    expect(r.detail).toContain("nothing wedged");
    expect(r.detail).not.toContain("SILENT");
  });

  it("reads the chart back from a REAL click, and takes its listener away after", async () => {
    // The one path a user travels that nothing can script. `setSelectedShapes`
    // is Office.js selecting a shape; a human clicking one is the same call in
    // theory and demonstrably not in practice on the web. So the battery asks
    // for a click and listens on `DocumentSelectionChanged` — a Common API
    // event that does not go through the wedging subsystem.
    //
    // What is asserted here is the half CI can honestly speak to: the
    // listen → click → read chain, and that the listener is taken back off.
    // Whether the EDIT that follows succeeds is `editViaSelection`'s ground,
    // covered twice over there; duplicating it here would only assert that the
    // fake host edits, which is not what this scenario is for.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    const run = runSelfTest("probe", "edit the chart YOU click");
    // Click only once the scenario is actually listening — the ordering that
    // exercises the handler rather than a lucky first read.
    await vi.waitFor(() => expect(selectionHandlerCount()).toBeGreaterThan(0));
    const chart = (await listChartsInDeck()).charts[0];
    expect(chart, "no probe chart for the click to land on").toBeTruthy();
    userClicksShape(chart.target.shapeId);
    const r = byName(await run)["edit the chart YOU click"];

    expect(r.skipped, `a click arrived and it still reported skipped: ${r.detail}`).toBeFalsy();
    expect(r.detail, "did not read the clicked chart back").toMatch(/read ".*" back from a real click/);
    // And it took its listener back off. A battery that leaked one handler per
    // run would leave dead runs answering the user's clicks.
    expect(selectionHandlerCount(), "left a selection handler behind").toBe(0);
  });

  it("does not blame the user for a click the host would not describe", async () => {
    // Two different findings that used to read identically. If the selection
    // read goes silent after a click, the person did their part and the HOST
    // did not — reporting "nobody clicked" tells them they failed at the one
    // thing they can see they did, which is how a report stops being believed.
    //
    // It also used to be reachable by timing alone: the read inherited the
    // 30-second CLICK budget, so a click at 29s started a read with 30s of rope
    // and the deadline cut it off one second later. Two unrelated quantities
    // sharing a name.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    _setClickWaitForTest(400);
    let r: ScenarioResult;
    try {
      const run = runSelfTest("probe", "edit the chart YOU click");
      // The probe charts only exist once the scenario is listening, and the
      // fault has to be armed AFTER reading them or it breaks the deck scan
      // rather than the selection read this test is about.
      await vi.waitFor(() => expect(selectionHandlerCount()).toBeGreaterThan(0));
      const chart = (await listChartsInDeck()).charts[0];
      // A click, on a host that then refuses the shape-collection read. This
      // is the "e.load is not a function" shape a real web host produced.
      faults.selectionReadThrows = true;
      userClicksShape(chart.target.shapeId);
      r = byName(await run)["edit the chart YOU click"];
    } finally {
      faults.selectionReadThrows = false;
      _setClickWaitForTest(30_000);
    }
    expect(r.skipped, "a host that went silent was reported as a pass or a failure").toBe(true);
    expect(r.detail, "blamed the user for the host's silence").not.toContain("nobody clicked");
    expect(r.detail).toContain("would not say what was selected");
    expect(selectionHandlerCount(), "left a selection handler behind").toBe(0);
  });

  it("skips rather than fails when nobody clicks", async () => {
    // Nobody clicked is "we did not check", never "it is broken". Reporting a
    // person's absence as a red line is how a battery teaches its reader to
    // ignore red lines.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    _setClickWaitForTest(60);
    let r: ScenarioResult;
    try {
      r = byName(await runSelfTest("probe", "edit the chart YOU click"))["edit the chart YOU click"];
    } finally {
      _setClickWaitForTest(30_000);
    }
    expect(r.ok, r.detail).toBe(false);
    expect(r.skipped, "a missing human was reported as a failure").toBe(true);
    expect(r.detail).toContain("nobody clicked");
    expect(selectionHandlerCount(), "left a selection handler behind after timing out").toBe(0);
  });

  it("reports a stop honestly, and will not call a leftover chart a clean stop", async () => {
    installHost([makeSlide("s1")]);
    const good = byName(await runSelfTest("probe"))["stop a run part-way"];
    expect(good.ok, good.detail).toBe(true);
    expect(good.detail).toContain("stopped at a batch boundary");

    // A chart already on the slide under the name a stopped insert would use.
    // The scenario must read that as "the stop left something behind" — if it
    // only checked that the throw was a stop, a host that stopped AND littered
    // would report a clean pass.
    vi.unstubAllGlobals();
    const slide = makeSlide("s1");
    const litter = makeShape("geometric", "rectangle", { left: 0, top: 0, width: 10, height: 10 });
    litter.name = "PowerChart";
    litter.tagStore.set(CHART_TAG, JSON.stringify({ ...sampleConfig("clustered"), title: "probe stopped" }));
    slide.created.push(litter);
    installHost([slide]);
    const bad = byName(await runSelfTest("probe"))["stop a run part-way"];
    expect(bad.ok, `called a stop clean with a chart left behind: ${bad.detail}`).toBe(false);
    expect(bad.detail).toContain("left a re-editable chart");
  });

  it("sees the chart on the slide, and fails when the host renders the same bytes regardless", async () => {
    installHost([makeSlide("s1")]);
    const good = byName(await runSelfTest("probe"))["the chart is actually visible"];
    expect(good.skipped, good.detail).toBeFalsy();
    expect(good.ok, good.detail).toBe(true);

    // A host whose rasteriser answers the same thing for an empty slide and a
    // slide with a chart on it. The comparison is then worthless, and saying so
    // is the only honest verdict — this is the one scenario whose whole value
    // is the difference between two images.
    vi.unstubAllGlobals();
    installHost([makeSlide("s1")]);
    faults.constantSlideImage = true;
    const bad = byName(await runSelfTest("probe"))["the chart is actually visible"];
    faults.constantSlideImage = false;
    expect(bad.ok, `passed while every render was identical: ${bad.detail}`).toBe(false);
    expect(bad.detail).toContain("nothing is visible");
  });

  /**
   * The failure this scenario actually hit in a real PowerPoint, on 2026-08-04.
   *
   * It got as far as "rasterising the empty slide" and the host answered
   * `GeneralException`, `errorLocation: SlideCollection.getItem`, `statement:
   * var slide = slides.getItem(...); slide.getImageAsBase64(...)`. The slide
   * was one `addScratchSlide` had just made and whose liveness check had just
   * passed — through the very same proxy, one sync earlier. Resolving a proxy
   * is what makes Office.js rewrite its object path to `getItem(id)`, and a new
   * slide's id does not round-trip through `getItem` on the web: the rule
   * `SlideThunk` is built on, in a function that held one handle across two
   * syncs.
   */
  it("rasterises a slide the host will only name once", async () => {
    installHost([makeSlide("s1")]);
    const r = byName(await runSelfTest("probe"))["the chart is actually visible"];
    expect(r.skipped, `gave up on a host that answers by-id handles once: ${r.detail}`).toBeFalsy();
    expect(r.ok, r.detail).toBe(true);
  });

  /**
   * The healthy-host half of the blind-scan change, and ONLY that.
   *
   * A real host reported this scenario FAILED — *"the chart already on the
   * slide is no longer re-editable"* — on a build whose insert path had not
   * changed since the run where it passed. The scenario checked its first deck
   * scan for blindness and then made its most alarming claim off a second one
   * it never checked, so "your chart is gone" and "I could not look" came out
   * as the same sentence. It now skips when either scan is blind.
   *
   * WHAT THIS TEST DOES NOT DO: fail against the pre-fix code. Reproducing the
   * real sequence needs the first scan to answer and a later one to go short,
   * and `hollowReads` blinds from the first read — so the scenario's own
   * opening check catches it and the new rung is never reached. `hollowReadsAfter`
   * was added to model a host that degrades mid-run, which is what the real one
   * did, but landing it on exactly the right read means guessing a scan count
   * three scenarios deep — the fragile magic number CLAUDE.md warns about, and
   * a guard tuned until it goes red is not evidence.
   *
   * So this pins the half that IS checkable: the change must not cost the
   * healthy-host case its pass. The blind case is argued from the host's own
   * log, and is honestly unguarded until something can drive that window.
   */
  it("still passes on a host whose deck scans answer", async () => {
    installHost([makeSlide("s1")]);
    const r = byName(await runSelfTest("probe", "insert onto a slide that already has content"));
    const verdict = r["insert onto a slide that already has content"];
    expect(verdict.ok, `the healthy-host case stopped working: ${verdict.detail}`).toBe(true);
    expect(verdict.skipped, "skipped a scenario whose scans were fine").toBeFalsy();
  });

  /**
   * A rasteriser that answers with nothing has to SAY so.
   *
   * `slideImageBase64` has four ways of returning undefined — no 1.8, the
   * slide would not resolve, the call threw, the call came back empty — and
   * one `catch` covering all of them. A real run showed how little that is
   * worth: the visibility scenario went from "rasterising the empty slide"
   * straight to "removing the scratch slide", with no error line and no
   * drawing step, and the reason existed nowhere. The round before, the same
   * call at least reported `GeneralException`; a fresh slide handle stopped it
   * throwing without making it work, and traded a loud wrong answer for a
   * silent one. Both are failures. Only one can be diagnosed.
   */
  it("says why it could not rasterise, instead of quietly reporting nothing", async () => {
    installHost([makeSlide("s1")]);
    setTracing(true);
    // The host takes the call and hands back an empty image — the quiet case,
    // which is the one that had no line at all.
    faults.constantSlideImage = false;
    faults.emptySlideImage = true;
    try {
      const r = byName(await runSelfTest("probe", "the chart is actually visible"))["the chart is actually visible"];
      expect(r.ok, "passed with no image to compare").toBe(false);
      const said = traceLog()
        .entries.filter((e) => e.scope === "host")
        .map((e) => e.message);
      expect(said, `nothing in the log says why: ${said.join(" | ")}`).toContain(
        "the host took the rasterise and returned nothing",
      );
    } finally {
      faults.emptySlideImage = false;
    }
  });

  /**
   * A real host died inside this scenario and the log's last line was the
   * scenario announcing itself — five host calls, and nothing to say which one
   * was outstanding. The scenario-level announcement was itself added for
   * exactly this reason one round earlier; it turned "somewhere in the battery"
   * into "somewhere in this scenario", and stopped there.
   *
   * The rasteriser is held open rather than made to fail, because a call that
   * throws leaves a verdict and a call that never returns leaves only the log.
   * Only the second one is the case this line exists for.
   */
  it("names the host call it is on, so a scenario that never ends is not a mystery", async () => {
    setTracing(true);
    let release!: () => void;
    try {
      // After installHost — it resets every fault, so arming one before it is
      // arming nothing.
      installHost([makeSlide("s1")]);
      faults.slideImageGate = new Promise<void>((r) => (release = r));
      const run = runSelfTest("probe", "the chart is actually visible");
      await vi.waitFor(() => {
        const last = traceLog()
          .entries.filter((e) => e.scope === "selftest")
          .at(-1);
        expect(last?.message).toBe("visibility step");
        expect(last?.data?.what).toBe("rasterising the empty slide");
      });
      release();
      await run;
      // And every call is named, in the order the scenario makes them.
      const steps = traceLog()
        .entries.filter((e) => e.message === "visibility step")
        .map((e) => e.data?.what);
      expect(steps).toEqual([
        "adding a scratch slide",
        "rasterising the empty slide",
        "drawing the chart",
        "rasterising the slide with the chart",
        "removing the scratch slide",
      ]);
    } finally {
      release?.();
      faults.slideImageGate = null;
      setTracing(false);
    }
  });

  it("says so when it cannot take its own scratch slide back", async () => {
    // The visibility scenario is the only one that borrows a slide and returns
    // it, so it is the only one that can leave litter. `deleteSlideById` is
    // best-effort and answers false rather than throwing, and the first version
    // of this scenario discarded that answer in a `finally` — so a host that
    // refused the delete left a tagged chart in the deck under a verdict that
    // read perfectly clean.
    installHost([makeSlide("s1")]);
    faults.refuseSlideDelete = true;
    const r = byName(await runSelfTest("probe"))["the chart is actually visible"];
    faults.refuseSlideDelete = false;
    expect(r.detail, "left a slide behind and said nothing").toMatch(/could not be removed/i);
  });

  it("runs one scenario plus the two inserts it needs, when asked for one", async () => {
    // A full round costs minutes and leaves slides behind: the first run that
    // produced a usable trace took 496 seconds to reach scenario seven and
    // wedged there, and the scenario before it left four charts on the deck as
    // loose piles of shapes. Iterating on the seventh by running the six in
    // front of it is not iteration.
    //
    // The first two run regardless, because every other scenario needs the
    // probe charts they insert — a targeted run that found none would report
    // SKIPPED, which is honest and useless.
    installHost([makeSlide("s1")]);
    const results = await runSelfTest("probe", "the chart is actually visible");
    expect(results.map((r) => r.name)).toEqual([
      "insert on top of an earlier run",
      "two slides claiming one slot",
      "the chart is actually visible",
    ]);
    const picked = results[2];
    expect(picked.skipped, picked.detail).toBeFalsy();
    expect(picked.ok, picked.detail).toBe(true);
  });

  it("offers exactly the scenarios it can run, in run order", async () => {
    // The picker is filled from this list. A hand-kept copy beside it would be
    // one rename away from offering a scenario that does not exist.
    installHost([makeSlide("s1")]);
    const results = await runSelfTest("probe");
    // A full run is the routine list exactly, in order.
    expect(results.map((r) => r.name)).toEqual([...ROUTINE_SCENARIO_NAMES]);
    // The picker offers those PLUS the ones a full run deliberately leaves out,
    // because being pickable is the only way such a scenario can ever run.
    expect(SCENARIO_NAMES.slice(0, ROUTINE_SCENARIO_NAMES.length)).toEqual([...ROUTINE_SCENARIO_NAMES]);
    for (const name of ROUTINE_SCENARIO_NAMES) expect(SCENARIO_NAMES).toContain(name);
    expect(SCENARIO_NAMES.length).toBeGreaterThan(ROUTINE_SCENARIO_NAMES.length);
  });

  it("runs everything when nothing is picked", async () => {
    installHost([makeSlide("s1")]);
    expect(await runSelfTest("probe", "")).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
  });

  it("announces a scenario BEFORE running it, so a crash names the right one", async () => {
    // Every verdict this battery emits is past tense. A run that dies mid
    // scenario therefore leaves the PREVIOUS scenario's line as its last word
    // — off by one, and pointing at the one thing that demonstrably worked.
    // That is not hypothetical: two real-host rounds have been diagnosed from
    // a screenshot of the pane rather than a log, because the log only exists
    // once a run ends and neither run ended.
    setTracing(true);
    try {
      installHost([makeSlide("s1")]);
      await runSelfTest("probe");
      const entries = traceLog().entries.filter((e) => e.scope === "selftest");
      const first = entries[0];
      expect(first?.message, "the first thing said about a scenario was its verdict").toBe("scenario starting");
      expect(first?.data?.name).toBe("insert on top of an earlier run");
      // And one announcement per scenario, each before its own verdict.
      const starts = entries.filter((e) => e.message === "scenario starting");
      expect(starts).toHaveLength(ROUTINE_SCENARIO_NAMES.length);
      for (const s of starts) expect(s.data?.name, "an announcement with no name").toBeTruthy();
    } finally {
      setTracing(false);
    }
  });

  it("can be stopped, and says which scenarios were never reached", async () => {
    // The battery had no stop check of its own. A user who pressed Stop got a
    // button that said "Stopping…" and a run that carried on to the next
    // scenario regardless — observed on a real host at 1819 seconds, where the
    // only way out was closing the tab, which is also how that run's log was
    // lost.
    //
    // Not-reached is reported as SKIPPED with a reason, never dropped: a report
    // missing its last four lines looks like a battery that crashed, and that
    // is a different diagnosis from one that was stopped.
    installHost([makeSlide("s1")]);
    requestStop();
    let results: ScenarioResult[];
    try {
      results = await runSelfTest("probe");
    } finally {
      resetStop();
    }
    expect(results, "a stopped battery dropped scenarios instead of reporting them").toHaveLength(
      ROUTINE_SCENARIO_NAMES.length,
    );
    expect(
      results.every((r) => r.skipped),
      "a stopped battery ran a scenario anyway",
    ).toBe(true);
    expect(results[0].detail).toMatch(/stopped/i);
  });

  it("does not let the stop scenario clear a stop the USER asked for", async () => {
    // `stopPartWay` is the only code in the add-in that calls resetStop(). If
    // it runs while the user's own stop is pending, its finally clears the
    // flag — and the battery resumes after the user cancelled it. The guard
    // above makes this unreachable; this pins it anyway, because "unreachable"
    // is how the flag came to be clobbered in the first place.
    installHost([makeSlide("s1")]);
    requestStop();
    try {
      await runSelfTest("probe");
      expect(isStopRequested(), "the stop scenario cleared the user's stop").toBe(true);
    } finally {
      resetStop();
    }
  });

  it("skips rather than fails on a host below PowerPointApi 1.5", async () => {
    // "We could not check" and "we checked and it is broken" are different
    // answers, and reporting the first as the second is how a requirement-set
    // gap gets diagnosed as a bug.
    installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.5");
    const r = byName(await runSelfTest("probe"))["edit the chart the user selected"];
    expect(r.skipped, r.detail).toBe(true);
    expect(r.detail).toContain("1.5");
  });
});

describe("the self-test against a host that can take a generated deck", () => {
  it("actually runs the insert scenarios, and they pass on a well-behaved host", async () => {
    // Until the fake could take a .pptx these two always reported "skipped",
    // so the battery's own insert logic had never executed anywhere. A skip is
    // an honest answer to give a user and a useless one to give a test suite.
    installHost([makeSlide("s1")]);
    const named = byName(await runSelfTest("probe"));
    const twice = named["insert on top of an earlier run"];
    expect(twice.skipped, twice.detail).toBeFalsy();
    expect(twice.ok, twice.detail).toBe(true);
    expect(twice.detail).toContain("deck grew by 4");

    const dup = named["two slides claiming one slot"];
    expect(dup.skipped, dup.detail).toBeFalsy();
    expect(dup.ok, dup.detail).toBe(true);
  });

  it("notices when the host quietly drops half the deck", async () => {
    // The negative control. A run that inserts four slides and gets two must
    // not read as a pass — that is the exact failure the scenario exists to
    // catch, and it is invisible to a check that only asks whether the call
    // threw. A real 12-item file insert came back "11 of 12 complete · 1 lost"
    // with nothing raised at all.
    installHost([makeSlide("s1")]);
    expect(byName(await runSelfTest("probe"))["insert on top of an earlier run"].ok).toBe(true);

    installHost([makeSlide("s1")]);
    faults.swallowDecks = 1; // the host takes the first deck and lands nothing
    const after = byName(await runSelfTest("probe"))["insert on top of an earlier run"];
    expect(after.ok).toBe(false);
    expect(after.detail).toContain("deck grew by 2");
  });
});

describe("scenarios that must not be able to pass without proving anything", () => {
  const named = (rs: ScenarioResult[]) => Object.fromEntries(rs.map((r) => [r.name, r]));

  it("skips the picture scenario rather than faking one when there is no rasteriser", async () => {
    // `render: "image"` on a config does NOT make a picture — the renderer
    // takes the picture path only when handed `pictureBase64`. An earlier
    // version passed `undefined`, quietly performed two ordinary shape
    // updates, and reported that the picture round-trip worked. A scenario
    // that cannot fail is worse than one that is missing: it reads as
    // evidence. With no rasteriser the honest answer is "did not check".
    setSelfTestRasterizer(undefined as unknown as (s: unknown) => Promise<string | undefined>);
    installHost([makeSlide("s1")]);
    const explode = named(await runSelfTest("probe"))["explode a degraded picture"];
    expect(explode.skipped, explode.detail).toBe(true);
    expect(explode.detail).toMatch(/rasteris/i);
  });

  it("hands the host a real picture when there is one", async () => {
    // The deck, not one slide: the insert scenarios append their own slides,
    // so the chart this one collapses is never on the slide we started with.
    const deck = [makeSlide("s1")];
    installHost(deck);
    let asked = 0;
    setSelfTestRasterizer(async () => {
      asked++;
      return "data:image/png;base64,UE5H";
    });
    const explode = named(await runSelfTest("probe"))["explode a degraded picture"];
    expect(asked, "the scenario never asked for a picture").toBeGreaterThan(0);
    expect(explode.skipped, explode.detail).toBeFalsy();
    // The picture really reached the host, rather than the renderer drawing
    // shapes and the scenario calling that a picture.
    const pictured = deck.flatMap((s) => s.created).filter((s) => s.imageBase64);
    expect(pictured.length, "no picture fill ever reached the host").toBeGreaterThan(0);
  });

  it("skips the rescale rather than comparing against -Infinity", async () => {
    // `Math.max()` of nothing is -Infinity, and JSON.stringify turns that into
    // `null` — so an all-blank data set wrote {"scale":{"max":null}} and then
    // compared it against -Infinity, which never matches. The scenario
    // reported a failure that was its own arithmetic.
    installHost([makeSlide("s1")]);
    setSelfTestRasterizer(async () => undefined);
    const scale = named(await runSelfTest("probe"))["same scale across the deck"];
    expect(scale.detail).toBeTruthy();
    expect(scale.detail).not.toContain("Infinity");
    expect(scale.detail).not.toContain("null");
  });
});

describe("the scenario for a slide that already has content", () => {
  const onto = (rs: ScenarioResult[]) => rs.find((r) => r.name === "insert onto a slide that already has content")!;

  it("passes on a host that behaves, and on one that does not", async () => {
    // The everyday action — "insert a chart on the slide I am looking at" —
    // and until this scenario existed, the only path with no real-host
    // coverage at all. Every other scenario, and the whole demo deck, works on
    // slides the run added BLANK.
    for (const hostile of [false, true]) {
      installHost([makeSlide("s1")]);
      if (hostile) applyWebProfile();
      setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
      const r = onto(await runSelfTest("probe"));
      expect(r.skipped, r.detail).toBeFalsy();
      expect(r.ok, `${hostile ? "hostile" : "clean"}: ${r.detail}`).toBe(true);
      expect(r.detail).toContain("re-editable");
    }
  }, 60_000);

  it("reports a failure when the charts land un-re-editable", async () => {
    // The negative control, and the whole point of the scenario. A host with
    // no tag support (below PowerPointApi 1.3) cannot make a chart
    // re-editable, so the scenario must SAY so rather than count the shapes
    // and call it a pass.
    //
    // Proven against the real defect too: reverting #243 — the fresh-proxy
    // re-fetch on `insertSceneIntoSlide` — makes this scenario report
    // "0 of 2 new charts are re-editable" under `strictGroup` and under the
    // full web profile, while still passing on a clean host. That is exactly
    // the shape of the bug: invisible except on a host that refuses stale
    // proxies, which is the web.
    installHost([makeSlide("s1")], [], undefined, (v) => v !== "1.3");
    setSelfTestRasterizer(async () => "data:image/png;base64,UE5H");
    const r = onto(await runSelfTest("probe"));
    expect(r.ok, r.detail).toBe(false);
    expect(r.detail).toMatch(/re-editable/);
  }, 60_000);
});

describe("the experiment that asks what makes a long run slow down", () => {
  /**
   * Every real-host artefact this project owns degrades, and not one of them
   * can say WHY: a run that gets slower grows its deck, ages its tab and holds
   * a request context open all at the same time. The repo's own rule for that
   * situation is to ask a question that separates the answers rather than to
   * reason about which is likelier — this is the question, and these are the
   * cases that check it actually separates them.
   *
   * `faults.syncCostMs` is what makes that checkable at all. Given a cost that
   * follows the syncs within a context it stands for proxy accumulation; given
   * one that follows the syncs since the tab opened it stands for a host
   * slowing down. The verdict has to name a different suspect for each, and a
   * verdict that cannot is worse than none: it would send a session's work at
   * the wrong thing.
   */
  const series = async (id: string, oneContext: boolean, rounds = 9) =>
    (await timeShapeRounds(id, { rounds, perRound: 2, oneContext, label: "t", budgetMs: 30_000 })).rounds.map(
      (r) => r.ms,
    );

  it("uses one context for the long arm and one per round for the other", async () => {
    // The structural half, and deterministic: no clock, no threshold. If the
    // two arms were ever wired the same way — both long, both fresh, or
    // swapped — every timing case below would still pass, comparing a series
    // against itself and reporting whatever the noise said.
    installHost([makeSlide("s1")]);
    const before = trips.contexts;
    await series("s1", true);
    const long = trips.contexts - before;
    const mid = trips.contexts;
    await series("s1", false);
    expect(long).toBe(1);
    expect(trips.contexts - mid).toBe(9);
  });

  it("blames the context when only a held context gets slower", async () => {
    installHost([makeSlide("s1")]);
    // Flat 10ms plus 8ms per sync already spent IN THIS CONTEXT. The fresh arm
    // never sees the second term, so its cost is constant by construction.
    faults.syncCostMs = ({ syncsInContext }) => 10 + syncsInContext * 8;
    const one = await series("s1", true);
    const fresh = await series("s1", false);
    const v = readDegradation(one, fresh);
    expect(v.freshContext, `fresh: ${fresh.join(",")}`).toBeLessThan(0.5);
    expect(v.oneContext, `one: ${one.join(",")}`).toBeGreaterThan(0.5);
    expect(v.suspect).toBe("context");
    expect(v.headline).toContain("THE CONTEXT");
  });

  it("blames the host when both arms get slower together", async () => {
    installHost([makeSlide("s1")]);
    // The same shape of cost, hung on the counter that does NOT reset at a
    // context boundary. A context boundary buys nothing here, which is exactly
    // what a growing deck or an ageing tab would look like.
    //
    // Rebased on the syncs already spent, so the case does not depend on where
    // in the file it runs — `trips.syncs` is a whole-suite counter and an
    // absolute cost read off it would make this test's answer depend on the
    // tests before it.
    const base = trips.syncs;
    faults.syncCostMs = ({ syncsTotal }) => 10 + (syncsTotal - base) * 6;
    const one = await series("s1", true);
    const fresh = await series("s1", false);
    const v = readDegradation(one, fresh);
    // The load-bearing assertion: the FRESH arm grew. That is the one thing a
    // context-accumulation story cannot produce, and the one thing that tells
    // a reader shortening contexts will not help.
    expect(v.freshContext, `fresh: ${fresh.join(",")}`).toBeGreaterThan(0.5);
    expect(["host", "both"]).toContain(v.suspect);
    expect(v.headline).toContain("THE HOST");
  });

  it("does not read a host that slows LINEARLY as a context problem", async () => {
    // The case the first version of this got wrong, kept as its own test
    // because it is the commonest shape a real slowdown takes and the mistake
    // was invisible without it. Comparing each arm's tail to its OWN head, the
    // long arm read ×2.64 and the fresh arm ×1.47 for a slowdown neither arm
    // caused — the fresh arm only looks flatter because it starts from a higher
    // baseline. A confident "the context did it", about a host.
    //
    // Same fault as the case above; what this one pins is the NUMBER, not just
    // the label. Both arms are hit equally hard in milliseconds, so the two
    // growth figures must come out within a hair of each other.
    installHost([makeSlide("s1")]);
    const base = trips.syncs;
    faults.syncCostMs = ({ syncsTotal }) => 10 + (syncsTotal - base) * 6;
    const one = await series("s1", true);
    const fresh = await series("s1", false);
    const v = readDegradation(one, fresh);
    expect(v.suspect, `one: ${one.join(",")} | fresh: ${fresh.join(",")}`).toBe("host");
    expect(Math.abs(v.oneContext - v.freshContext), "the arms were hit equally and must read equally").toBeLessThan(
      0.35,
    );
  });

  it("blames nothing when the host is steady", async () => {
    installHost([makeSlide("s1")]);
    faults.syncCostMs = () => 6;
    const v = readDegradation(await series("s1", true), await series("s1", false));
    expect(v.suspect, "a steady host must read as steady").toBe("none");
  });

  it("says so rather than guessing when there are too few rounds to read", async () => {
    // Three rounds is not a curve. Reporting "no degradation" from it would be
    // the same mistake as reading a blind deck scan as an empty deck, which
    // this file already has a whole third result state for.
    installHost([makeSlide("s1")]);
    const v = readDegradation(await series("s1", true, 3), await series("s1", false, 3));
    expect(v.suspect).toBe("unknown");
    expect(v.headline).toContain("not enough rounds");
  });

  it("keeps the rounds it managed when the host stops answering", async () => {
    // A curve with four points is a finding; an exception is not. The long arm
    // is the one at risk, because its whole series sits under one deadline.
    installHost([makeSlide("s1")]);
    faults.wedgeAfterSyncs = 4;
    const got = await timeShapeRounds("s1", {
      rounds: 9,
      perRound: 2,
      oneContext: true,
      label: "t",
      budgetMs: 150,
    });
    expect(got.rounds).toHaveLength(4);
    expect(got.cutShort, "a series that lost five rounds must say so").toBeTruthy();
  });

  it("is offered by the picker, left out of a full run, and draws its arms on two slides", async () => {
    // The fake pushes `slides.add()` onto the array it was handed, so the deck
    // the experiment built its two scratch slides in is readable here — which
    // is the only way to check the arms went to different ones.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const deckShapeNames = () => deck.map((s) => s.created.map((c) => c.name ?? ""));
    _setDegradeSizeForTest(4, 2);
    try {
      expect(SCENARIO_NAMES).toContain("what makes a long run slow down");
      expect(ROUTINE_SCENARIO_NAMES).not.toContain("what makes a long run slow down");
      const r = (await runSelfTest("probe", "what makes a long run slow down")).find(
        (x) => x.name === "what makes a long run slow down",
      )!;
      expect(r.detail).toContain("one context:");
      expect(r.detail).toContain("fresh contexts:");
      // Two arms, two slides. Sharing one would make the second arm draw onto a
      // slide already holding the first arm's shapes, and a fat slide is a
      // fourth variable in an experiment built to have exactly one.
      const named = (frag: string) =>
        deckShapeNames().filter((names) => names.some((n) => n.includes(frag))).length;
      expect(named("one-context")).toBe(1);
      expect(named("fresh-context")).toBe(1);
    } finally {
      _setDegradeSizeForTest(8, 12);
    }
  }, 60_000);
});
