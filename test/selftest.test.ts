// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, makeShape, applyWebProfile, faults } from "./helpers/office-host";
import { CHART_TAG, requestStop, resetStop, isStopRequested } from "../src/render/powerpoint";
import { sampleConfig } from "../src/core/samples";
import { setTracing, traceLog } from "../src/core/trace";
import { runSelfTest, describeSelfTest, setSelfTestRasterizer, type ScenarioResult } from "../src/taskpane/selftest";

/**
 * The host self-test — nine paths the demo deck never touches.
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
    expect(results).toHaveLength(9);
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
    expect(results).toHaveLength(9);
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
    expect(results).toHaveLength(9);
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
      expect(starts).toHaveLength(9);
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
    expect(results, "a stopped battery dropped scenarios instead of reporting them").toHaveLength(9);
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
