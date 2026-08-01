// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, applyWebProfile, faults } from "./helpers/office-host";
import { runSelfTest, describeSelfTest, setSelfTestRasterizer, type ScenarioResult } from "../src/taskpane/selftest";

/**
 * The host self-test — six paths the demo deck never touches.
 *
 * The battery's own value is that it runs against a REAL PowerPoint, which
 * nothing here can do. What these cases pin is the property that makes it
 * worth clicking: that it comes back with a verdict for every scenario, that
 * the verdicts say what was observed, and above all that a scenario which
 * blows up does not take the other five with it. A battery that stops at the
 * first error spends a whole real-host session to learn one thing — and the
 * scenarios after the failure are precisely the ones nobody has data for.
 */

afterEach(() => vi.unstubAllGlobals());

const byName = (rs: ScenarioResult[]) => Object.fromEntries(rs.map((r) => [r.name, r]));

describe("the host self-test battery", () => {
  it("returns a verdict for every scenario, in order", async () => {
    installHost([makeSlide("s1")]);
    const results = await runSelfTest("probe");
    expect(results).toHaveLength(6);
    expect(results.map((r) => r.name)).toEqual([
      "insert on top of an earlier run",
      "two slides claiming one slot",
      "edit a chart on the visible slide",
      "insert onto a slide that already has content",
      "same scale across the deck",
      "explode a degraded picture",
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
    expect(results).toHaveLength(6);
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
    expect(results).toHaveLength(6);
    const named = byName(results);
    expect(named["insert on top of an earlier run"].detail).toBeTruthy();
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
