// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, makeShape, applyWebProfile, faults, stallSyncOn } from "./helpers/office-host";
import {
  insertDemoDeck,
  insertSceneIntoSlide,
  updateChartInSlide,
  DEMO_SLOT_TAG,
  CHART_TAG,
  _setBatchTimeoutForTest,
} from "../src/render/powerpoint";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import type { ChartConfig, ChartKind } from "../src/core/types";

/**
 * The demo path against a host behaving as badly as PowerPoint on the web has
 * actually been observed to behave — every misbehaviour at once.
 *
 * Every bug this project has fixed in the last week was invisible to the local
 * suite and found by a human running the add-in in a real PowerPoint, saving
 * the deck, and uploading it: a stale shape proxy rejected by `getItem(id)`, a
 * shape collection reading back shorter than it is without throwing, a group
 * call refused outright, a sync that commits minutes after the add-in gave up
 * on it. Each cost a deploy, a run someone sat through, an upload and an
 * analysis — and each fix then added one more knob to the fake host so it
 * could never come back.
 *
 * Those knobs were only ever used one or two at a time, in the test that
 * introduced them, against an otherwise well-behaved host. Turned on together
 * they describe something no suite here had ever run against. That is what
 * this file does, and it is why it exists: to make the NEXT bug of this class
 * findable without leaving CI.
 *
 * What is asserted throughout is not that the run succeeds — under this
 * profile it should not, entirely — but that it stays HONEST: it finishes, it
 * does not lose track of what it drew, it repairs what it can, and what it
 * reports matches what the host was left holding.
 */

const cfgFor = (kind: ChartKind): ChartConfig => ({ ...sampleConfig(kind), ...DEFAULT_SIZE });

/** A demo-shaped run: charts with configs, plus untagged harness pages. */
function demoItems(n: number) {
  const kinds: ChartKind[] = ["stacked", "line", "clustered", "pie", "area", "scatter"];
  return Array.from({ length: n }, (_, i) => {
    const cfg = cfgFor(kinds[i % kinds.length]);
    return {
      scene: buildChart(cfg),
      title: `Item ${i}`,
      // Every third item is a harness page: no config, so it is never expected
      // to come back re-editable. Mixing them in matters — a repair pass that
      // treats "no config" as "lost its config" rewrites them.
      tagData: i % 3 === 0 ? undefined : JSON.stringify(cfg),
    };
  });
}

beforeEach(() => {
  _setBatchTimeoutForTest(50);
});
afterEach(() => {
  _setBatchTimeoutForTest(0);
  vi.unstubAllGlobals();
});

describe("the demo run against a web host at its worst", () => {
  it("finishes, and reports one result per item", async () => {
    installHost([makeSlide("s1")]);
    applyWebProfile();
    const items = demoItems(9);
    const report = await insertDemoDeck(items, undefined, { reconcile: true });
    // The bar this profile sets is "finishes and accounts for itself", not
    // "succeeds". A run that threw here, or silently returned fewer results
    // than it was given items, is the failure mode that costs a real round.
    expect(report.results).toHaveLength(items.length);
    expect(report.run).toBeTruthy();
    expect(report.totalMs).toBeGreaterThanOrEqual(0);
  });

  it("never reports more slides added than the deck actually grew", async () => {
    // The measurement that unmasks a lost slide. Reading it against
    // `results.length` instead of the settled count once reported 0 lost
    // during real corruption, because a stray from a retry cancelled it out.
    const slides = [makeSlide("s1")];
    installHost(slides);
    applyWebProfile();
    const before = slides.length;
    const report = await insertDemoDeck(demoItems(6), undefined, { reconcile: true });
    expect(report.slidesAdded).toBe(slides.length - before);
    expect(report.addsIssued).toBeGreaterThanOrEqual(report.slidesAdded);
  });

  it("stamps every slide it added with this run's token, and only this run's", async () => {
    // The token is what lets the repair pass — and `npm run triage` — tell
    // this run's slides from whatever an earlier one left in the same deck.
    // A run that failed to write it left the pass unable to identify its own
    // work, and "cannot tell" has ended in a delete.
    const slides = [makeSlide("s1")];
    installHost(slides);
    applyWebProfile();
    const report = await insertDemoDeck(demoItems(6), undefined, { reconcile: true });
    const tokens = slides
      .map((s) => s.tagStore.get(DEMO_SLOT_TAG))
      .filter((v): v is string => !!v)
      .map((v) => JSON.parse(v).run);
    expect(tokens.length).toBeGreaterThan(0);
    expect(new Set(tokens)).toEqual(new Set([report.run]));
  });

  it("does not claim a chart is re-editable when the host swallowed the tag", async () => {
    // `strictTags` refuses a tag write on a proxy more than one sync old,
    // which is the exact failure a real run hit 28 times in one pass. What
    // matters is not that it happens — it will — but that the run's own
    // account of it matches the slides: a chart reported tagged that carries
    // no config tag is a report that sends the next diagnosis the wrong way.
    const slides = [makeSlide("s1")];
    installHost(slides);
    applyWebProfile();
    const report = await insertDemoDeck(demoItems(6), undefined, { reconcile: true });
    for (const [i, r] of report.results.entries()) {
      if (!r.tagged) continue;
      const slide = slides.find((s) => {
        const slot = s.tagStore.get(DEMO_SLOT_TAG);
        return !!slot && JSON.parse(slot).i === i;
      });
      if (!slide) continue;
      const tagged = slide.shapes.items.some((sh) => sh.tagStore.has(CHART_TAG));
      expect(tagged, `item ${i} reported tagged but its slide carries no config`).toBe(true);
    }
  });

  it("survives a host that also stalls mid-run", async () => {
    // Everything above, plus a sync that never answers. The web host does all
    // of this in the same run; a profile that stopped short of it would be
    // describing a nicer host than the one that crashed the tab.
    installHost([makeSlide("s1")]);
    applyWebProfile();
    stallSyncOn.add(4);
    const report = await insertDemoDeck(demoItems(6), undefined, { reconcile: true });
    expect(report.results).toHaveLength(6);
    // A stalled sync must not be recorded as a rendered chart.
    for (const r of report.results) expect(["rendered", "failed", "skipped"]).toContain(r.status);
  });

  it("still gets its degraded pictures tagged, despite the stale-proxy trap", async () => {
    // The case that cost a whole real-host round. Past the shape budget the
    // run stops drawing shapes and inserts a picture instead — ONE shape — and
    // the tag write then had no fresh proxy to aim at, because the re-fetch
    // only ran for items with enough shapes to be worth grouping. A real run
    // failed 28 tag writes this way with `InvalidParam passed to GetItem(id)`,
    // and every one of those charts came back not re-editable.
    //
    // `strictTags` is what makes this reproducible here: a proxy older than
    // one sync is refused, exactly as the web host refuses it.
    const slides = [makeSlide("s1")];
    installHost(slides);
    applyWebProfile();
    const items = demoItems(6);
    const report = await insertDemoDeck(items, undefined, {
      reconcile: true,
      // Low enough that everything after the first item degrades.
      shapeBudget: 1,
      pictureFor: async () => "data:image/png;base64,UE5H",
    });
    expect(report.degradedAt, "the budget did not force any degradation").toBeGreaterThan(0);
    // Charts (not the untagged harness pages) drawn as pictures must still
    // carry their config. Without the single-shape re-fetch they carry none.
    const from = report.degradedAt!;
    const wanted = items.map((it, i) => ({ it, i })).filter(({ it, i }) => it.tagData && i >= from);
    expect(wanted.length).toBeGreaterThan(0);
    for (const { i } of wanted) {
      expect(report.results[i].tagged, `degraded chart ${i} never got its config tag`).toBe(true);
    }
  });

  it("is a genuinely hostile profile — the same run on a clean host does better", async () => {
    // The control. Without it this file could be asserting nothing: a profile
    // that quietly failed to apply would pass every case above, and the suite
    // would read as coverage of a hostile host while exercising a polite one.
    const hostile = [makeSlide("s1")];
    installHost(hostile);
    applyWebProfile();
    const bad = await insertDemoDeck(demoItems(6), undefined, { reconcile: true });
    vi.unstubAllGlobals();

    const clean = [makeSlide("s1")];
    installHost(clean);
    const good = await insertDemoDeck(demoItems(6), undefined, { reconcile: true });

    const grouped = (r: typeof bad) => r.results.filter((x) => x.grouped).length;
    expect(grouped(good)).toBeGreaterThan(grouped(bad));
  });
});

describe("the web profile itself", () => {
  it("turns on every fault a real host has shown us", () => {
    installHost([makeSlide("s1")]);
    applyWebProfile();
    expect(faults.strictGroup).toBe(true);
    expect(faults.strictTags).toBe(true);
    expect(faults.hollowReads).toBeGreaterThan(0);
    expect(faults.refuseGroups).toBeGreaterThan(0);
  });

  it("is undone by installing a host, so it cannot leak into the next test", () => {
    installHost([makeSlide("s1")]);
    applyWebProfile();
    installHost([makeSlide("s2")]);
    expect(faults.strictGroup).toBe(false);
    expect(faults.strictTags).toBe(false);
    expect(faults.hollowReads).toBe(0);
  });
});

describe("the everyday paths on a host that refuses stale proxies", () => {
  /** A chart big enough to need more than one batch — where the trap lives. */
  const bigChart = () => buildChart({ ...sampleConfig("clustered"), ...DEFAULT_SIZE });

  it("keeps an ordinary insert grouped and re-editable", async () => {
    // The demo path learned this lesson in #238 and the everyday path did not.
    // `insertSceneIntoSlide` never asked for a proxy refresh — the re-fetch
    // matched shapes by "the last N on the slide", which is true of a blank
    // slide a run just added and false of the slide the user is looking at, so
    // the ordinary path could not opt in. A 24-shape chart takes three batches
    // at ten a sync; by grouping time the first batch's proxies are three syncs
    // old, the host refuses them, and the chart lands as a heap of shapes with
    // no config tag on it — not re-editable at all, on the single most-used
    // action in the add-in.
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.strictGroup = true;
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    await insertSceneIntoSlide(bigChart(), { tagData: JSON.stringify(cfg) });
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.length, "the chart took more than one batch").toBeGreaterThan(10);
    expect(live.filter((s) => s.grouped).length, "chart did not survive as one group").toBe(1);
    expect(live.filter((s) => s.tagStore.has(CHART_TAG)).length, "chart is not re-editable").toBe(1);
  });

  it("keeps a chart re-editable even when grouping itself is refused", async () => {
    // Grouping and tagging fail for the SAME reason and are recovered
    // separately: losing the group costs a tidy object on the slide, losing
    // the tag costs the chart. On a host that simply will not group, the tag
    // must still land — and it used to be aimed at `created[0]`, the oldest
    // proxy the run holds and the one guaranteed to be refused when staleness
    // was what broke grouping in the first place.
    const slide = makeSlide("s1");
    installHost([slide]);
    applyWebProfile();
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    await insertSceneIntoSlide(bigChart(), { tagData: JSON.stringify(cfg) });
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.filter((s) => s.tagStore.has(CHART_TAG)).length, "chart is not re-editable").toBe(1);
  });

  it("tags its own shape, not a bystander, when the slide already had shapes on it", async () => {
    // What made the re-fetch safe to turn on here. It used to identify a
    // chart's shapes as "the last N on the slide", which holds only for a
    // slide the run added blank. `insertSceneIntoSlide` draws onto whatever
    // the user is looking at — a slide that can already hold anything — and
    // there the rule picks up the user's own shapes and tags one of them.
    // Matching by id is exact and does not care what else is on the slide.
    const slide = makeSlide("s1");
    const theirs = Array.from({ length: 5 }, (_, k) => {
      const sh = makeShape("geometric", "rectangle", { left: k, top: 0, width: 5, height: 5 });
      sh.name = `user shape ${k}`;
      slide.created.push(sh);
      return sh;
    });
    installHost([slide]);
    faults.strictGroup = true;
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    await insertSceneIntoSlide(bigChart(), { tagData: JSON.stringify(cfg) });
    for (const sh of theirs) {
      expect(sh.tagStore.has(CHART_TAG), `${sh.name} was tagged as if it were the chart`).toBe(false);
      expect(sh.grouped, `${sh.name} was swept into the chart's group`).toBeUndefined();
    }
    const group = slide.created.filter((s) => !s.deleted && s.grouped);
    expect(group).toHaveLength(1);
    expect(group[0].grouped, "the group swallowed the user's shapes").toHaveLength(
      slide.created.filter((s) => !s.deleted).length - theirs.length - 1,
    );
  });

  it("does not sweep the user's own shapes into a chart that lost a batch", async () => {
    // The destructive version of the bystander case, and the reason a partial
    // id match beats a positional guess. When a batch's sync fails the host
    // discards it, so the slide holds FEWER shapes than the run drew — and
    // "the chart is the last N shapes" then reaches back past the chart into
    // whatever was already on the slide. Those shapes get grouped into the
    // chart and carried in its parts list, so the next edit deletes them.
    const slide = makeSlide("s1");
    const theirs = Array.from({ length: 6 }, (_, k) => {
      const sh = makeShape("geometric", "rectangle", { left: k, top: 0, width: 5, height: 5 });
      sh.name = `user shape ${k}`;
      slide.created.push(sh);
      return sh;
    });
    installHost([slide]);
    faults.strictGroup = true;
    faults.failSyncOn = 2; // the chart's second batch is discarded by the host
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    await insertSceneIntoSlide(bigChart(), { tagData: JSON.stringify(cfg) }).catch(() => {});
    for (const sh of theirs) {
      expect(sh.tagStore.has(CHART_TAG), `${sh.name} was tagged as the chart`).toBe(false);
      const inSomeGroup = slide.created.some((g) => (g.grouped as unknown[] | undefined)?.includes(sh));
      expect(inSomeGroup, `${sh.name} was grouped into the chart and will be deleted with it`).toBe(false);
    }
  });

  it("keeps an edit in place re-editable", async () => {
    // An update redraws every shape, so it spans the same batches and hits the
    // same trap. A chart that stops being re-editable when you edit it is
    // worse than one that never was: the pane hands back a target it cannot
    // use again, and the next edit silently does nothing.
    const slide = makeSlide("s1");
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const chart = makeShape("geometric", "rectangle", { left: 10, top: 10, width: 100, height: 100 });
    chart.name = "PowerChart";
    chart.tagStore.set(CHART_TAG, JSON.stringify(cfg));
    slide.created.push(chart);
    installHost([slide], [chart], slide);
    faults.strictGroup = true;
    const next = { ...cfg, title: "edited" };
    const target = await updateChartInSlide(
      buildChart(next),
      { slideId: "s1", shapeId: chart.id, left: 10, top: 10 },
      { tagData: JSON.stringify(next) },
    );
    expect(target, "the update lost its target").toBeTruthy();
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.filter((s) => s.tagStore.has(CHART_TAG)).length, "chart is not re-editable").toBe(1);
  });
});
