// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  installHost,
  makeSlide,
  makeShape,
  applyWebProfile,
  faults,
  stallSyncOn,
  failSyncsOn,
  trips,
  unansweredShapeReads,
  unansweredNullChecks,
} from "./helpers/office-host";
import {
  insertDemoDeck,
  insertSceneIntoSlide,
  listChartsInDeck,
  listChartsInSelection,
  loadChartFromSelection,
  updateChartInSlide,
  updateChartsInSlides,
  errorText,
  DEMO_SLOT_TAG,
  CHART_TAG,
  _setBatchTimeoutForTest,
} from "../src/render/powerpoint";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { setTracing, traceLog } from "../src/core/trace";
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

  /**
   * The cascade the self-test's `same scale across the deck` scenario reported
   * as "3 of 8 charts carry the shared scale; 3 still re-editable".
   *
   * A real host refused the grouping with `InvalidParam passed to GetItem(id)`,
   * and the tag target it fell back to then answered `.tags` as UNDEFINED:
   * "Cannot read properties of undefined (reading 'add')", once per chart,
   * five charts in a row. `groupAndTagAll` survives that per chart — it stopped
   * costing the whole batch a round ago — but surviving it still meant shipping
   * a chart with no config on it, and the ordinary paths had no recovery.
   *
   * The demo path always had one: `insertDemoDeck` re-reads the settled deck
   * and plans a `retag`, which is why the same run's 23 lost tags came back and
   * these did not. `settleAndTagChart` gives the insert and update paths the
   * same second chance, from a context the failed one cannot poison.
   */
  it("settles the config tag from a fresh context when the drawing one could not write it", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    applyWebProfile();
    // The drawing context's tag target hands back no `.tags` at all — the
    // real host's answer, and a synchronous throw where the write is queued.
    faults.tagsUndefinedOn = 1;
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const target = await insertSceneIntoSlide(bigChart(), { tagData: JSON.stringify(cfg) });
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.filter((s) => s.tagStore.has(CHART_TAG)).length, "chart is not re-editable").toBe(1);
    // And the caller is told so. A chart that was settled but still reported
    // `lost: "no-config"` makes the pane tell the user their chart cannot be
    // re-opened, about a chart that can.
    expect(target?.lost, "reported as lost after the tag had been settled").toBeUndefined();
  });

  /**
   * Editing a chart on a slide this session added.
   *
   * The case every existing test missed, because they all edit a slide that was
   * already in the deck — where a by-id handle round-trips and holding one
   * across a sync is free. A demo deck's slides are minutes old when someone
   * edits a chart on one, and `same scale across the deck` updates charts on
   * slides the battery itself just inserted.
   *
   * `updateChartsInSlides` resolved the slide, synced, and then reached through
   * that held proxy for its shapes and for every redraw batch — the pattern
   * PowerPoint web answers with `GeneralException` at `SlideCollection.getItem`.
   * Not a degraded result: it threw out of the whole update. The comment on the
   * thunk said "an existing slide's proxy is stable across syncs — hold it",
   * which is true, and load-bearing on an assumption that was never checked.
   */
  it("edits a chart on a slide added in this session, not just a pre-existing one", async () => {
    installHost([makeSlide("s1")]);
    const { addScratchSlide } = await import("../src/render/powerpoint");
    const slideId = await addScratchSlide();
    expect(slideId, "no freshly-added slide to edit on").toBeTruthy();
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const first = await insertSceneIntoSlide(buildChart(cfg), { slideId: slideId!, tagData: JSON.stringify(cfg) });
    expect(first, "the setup insert produced no target").toBeTruthy();
    const next = { ...cfg, scale: { max: 99 } };
    const after = await updateChartInSlide(buildChart(next), first!, { tagData: JSON.stringify(next) });
    expect(after, "the update produced no target at all").toBeTruthy();
    expect(after?.lost, `the update reported the chart lost: ${after?.lost}`).toBeUndefined();
  });

  it("settles an UPDATE's config tag too, which is the path same-scale drives", async () => {
    // `same scale across the deck` edits every probe chart through
    // `updateChartInSlide` and then counts how many still carry a config. On a
    // real host it counted three of eight: an update redraws every shape, so it
    // meets the same staleness a fresh insert does, and it had the same lack of
    // recovery. The insert path is not a proxy for this one — they are separate
    // functions with separate contexts, and only one of them was fixed first.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const first = await insertSceneIntoSlide(bigChart(), { tagData: JSON.stringify(cfg) });
    expect(first, "the setup insert did not produce a target").toBeTruthy();
    applyWebProfile();
    faults.tagsUndefinedOn = 1;
    const next = { ...cfg, scale: { max: 99 } };
    const after = await updateChartInSlide(buildChart(next), first!, { tagData: JSON.stringify(next) });
    expect(after?.lost, "an updated chart was left reported as un-re-editable").toBeUndefined();
    const live = slide.created.filter((s) => !s.deleted);
    expect(
      live.filter((s) => s.tagStore.get(CHART_TAG) === JSON.stringify(next)).length,
      "the updated chart does not carry the new config",
    ).toBe(1);
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

/**
 * Reading a proxy the host never answered for.
 *
 * Three shapes of one mistake, all three found by the self-test battery on a
 * real PowerPoint and none of them reachable here before the fake learned to
 * refuse an unloaded read. Office.js does not hand back `undefined` for a
 * property it has no value for — it throws `PropertyNotLoaded`, from the
 * getter, at whatever line happens to read it. So the failure never surfaces
 * where the mistake is; it surfaces as a crash in code that looks correct.
 *
 * Together they cost three of six self-test scenarios in one run, and what
 * they had in common is that the WORK had already succeeded: the slide was
 * there, the deck was scannable, the charts were drawn. Only the reading of it
 * fell over.
 */
describe("proxies the host would not answer for", () => {
  it("edits a chart in place after resolving its slide", async () => {
    // `updateChartsInSlides` resolved the target's slide with
    // getItemOrNullObject and asked for `isNullObject` BY NAME, which selects
    // nothing and leaves the flag unreadable. It then read the flag on the
    // very next line — so an in-place edit threw before it deleted anything,
    // and "edit a chart on the visible slide" failed on a host where the
    // slide, the chart and the tag were all exactly where they should be.
    const slide = makeSlide("s1");
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const chart = makeShape("geometric", "rectangle", { left: 10, top: 10, width: 100, height: 100 });
    chart.name = "PowerChart";
    chart.tagStore.set(CHART_TAG, JSON.stringify(cfg));
    slide.created.push(chart);
    installHost([slide], [chart], slide);
    const next = { ...cfg, title: "edited" };
    const target = await updateChartInSlide(
      buildChart(next),
      { slideId: "s1", shapeId: chart.id, left: 10, top: 10 },
      { tagData: JSON.stringify(next) },
    );
    expect(target, "the edit resolved no target — its slide read as gone").toBeTruthy();
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.filter((s) => s.tagStore.has(CHART_TAG)).length, "the edited chart is not re-editable").toBe(1);
  });

  it("scans the rest of the deck when one slide will not answer", async () => {
    // The deck-wide scan queued one shape-collection load per slide into a
    // SINGLE sync and read every `.items` straight afterwards. On a 38-slide
    // deck the web host answers that request incompletely; one unanswered
    // collection threw out of the whole call, so Same Scale reported a crash
    // instead of rescaling the charts it could see — and so did every
    // self-test scenario that begins by asking what is in the deck.
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const slides = ["s1", "s2"].map((id) => {
      const s = makeSlide(id);
      const chart = makeShape("geometric", "rectangle", { left: 10, top: 10, width: 100, height: 100 });
      chart.name = "PowerChart";
      chart.tagStore.set(CHART_TAG, JSON.stringify({ ...cfg, title: id }));
      s.created.push(chart);
      return s;
    });
    installHost(slides);
    unansweredShapeReads.add("s1");
    const found = await listChartsInDeck();
    expect(
      found.charts.map((c) => JSON.parse(c.configJson).title),
      "the silent slide took the whole scan with it",
    ).toEqual(["s2"]);
    // And the scan SAYS it missed one, rather than handing back a short list
    // that reads exactly like a deck with one chart in it.
    expect(found.unread, "the silent slide was not reported").toBe(1);
  });

  it("never turns a refused sync into a property-read crash", async () => {
    // The tagging sync is where the web host refuses a stale proxy
    // ("InvalidParam passed to GetItem(id)", code 5010) — a caught, logged,
    // survivable outcome: the shapes are on the slide and the only loss is
    // re-editability. But that same sync carries the `load("id,left,top")`
    // that tells the caller WHERE the chart landed, so when it went down the
    // caller was handed a proxy holding nothing and threw `PropertyNotLoaded`
    // at `Shape.left`. A deck-wide rescale then reported a crash for charts it
    // had correctly redrawn.
    //
    // Every sync in the insert, not just the tagging one: which sync a host
    // refuses is not ours to choose, and the rule is the same for all of them.
    // A refusal may fail the insert. It may never fail it by reading.
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    installHost([makeSlide("s1")]);
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg) });
    const syncs = trips.syncs;
    expect(syncs, "the insert made no syncs to refuse").toBeGreaterThan(1);

    for (let n = 1; n <= syncs; n++) {
      installHost([makeSlide("s1")]);
      faults.strictShapeReads = true;
      failSyncsOn.add(n);
      let thrown: unknown;
      await insertSceneIntoSlide(scene, { tagData: JSON.stringify(cfg) }).catch((err) => (thrown = err));
      faults.strictShapeReads = false;
      expect(
        String((thrown as Error)?.message ?? ""),
        `sync ${n} of ${syncs} was refused and the insert read on`,
      ).not.toMatch(/PropertyNotLoaded|is not available/i);
    }
  });

  it("reports where a chart landed even when there was nothing to tag", async () => {
    // No `tagData` means no tagging phase, and the load that resolves the
    // target's position rode inside it — so the one path that never tags was
    // also the one that always handed back an unreadable target. Same on any
    // host below PowerPointApi 1.3, where the tagging phase is skipped outright.
    installHost([makeSlide("s1")]);
    faults.strictShapeReads = true;
    const target = await insertSceneIntoSlide(buildChart({ ...sampleConfig("clustered"), ...DEFAULT_SIZE }), {});
    faults.strictShapeReads = false;
    expect(target, "an untagged insert lost track of its own chart").toBeTruthy();
    expect(Number.isFinite(target!.left) && Number.isFinite(target!.top), "the target carries no position").toBe(true);
  });
});

/**
 * The everyday paths, run end to end with every proxy read held to the
 * contract Office.js actually enforces.
 *
 * The three failures that started this were found one at a time, from a real
 * host, by reading a log. Fixing them one at a time is how the next three get
 * found the same expensive way — the bug is not any one line, it is that a
 * property read looks identical whether or not the host answered for it. So
 * these drive the whole insert / edit / scan / select surface with
 * `strictShapeReads` on, and assert only one thing: the add-in never fails by
 * READING. Failing because a host refused something is allowed. Failing on the
 * line that reads the answer is not.
 */
describe("the everyday paths under a host that answers nothing it was not asked for", () => {
  const cfg = () => ({ ...sampleConfig("clustered"), ...DEFAULT_SIZE });

  /** A slide holding one re-editable chart, as every entry point expects. */
  function deckWithChart(id = "s1") {
    const slide = makeSlide(id);
    const chart = makeShape("geometric", "rectangle", { left: 10, top: 10, width: 100, height: 100 });
    chart.name = "PowerChart";
    chart.tagStore.set(CHART_TAG, JSON.stringify({ ...cfg(), title: id }));
    slide.created.push(chart);
    return { slide, chart };
  }

  /** Every host entry point the pane reaches, run once. */
  async function everydayPaths() {
    await loadChartFromSelection();
    await insertSceneIntoSlide(buildChart(cfg()), { tagData: JSON.stringify(cfg()) });
    const deck = await listChartsInDeck();
    if (deck.charts.length) {
      await updateChartInSlide(buildChart({ ...cfg(), title: "edited" }), deck.charts[0].target, {
        tagData: JSON.stringify({ ...cfg(), title: "edited" }),
      });
    }
    await listChartsInSelection();
    return deck.charts;
  }

  it("runs the pane's whole surface against a host that answers only some of it", async () => {
    // Not one quiet answer — a scattering of them, across every entry point,
    // with the shape contract enforced throughout. Each individual gap is
    // covered below; this is the one that asks whether they compose, because
    // a real host does not hand out its silences one per call and stop.
    const { slide, chart } = deckWithChart();
    installHost([slide], [chart], slide);
    // Read the id BEFORE arming the gate: the fault is addressed by id, and
    // asking for it under the gate is the test tripping over its own fixture.
    unansweredNullChecks.add(chart.id); // the update gets a quiet shape
    faults.strictShapeReads = true;
    faults.unansweredTagLoads = 1; // the selection read gets a quiet tag
    let thrown: unknown;
    try {
      const deck = await everydayPaths();
      expect(deck.length, "the deck scan found nothing").toBeGreaterThan(0);
    } catch (err) {
      thrown = err;
    } finally {
      faults.strictShapeReads = false;
      faults.unansweredTagLoads = 0;
    }
    expect(thrown && errorText(thrown), "an everyday path threw").toBeFalsy();
  });

  it("treats a tag the host stayed quiet about as 'not a chart', not as a crash", async () => {
    // `isNullObject` and `value` on a config tag are what every "is this shape
    // a chart" question comes down to, and both were read raw. A host that
    // takes the load and answers nothing then does not mean "no chart here" —
    // it means the pane's most-used reads throw. One quiet tag per entry point
    // is enough: what has to survive is the CALL, whatever it concludes.
    const { slide, chart } = deckWithChart();
    installHost([slide], [chart], slide);
    faults.unansweredTagLoads = 3;
    let thrown: unknown;
    try {
      await loadChartFromSelection();
      faults.unansweredTagLoads = 3;
      await listChartsInDeck();
      faults.unansweredTagLoads = 3;
      await listChartsInSelection();
    } catch (err) {
      thrown = err;
    } finally {
      faults.unansweredTagLoads = 0;
    }
    expect(thrown && errorText(thrown), "a quiet tag took a read down with it").toBeFalsy();
  });

  it("leaves a sibling it cannot see alone, and still finishes the redraw", async () => {
    // An ungrouped chart's siblings travel in its parts tag, and the update
    // deletes the set. A part the host will not confirm is a part we cannot
    // prove is ours — and the shapes on that slide include whatever the user
    // put there. Covering a stray with the redraw is visible and fixable;
    // deleting somebody else's shape is neither.
    //
    // Both halves are asserted, and the second is what makes this a test. The
    // raw read threw, which also left the shape undeleted — so "the sibling
    // survived" alone is true of the bug as well as the fix. What separates
    // them is that the bug took the whole update down with it.
    const { slide, chart } = deckWithChart();
    const theirs = makeShape("geometric", "rectangle", { left: 300, top: 10, width: 20, height: 20 });
    theirs.name = "not ours";
    slide.created.push(theirs);
    installHost([slide], [chart], slide);
    unansweredNullChecks.add(theirs.id); // it resolves, and the host says nothing back
    const next = { ...cfg(), title: "edited" };
    const target = await updateChartInSlide(
      buildChart(next),
      { slideId: "s1", shapeId: chart.id, left: 10, top: 10, partIds: [theirs.id] },
      { tagData: JSON.stringify(next) },
    ).catch(() => null);
    expect(theirs.deleted, "a shape the host would not confirm was deleted anyway").toBe(false);
    expect(target, "one unreadable sibling failed the whole redraw").toBeTruthy();
  });

  it("tags the charts it can when one target has no tags at all", async () => {
    // A real host answered `shape.tags` as UNDEFINED — "Cannot read properties
    // of undefined (reading 'add')", four times in one run, each on a chart
    // whose grouping had just been refused with InvalidParam 5010. That throw
    // is SYNCHRONOUS, so it escaped the tagging loop and took the whole
    // batch's tagging with it: every chart after the bad one lost its config
    // without ever being attempted.
    const cfgJson = JSON.stringify(cfg());
    const slide = makeSlide("s1");
    installHost([slide]);
    faults.tagsUndefinedOn = 1; // the FIRST shape to be tagged has no .tags
    let thrown: unknown;
    const targets = await Promise.all([
      insertSceneIntoSlide(buildChart(cfg()), { tagData: cfgJson }).catch((e) => {
        thrown = e;
        return null;
      }),
      insertSceneIntoSlide(buildChart(cfg()), { tagData: cfgJson }).catch(() => null),
    ]);
    faults.tagsUndefinedOn = 0;
    expect(thrown && errorText(thrown), "a missing .tags took the insert down").toBeFalsy();
    // The second chart must still be re-editable: one unusable target is not a
    // reason to abandon the charts either side of it.
    const tagged = slide.created.filter((sh) => !sh.deleted && sh.tagStore.has(CHART_TAG));
    expect(tagged.length, "a chart with a usable target went untagged").toBeGreaterThan(0);
    expect(targets.filter(Boolean).length).toBeGreaterThan(0);
  });

  it("says WHICH phase an error escaped, in the message and in the log", async () => {
    // The run log used to carry what the host refused and nothing about what
    // the add-in was doing when it refused. Placing three real-host failures
    // meant reasoning from timestamps and call order back to a line, for each
    // one. `errorLocation` names the Office.js type; this names the phase.
    setTracing(true);
    try {
      installHost([makeSlide("s1")]);
      failSyncsOn.add(2); // mid-render, after the first batch has committed
      let thrown: unknown;
      const cfgJson = JSON.stringify(cfg());
      await insertSceneIntoSlide(buildChart(cfg()), { tagData: cfgJson }).catch((err) => (thrown = err));
      expect(thrown, "the refused sync did not surface").toBeTruthy();
      expect(errorText(thrown), "the error does not say what was running").toMatch(/at=drawing the chart's shapes/);
      // And in the log, at the moment it happened — so a phase is on record
      // even when the error is swallowed by a best-effort catch further up.
      const phases = traceLog()
        .entries.filter((e) => e.scope === "error")
        .map((e) => e.message);
      expect(phases, "no phase was traced").toContain("drawing the chart's shapes");
    } finally {
      setTracing(false);
    }
  });

  it("skips a slide it cannot confirm instead of failing the update", async () => {
    // The slide resolve is the FIRST thing an in-place update does, for every
    // chart in the batch at once. A host that stays quiet about one of them
    // used to throw out of the whole call — so Same Scale across a deck lost
    // every chart to one unanswered slide, including the charts it had not
    // reached yet.
    const a = deckWithChart("s1");
    const b = deckWithChart("s2");
    installHost([a.slide, b.slide], [a.chart], a.slide);
    unansweredNullChecks.add("s1");
    const next = { ...cfg(), title: "edited" };
    const out = await updateChartsInSlides([
      {
        scene: buildChart(next),
        target: { slideId: "s1", shapeId: a.chart.id, left: 10, top: 10 },
        opts: { tagData: JSON.stringify(next) },
      },
      {
        scene: buildChart(next),
        target: { slideId: "s2", shapeId: b.chart.id, left: 10, top: 10 },
        opts: { tagData: JSON.stringify(next) },
      },
    ]).catch(() => []);
    expect(out.length, "the quiet slide took the readable one down with it").toBe(1);
    expect(a.chart.deleted, "a chart on a slide we could not confirm was deleted").toBe(false);
  });
});

/**
 * The fake host's own strictness, asserted directly.
 *
 * Every other test here drives production code and trusts the double
 * underneath it. These drive the double, because a fake that is kinder than
 * the host it models does not fail — it passes, wrongly, and takes a real
 * defect with it. That is not hypothetical: resolving a tag proxy and reading
 * it without `load()` is `PropertyNotLoaded` on PowerPoint web and was a plain
 * property read here, which is how editing an ungrouped chart came to be
 * impossible in the field with 1700 green tests, and how `deleteSlide` came to
 * refuse every guarded delete it was ever asked for.
 *
 * Softening any of these makes the suite lie again. They are here so that
 * doing so fails loudly rather than silently widening what the tests accept.
 */
describe("the fake host models Office.js strictness, not convenience", () => {
  it("throws PropertyNotLoaded when a shape tag is read without load()", () => {
    const shape = makeShape("geometric", "rectangle", { left: 0, top: 0, width: 10, height: 10 });
    shape.tagStore.set(CHART_TAG, "{}");
    installHost([makeSlide("s1")]);
    const tag = shape.tags.getItemOrNullObject(CHART_TAG);
    expect(() => tag.isNullObject).toThrow(/not available|PropertyNotLoaded/i);
    expect(() => tag.value).toThrow(/not available|PropertyNotLoaded/i);
  });

  it("keeps a tag unreadable until the sync its load() was queued in lands", async () => {
    // load() alone is not enough — the value arrives with the SYNC. A fake that
    // resolved on load() would accept code that reads a tag it queued in the
    // same breath, which the real host refuses.
    const slide = makeSlide("s1");
    slide.tagStore.set(DEMO_SLOT_TAG, JSON.stringify({ i: 0, title: "Line", run: "r" }));
    installHost([slide]);
    await PowerPoint.run(async (context) => {
      const s = context.presentation.slides.getItemOrNullObject("s1");
      s.load("id");
      await context.sync();
      const tag = (
        s as unknown as { tags: { getItemOrNullObject(k: string): { value: string; load(p?: string): void } } }
      ).tags.getItemOrNullObject(DEMO_SLOT_TAG);
      tag.load("value");
      expect(() => tag.value, "readable before its sync").toThrow(/not available|PropertyNotLoaded/i);
      await context.sync();
      expect(JSON.parse(tag.value).title).toBe("Line");
    });
  });

  it("throws PropertyNotLoaded when a slide's isNullObject is read without load()", async () => {
    installHost([makeSlide("s1")]);
    await PowerPoint.run(async (context) => {
      const missing = context.presentation.slides.getItemOrNullObject("nope");
      expect(() => missing.isNullObject).toThrow(/not available|PropertyNotLoaded/i);
      missing.load("id");
      await context.sync();
      expect(missing.isNullObject).toBe(true);
    });
  });

  it('does not count load("isNullObject") as a load', async () => {
    // The flag is not a property the host holds — it is set from the response
    // to a load of REAL properties. Asking for it by name selects nothing, so
    // the proxy never joins the sync and the flag stays unreadable. The fake
    // used to accept it, and five resolves in powerpoint.ts were written that
    // way; the one on the in-place update path is why editing a chart threw
    // before it deleted anything on PowerPoint on the web.
    installHost([makeSlide("s1")]);
    await PowerPoint.run(async (context) => {
      const missing = context.presentation.slides.getItemOrNullObject("nope");
      missing.load("isNullObject");
      await context.sync();
      expect(() => missing.isNullObject, 'load("isNullObject") resolved the flag').toThrow(
        /not available|PropertyNotLoaded/i,
      );
    });
  });
});
