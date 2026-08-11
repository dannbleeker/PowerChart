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
  CHART_PARTS_TAG,
  _setBatchTimeoutForTest,
  _setReadbackTimeoutForTest,
  addScratchSlide,
  chooseGroupMembers,
  OFFSCREEN_BATCH,
  mayTakeConfig,
  afterSettle,
  targetWithNoTagResult,
  rasteriseTimeoutMs,
  readbackTimeoutMs,
  enableExtendedErrorLogging,
  trimDebugInfo,
  untaggedIndices,
  insertAgendaSlides,
  MAX_ADD_RETRY_ROUNDS,
} from "../src/render/powerpoint";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { buildAgendaScene } from "../src/core/agenda";
import { sampleConfig } from "../src/core/samples";
import { setTracing, traceLog, traceMark } from "../src/core/trace";
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
    // carry their config — RE-EDITABLE, by whichever writer got there.
    //
    // This asserted `report.results[i].tagged`, which is narrower: it is true
    // only when the DRAWING CONTEXT's write landed. That stopped being the right
    // question once the re-read above started asking for `items/id`, because the
    // hollow-read model then reaches it and the single-shape re-fetch does not
    // run — and on the real host it never runs anyway, since the re-read comes
    // back empty there whatever projection is asked for (four confirmations in
    // the 2026-08-07 round). The demo path repairs those charts and does it well:
    // three `applying retag` lines, and every chart on the deck carrying its
    // config afterwards.
    //
    // So the property is re-editability, not which writer achieved it. Keeping
    // the narrow assertion would have pinned a code path this host does not have.
    const from = report.degradedAt!;
    const wanted = items.map((it, i) => ({ it, i })).filter(({ it, i }) => it.tagData && i >= from);
    expect(wanted.length).toBeGreaterThan(0);
    const tagged = new Set(
      slides.flatMap((s) =>
        s.created.filter((sh) => !sh.deleted && sh.tagStore.has(CHART_TAG)).map((sh) => sh.tagStore.get(CHART_TAG)),
      ),
    );
    for (const { it, i } of wanted) {
      expect(tagged.has(it.tagData!), `degraded chart ${i} is not re-editable — no shape carries its config`).toBe(
        true,
      );
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

  /**
   * A chart inserted onto a freshly-added slide must land as ONE object.
   *
   * Nothing asserted this, and the cost of that gap was measured rather than
   * imagined: a change that made `insertSceneIntoSlide` re-acquire its slide
   * handle per batch turned a 40-shape chart into an UNGROUPED heap with no
   * `CHART_PARTS_TAG`, and the next edit — unable to find the other 39 shapes —
   * drew a second chart on top of the first. Forty-one shapes became
   * seventy-nine. The whole suite stayed green: 2016 tests, and not one of them
   * looked at grouping or parts on this path.
   *
   * The parts tag is the half that bites later. Its own docstring says it: an
   * in-place update without it "deletes 1 of the chart's 13 shapes and redraws
   * all 13", so the chart grows by a whole chart on every edit.
   */
  it("lands a multi-batch chart on a freshly-added slide as one grouped, re-editable object", async () => {
    const deck = [makeSlide("s1")];
    installHost(deck);
    const { addScratchSlide } = await import("../src/render/powerpoint");
    const slideId = await addScratchSlide();
    expect(slideId, "no freshly-added slide to draw on").toBeTruthy();
    const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
    const scene = buildChart(cfg);
    // More than one batch, which is the case the stale-handle traps live in.
    expect(scene.nodes.length, "this chart is too small to span batches").toBeGreaterThan(10);
    await insertSceneIntoSlide(scene, { slideId: slideId!, tagData: JSON.stringify(cfg) });

    const fresh = deck.find((s) => s.id === slideId);
    const live = fresh?.created.filter((sh) => !sh.deleted) ?? [];
    const groups = live.filter((sh) => sh.grouped);
    const tagged = live.filter((sh) => sh.tagStore.has(CHART_TAG));
    expect(tagged, "the chart carries no config — it is not re-editable").toHaveLength(1);
    // One object, or — if this host would not group — an anchor that knows the
    // rest of its shapes. Either is survivable; neither-nor is what orphans a
    // chart on the next edit.
    const anchor = tagged[0];
    const knowsItsParts = anchor.tagStore.has(CHART_PARTS_TAG);
    expect(
      groups.length === 1 || knowsItsParts,
      `chart landed as ${live.length} loose shapes with no group and no parts tag — the next edit will orphan them`,
    ).toBe(true);
  });

  /**
   * The half that bites later, on the host that actually exists.
   *
   * `ungroupedFallback` reads its ids off `it.created` — the proxies
   * `addGeometricShape` handed back — and by the time it runs those span several
   * batches, so Office.js has rewritten their object paths to
   * `shapes.getItem(id)`. This host refuses that call: `reading back an ungrouped
   * chart's shape ids` failed three times in the 2026-08-07 run with
   * `InvalidParam passed to GetItem(id)`, `errorLocation:
   * ShapeCollection.getItem`, and each failure cost that chart its parts list.
   *
   * The blast radius is the reason this is not a footnote. PowerPoint on the web
   * ungroups every chart it cannot group, so on that host EVERY chart takes this
   * path — and the parts tag's own docstring says what its absence costs: an
   * in-place update "deletes 1 of the chart's 13 shapes and redraws all 13", so
   * the chart grows by a whole chart on every edit.
   *
   * The recovery is the one this host honours everywhere else: members of a
   * collection read. `groupAndTagAll` already performs that read for grouping;
   * this is the same answer, used twice.
   */
  it("writes an ungrouped chart's parts tag from the re-read when created proxies are refused", async () => {
    const deck = [makeSlide("s1")];
    installHost(deck);
    const slideId = await addScratchSlide();
    expect(slideId, "no freshly-added slide to draw on").toBeTruthy();
    // Both halves of the real host. `refuseGroups` is what puts the chart down
    // the ungrouped path at all — on the web that is every chart — and
    // `strictIdLoads` is the host refusing to read an id off a proxy that has
    // outlived its batch, which is what `it.created` are by then.
    faults.refuseGroups = 99;
    faults.strictIdLoads = true;
    try {
      const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
      const scene = buildChart(cfg);
      expect(scene.nodes.length, "this chart is too small to span batches").toBeGreaterThan(10);
      await insertSceneIntoSlide(scene, { slideId: slideId!, tagData: JSON.stringify(cfg) });

      const fresh = deck.find((s) => s.id === slideId);
      const live = fresh?.created.filter((sh) => !sh.deleted) ?? [];
      const tagged = live.filter((sh) => sh.tagStore.has(CHART_TAG));
      expect(tagged, "the chart carries no config at all").toHaveLength(1);
      expect(tagged[0].grouped, "this case is only interesting while the chart is ungrouped").toBeFalsy();
      const parts = tagged[0].tagStore.get(CHART_PARTS_TAG);
      expect(parts, "an ungrouped chart with no parts list — the next edit leaves its shapes behind").toBeTruthy();
      // And it must name the chart's OTHER shapes, not a truncated remnant: a
      // partial list orphans exactly the shapes it left out.
      expect(JSON.parse(parts!), "the parts list is too short to be the rest of the chart").toHaveLength(
        live.length - 1,
      );
    } finally {
      faults.refuseGroups = 0;
      faults.strictIdLoads = false;
    }
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

describe("a chart the update could not name afterwards", () => {
  /**
   * The state a real PowerPoint produced on 2026-08-06, and the one the settle
   * pass was silently never asked about.
   *
   * `same scale across the deck` reported "3 of 8 charts carry the shared
   * scale" and the run log carried FIVE `tagging failed — charts are not
   * re-editable until repaired` events and not a single `settle pass:` line.
   * That trace is unconditional whenever `settleUntaggedCharts` is handed
   * anything, so its absence is proof the settle was never invoked — not proof
   * it failed. The two look identical from a deck and want completely
   * different fixes. (It printed as one fixed sentence then; it now names its
   * outcome, and all three forms share that prefix so this reading survives.)
   *
   * It reaches the update through a hole four host failures wide: grouping
   * refused, tagging refused, the id read-back refused, and the shape
   * unreadable without it. Driving all four end-to-end through the fake is
   * possible but tells you less than asking the rule directly.
   */
  const old = {
    slideId: "s1",
    shapeId: "shape-25",
    left: 60,
    top: 90,
    origin: { left: 60, top: 90, anchorLeft: 60, anchorTop: 90 },
  };

  it("marks a chart that DREW but could not be named as un-re-editable", () => {
    // The whole fix. Bare, this returned `old` untouched, `lost` stayed
    // undefined, and the `no-config` filter that feeds the settle matched
    // nothing.
    const back = targetWithNoTagResult(old, { drew: true, wrecked: false });
    expect(back.lost, "a drawn, untagged, unnameable chart came back looking fine").toBe("no-config");
  });

  it("says no-config, not unknown-shape — the chart is on the slide", () => {
    // `unknown-shape` is what the pane turns into "your chart is gone". This
    // chart is emphatically not gone; it just cannot be re-opened, and telling
    // the user otherwise sends them looking for something that is in front of
    // them.
    const back = targetWithNoTagResult(old, { drew: true, wrecked: true });
    expect(back.lost).toBe("no-config");
  });

  it("still reports a chart wrecked BEFORE it drew as unknown-shape", () => {
    // The pre-existing rule, kept: the delete committed and nothing replaced
    // it, so the old shapeId names something deleted.
    expect(targetWithNoTagResult(old, { drew: false, wrecked: true }).lost).toBe("unknown-shape");
  });

  it("leaves a chart the stop caught before the delete completely alone", () => {
    // Its shapes are still there and its target is still true. Marking it would
    // be a false alarm, and an existing test caught a version that marked both.
    const back = targetWithNoTagResult(old, { drew: false, wrecked: false });
    expect(back.lost).toBeUndefined();
    expect(back).toEqual(old);
  });
});

describe("how long to wait for the host to draw a slide", () => {
  /**
   * Rasterising is not a readback and must not borrow its budget.
   *
   * Every successful `getImageAsBase64` this project has recorded answered in
   * about a second, and PowerPoint on the web has failed the same call three
   * different ways across three rounds: `GeneralException` at
   * `SlideCollection.getItem`, then taking the call and silently producing
   * nothing, then — 2026-08-06 — never answering the sync at all.
   *
   * That last one cost a whole round. `the chart is actually visible` sat on
   * the full ninety-second readback budget and the tab died moments later on
   * the delete that followed, taking the run's own report with it. The honest
   * verdict for that scenario is one word, `skipped`; ninety seconds buys
   * nothing towards it and costs every scenario after it.
   */
  it("gives up on a rasterise far sooner than on a readback", () => {
    _setReadbackTimeoutForTest(90_000);
    try {
      expect(
        rasteriseTimeoutMs(),
        "a rasterise that will answer answers in about a second — this waits as long as a 20-slide page read",
      ).toBeLessThan(readbackTimeoutMs());
      expect(rasteriseTimeoutMs()).toBeLessThanOrEqual(20_000);
    } finally {
      _setReadbackTimeoutForTest(90_000);
    }
  });

  it("never out-waits a readback budget a test has shortened", () => {
    // The trap `readbackTimeoutMs` documents: a hard-coded number is shorter
    // than ninety seconds in production and far LONGER than the milliseconds a
    // test shortens the budget to, which would make the wait untestable at
    // exactly the site that most needed bounding.
    _setReadbackTimeoutForTest(40);
    try {
      expect(rasteriseTimeoutMs(), "a shortened readback budget did not shorten the rasterise").toBeLessThanOrEqual(40);
    } finally {
      _setReadbackTimeoutForTest(90_000);
    }
  });
});

describe("grouping when the host will not name every member", () => {
  /**
   * The state five charts died in on 2026-08-07.
   *
   * `groupAndTagAll` re-reads the shape collection and re-resolves each member
   * by id off a fresh slide handle — right, and documented. But both fallbacks
   * reached for proxies whose PARENT Office.js had rewritten to
   * `slides.getItem(id)`, which this host refuses for a slide the run added.
   * The host listed the statements:
   *
   *   var itemOrNullObject = slides.getItemOrNullObject(...);   // fresh
   *   var slide            = slides.getItem(...);               // rewritten
   *   var shapes1          = slide.shapes;
   *   var shape            = shapes1.getItem(...);              // refused, 5010
   *
   * Handing those to `addGroup` does not merely fail to group — it THROWS, and
   * the throw takes the batch's tagging with it (`Cannot read properties of
   * undefined (reading 'add')`). The chart loses its group AND its config.
   *
   * Asked as the rule rather than end-to-end: reaching it through the fake
   * needs the re-read to answer short, and `groupAndTagAll` loads `"items"`,
   * which neither `hollowReads` (`items/id`) nor `hollowNameReads`
   * (`items/name`) shortens. Bending the fake to that shape would test the
   * bend.
   */
  it("re-resolves by id when every member can be named", () => {
    expect(chooseGroupMembers({ refreshedIds: ["a", "b"], askedForRefresh: true })).toEqual({
      use: "ids",
      ids: ["a", "b"],
    });
  });

  it("groups NOTHING when even one member cannot be named", () => {
    // Not a partial group: the loose remainder would be deleted by the next
    // in-place update, which is worse than never grouping.
    expect(chooseGroupMembers({ refreshedIds: ["a", undefined], askedForRefresh: true })).toEqual({ use: "none" });
  });

  it("groups nothing when the re-read produced no members at all", () => {
    // The `short-0` host: twenty-four shapes drawn, none listed. `created` is
    // NOT a safe fallback — its proxies carry the same rewritten parent, and a
    // host that will not list the slide's shapes is exactly the host that
    // refuses them.
    expect(chooseGroupMembers({ refreshedIds: undefined, askedForRefresh: true })).toEqual({ use: "none" });
  });

  it("still uses the created handles for a chart that never asked for a refresh", () => {
    // Nothing has been re-resolved behind it, so `created` is the only handle
    // there has ever been. The small-chart path that has always worked, and
    // must keep working.
    expect(chooseGroupMembers({ refreshedIds: undefined, askedForRefresh: false })).toEqual({ use: "created" });
  });
});

describe("a slide that was added a moment ago", () => {
  /**
   * office-js#2903's workaround was tried on this host and made it strictly
   * worse. Recorded here so nobody adds it back on the strength of the issue.
   *
   * The issue says a slide added on PowerPoint Online is not usable for a
   * couple of seconds, and its reporter's fix is to wait. `addScratchSlide` did
   * that on 2026-08-07. The next real-host round answered **1 of 25** probe
   * questions against 19 of 26 the build before: the add landed, the wait ran,
   * and the liveness check that follows it then found nothing — so every
   * question came back `no-scratch-slide` and the only row with an answer was
   * the cleanup's own.
   *
   * This host resolves a freshly-added slide's id ONCE and refuses it ever
   * after — the behaviour `deleteSlideById` and `SlideThunk` are already built
   * around. Waiting does not buy it time; it spends the one resolution later,
   * by which point the id is gone.
   */
  it("is used immediately, with no settling pause", async () => {
    installHost([makeSlide("s1")]);
    setTracing(true);
    const mark = traceMark();
    const id = await addScratchSlide();
    expect(id, "the fake would not add a scratch slide").toBeTruthy();
    const settled = traceLog(mark).entries.filter((e) => /settle/i.test(e.message));
    expect(
      settled,
      "a pause was reintroduced after adding a slide — it cost 18 of 19 probe answers last time",
    ).toHaveLength(0);
  });
});

/**
 * What the host is allowed to tell us about a failure, and how much of it we keep.
 *
 * Every `debugInfo` in every real-host log this project owns ends the same way:
 * `"fullStatements":["Please enable config.extendedErrorLogging to see full
 * statements."]`. Without it, all a reader gets is `surroundingStatements` — a
 * pretty-printed excerpt — and on 2026-08-07 that excerpt was not enough to
 * decide the question the whole round turned on: whether the batch that failed
 * was `settleAndTagChart`'s own fresh one, or an older handle wearing its
 * shape. Two readings, no way to choose between them from the evidence, and one
 * of them says the settle is repairable while the other says it never ran.
 */
describe("asking the host what it actually ran", () => {
  it("turns on extended error logging, and says whether the host took it", () => {
    const cfg: { extendedErrorLogging?: boolean } = {};
    vi.stubGlobal("OfficeExtension", { config: cfg });
    expect(enableExtendedErrorLogging(), "asked and did not report success").toBe(true);
    expect(cfg.extendedErrorLogging, "never actually set the flag").toBe(true);
  });

  it("reports false rather than throwing on a host with no OfficeExtension", () => {
    // An older host, or the pane running outside Office at all. "The host would
    // not" and "nobody asked" are different diagnoses and the environment line
    // carries this so a reader can tell them apart.
    vi.stubGlobal("OfficeExtension", undefined);
    expect(enableExtendedErrorLogging()).toBe(false);
  });

  /**
   * Extended logging fills `fullStatements` with the WHOLE batch, and the whole
   * batch is not something a run log can carry: one round held 66 of these
   * errors, each one a 24-shape draw.
   *
   * BOTH ends, and the first version of this case pinned the wrong one. It kept
   * the tail, on the assumption that the failing statement is last — and in the
   * 2026-08-07 round the `>>>>>` marker sat on the FIRST statement of the batch
   * while the log read "… 37 earlier statement(s) dropped". The one line worth
   * reading was the one line thrown away. The head is also where the batch's
   * opening handles are, which is what settled whether a printed `getItem` means
   * a held handle.
   */
  it("keeps both ends of a long statement list and says how much it dropped", () => {
    const full = Array.from({ length: 200 }, (_, i) => `var shape${i} = shapes.getItem(...);`);
    const trimmed = trimDebugInfo({ code: "5010", fullStatements: full }) as {
      code: string;
      fullStatements: string[];
    };
    expect(trimmed.code, "dropped the rest of the debugInfo").toBe("5010");
    expect(trimmed.fullStatements.length, "kept the whole batch").toBeLessThan(full.length);
    // The batch's opening handles — where the failing statement sat in the round
    // that prompted this — and what it was doing when it stopped. Both.
    expect(trimmed.fullStatements[0], "dropped the head, where the marker was").toBe(full[0]);
    expect(trimmed.fullStatements.at(-1), "dropped the tail").toBe(full.at(-1));
    expect(trimmed.fullStatements.join("\n"), "never said what it dropped").toMatch(/dropped/);
  });

  it("leaves a short statement list exactly as it found it", () => {
    const info = { code: "5010", fullStatements: ["var shape = shapes.getItem(...);"] };
    expect(trimDebugInfo(info)).toBe(info);
  });

  it("carries the trimmed statements into the error text a run log records", () => {
    const full = Array.from({ length: 200 }, (_, i) => `var shape${i} = shapes.getItem(...);`);
    const text = errorText(
      Object.assign(new Error("InvalidParam passed to GetItem(id)"), {
        code: "5010",
        debugInfo: { errorLocation: "ShapeCollection.getItem", fullStatements: full },
      }),
    );
    expect(text, "the error text lost the statements entirely").toContain("ShapeCollection.getItem");
    expect(text).toContain("dropped");
    expect(text.length, "put a whole batch into one log line").toBeLessThan(4000);
  });
});

/**
 * A settle that found nothing says so.
 *
 * Three things end at `settled=0 lost=1` and want different fixes: a settle that
 * never ran, one that ran and found nothing, and one that ran and was refused.
 * On 2026-08-07 four of five settles ended on an empty collection read — the
 * same empty read that defeats the grouping — and the log did not say so once,
 * so working out what had actually happened meant reading the SHAPES of the
 * Office.js statements in the error payloads. That is not a diagnosis anyone
 * should have to make twice.
 */
describe("what the settle says when it comes up empty", () => {
  it("traces an empty re-read instead of returning silently", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    setTracing(true);
    const mark = traceMark();
    // TWO refusals, not one: the drawing context's write, and then the settle's
    // by-id write. With only the first, the settle's by-id attempt SUCCEEDS and
    // the collection read is never reached — the case this is about never runs.
    faults.refuseTagWrites = 2;
    faults.hollowReads = 50;
    try {
      const cfg = { ...sampleConfig("clustered"), ...DEFAULT_SIZE };
      await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg) });
      const said = traceLog(mark).entries.map((e) => e.message);
      expect(said, `the settle ended silently — said only: ${said.join(" | ")}`).toContain(
        "the settle's re-read came back empty",
      );
    } finally {
      faults.refuseTagWrites = 0;
      faults.hollowReads = 0;
    }
  });
});

/**
 * Every collection load names the properties it will read.
 *
 * Microsoft's own guidance, not a house style: `load("items")` loads the
 * COLLECTION and not the properties of the items in it — "you must explicitly
 * specify each property you need from collection items, as they won't be loaded
 * by default, including scalar properties". A property nobody asked for throws
 * `PropertyNotLoaded` at the READ, which is usually several lines and one sync
 * away from the load, and often inside a best-effort `try` that swallows it.
 *
 * The file had exactly one violation: the re-read before grouping, whose only
 * job is to read `id` off the items it gets back. That is also the read a real
 * host reported as `refreshed=0`.
 *
 * This guard was added, removed, and added again, which is worth recording. The
 * first attempt shipped without the docs behind it and was reverted when the
 * suite went red on `still gets its degraded pictures tagged` — trading a guess
 * for a measured regression. The research settled the guess, and the red test
 * turned out to be asserting the wrong property.
 */
describe("what a collection load asks the host for", () => {
  it("never loads a bare items — the property it will read has to be named", async () => {
    const { readFileSync } = await import("node:fs");
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    const bare = src.split("\n").flatMap((line, i) => (/\.load\(\s*["']items["']\s*\)/.test(line) ? [i + 1] : []));
    expect(
      bare,
      `bare load("items") at line(s) ${bare.join(", ")} — name the properties, as every other collection load here does`,
    ).toEqual([]);
  });
});

/**
 * A run log that leaves the reader a subtraction.
 *
 * Both of these cost real time on 2026-08-07. `read the deck back … tagged=38`
 * sat two lines from `tag pass … tagsFound=17` on the same 38-slide deck and
 * read as two passes disagreeing about how many charts were re-editable. They
 * were counting different tags: the first is DEMO_SLOT_TAG, the harness's own
 * bookkeeping, and it says nothing about a chart's config at all.
 *
 * And `tagsFound=14 slides=18` leaves four slides unaccounted for with no way to
 * settle what they are — four harness pages that never had a config, or four
 * charts that lost theirs. Opposite responses, same number.
 */
describe("what the repair pass's numbers actually count", () => {
  it("does not call the slot-tag count 'tagged'", async () => {
    const { readFileSync } = await import("node:fs");
    const src = readFileSync("src/render/powerpoint.ts", "utf8");
    const deckRead = src.slice(src.indexOf('trace("repair", "read the deck back"'));
    const body = deckRead.slice(0, deckRead.indexOf("});"));
    expect(body, "the deck read still reports a slot count under the name 'tagged'").not.toMatch(/\btagged:/);
    expect(body, "the slot count lost its name entirely").toMatch(/withSlotTag:/);
  });

  it("names the slides that came back without a config, not just the count", () => {
    const page = [{ index: 3 }, { index: 4 }, { index: 5 }];
    expect(untaggedIndices(page, new Set([4]))).toEqual([3, 5]);
    expect(untaggedIndices(page, new Set([3, 4, 5])), "invented a missing slide").toEqual([]);
  });

  it("caps the list so one bad page cannot bury the run log", () => {
    const page = Array.from({ length: 40 }, (_, i) => ({ index: i }));
    const named = untaggedIndices(page, new Set(), 12);
    expect(named).toHaveLength(13);
    expect(named[0]).toBe(0);
    // And it SAYS it truncated — a list that silently stops reads as a complete
    // answer, which is the same defect this whole case is about.
    expect(String(named.at(-1)), "truncated without saying so").toMatch(/28 more/);
  });
});

/**
 * A degraded item is ONE picture, and the reconcile has to expect one.
 *
 * `expected` came from `estimateOfficeShapes(scene)` unconditionally — what the
 * chart would be as native shapes. Past `degradedAt` the run draws a single
 * picture on purpose, so that asks the reconcile to compare a healthy chart
 * against a shape count nothing ever tried to put there, and the honest verdict
 * for that slide is `wreckage`.
 *
 * It was never SEEN as wreckage, and that is the uncomfortable part: on a host
 * that cannot count a group's children the comparison is short-circuited, so
 * this bug and the missing `measured` flag hid each other. A run on 2026-08-08
 * drew 36 of 38 slides as pictures and reported "35 of 38 complete" with
 * per-slide counts up to 253 — numbers describing a chart that was not there.
 */
describe("what the reconcile expects of a degraded item", () => {
  it("expects one shape per picture, not the whole scene's worth", async () => {
    const slides = [makeSlide("s1")];
    installHost(slides);
    const report = await insertDemoDeck(demoItems(6), undefined, {
      reconcile: true,
      shapeBudget: 1,
      pictureFor: async () => "data:image/png;base64,UE5H",
    });
    const from = report.degradedAt!;
    expect(from, "the budget did not force any degradation").toBeGreaterThan(0);
    const verdicts = report.reconcile?.plan.verdicts ?? [];
    expect(verdicts.length, "no reconcile ran").toBeGreaterThan(from);
    for (const v of verdicts.slice(from)) {
      expect(
        v.expected,
        `slot ${v.slot} (${v.title}) is one picture but the reconcile expected ${v.expected} shapes`,
      ).toBe(1);
    }
    // And the items BEFORE the switch still expect their real shape counts —
    // otherwise this "fix" would be switching the check off for everything.
    for (const v of verdicts.slice(0, from)) {
      expect(v.expected, `slot ${v.slot} drew native shapes; its expectation collapsed to 1`).toBeGreaterThan(1);
    }
  });
});

/**
 * Two index spaces, and the settle pass used to mix them.
 *
 * `updateChartsInSlides` filters `items` down to the ones whose slide and shape
 * the host still resolves, then maps over THAT. The settle plan paired each
 * result with `items[i]` — a different space the moment one item is dropped,
 * which is ordinary here: the user deletes a chart and runs Same Scale, or the
 * host declines one shape with a 5010.
 */
describe("the settle pass tags the chart it is holding", () => {
  const conf = (i: number) => ({ ...sampleConfig("clustered"), title: `chart-${i}` });

  it("writes each chart's own config, even when an earlier item was filtered out", async () => {
    const slide = makeSlide("s1");
    const kept = makeShape("geometric", "rectangle", { left: 10, top: 20, width: 100, height: 100 });
    kept.name = "PowerChart";
    kept.tagStore.set(CHART_TAG, JSON.stringify(conf(1)));
    slide.created.push(kept);
    installHost([slide], [kept], slide);
    // The config-tag write inside the drawing context is refused, so the settle
    // pass runs — the ordinary documented outcome on the web host.
    faults.refuseTagWrites = 1;

    await updateChartsInSlides([
      // Filtered out: the user deleted this chart between reading and pushing.
      {
        scene: buildChart(conf(0)),
        target: { slideId: "s1", shapeId: "the-user-deleted-this", left: 10, top: 20 },
        opts: { tagData: '{"i":0}' },
      },
      {
        scene: buildChart(conf(1)),
        target: { slideId: "s1", shapeId: kept.id, left: 10, top: 20 },
        opts: { tagData: '{"i":1}' },
      },
    ]);

    const written = slide.created
      .map((sh: { tagStore: Map<string, string> }) => sh.tagStore.get(CHART_TAG))
      .filter((v: string | undefined) => v === '{"i":0}' || v === '{"i":1}');
    // The surviving chart is chart 1. The old code wrote chart 0's config onto
    // it and then reported it re-editable, so opening it would have loaded the
    // DELETED chart's data and the next Update would have written that in.
    expect(written, "the settle tagged a chart with another chart's config").not.toContain('{"i":0}');
  });
});

/**
 * The guard `addAndRenderItem` has, and the sibling call site never got.
 */
describe("insertAgendaSlides when the host drops a slide add", () => {
  it("reports the host's reason, not a TypeError from renderer internals", async () => {
    installHost([makeSlide("s1")]);
    const chapters = ["Intro", "Body", "Close"];
    const scenes = chapters.map((_, i) => buildAgendaScene(chapters, { highlight: i }));
    // Every add swallowed, retry budget included: `addSlides` then hands back
    // fewer thunks than scenes, which is its documented contract.
    faults.swallowAdds = chapters.length * (1 + MAX_ADD_RETRY_ROUNDS) - 1;
    let err: unknown;
    try {
      await insertAgendaSlides(scenes);
    } catch (e) {
      err = e;
    }
    expect(err, "the dropped add went unreported").toBeDefined();
    expect(String((err as Error).message)).toMatch(/agenda slides/i);
    // "getSlide is not a function" is a TypeError from renderer internals, for
    // a condition `addSlides` diagnosed precisely one frame earlier.
    expect(String((err as Error).message)).not.toMatch(/is not a function/);
  });
});

/**
 * The demo/deck path drew at the OFF-SCREEN batch size (40) and then decided
 * whether the draw had spanned batches by comparing against the LIVE one (10).
 * So every chart of 11 to 40 shapes drew in a single batch — every proxy still
 * inside its own sync's window — and was told to go and re-read the collection
 * anyway.
 *
 * That is not a spare round trip on this host. It does not list the shapes a
 * run has just added, so the re-read comes back empty, and `chooseGroupMembers`
 * reads an empty answer to a refresh it ASKED for as "group nothing" — where not
 * asking would have grouped the perfectly good created proxies. The chart loses
 * its group, and with it the shape id the settle needs to write the config
 * through.
 */
describe("an off-screen chart that fitted in one batch is not sent to re-read the slide", () => {
  const midSizedScene = () => ({
    width: 480,
    height: 300,
    nodes: Array.from({ length: 24 }, (_, i) => ({
      kind: "rect" as const,
      x: i * 10,
      y: 10,
      w: 8,
      h: 40,
      fill: "#2a78d6",
      name: `seg-0-${i}`,
    })),
  });

  it("groups it even on a host that will not list a slide's shapes", async () => {
    const slides = [makeSlide("s1")];
    installHost(slides);
    applyWebProfile();
    // The refusal this is about, and only it: the collection answers nothing,
    // for good. `applyWebProfile` also refuses the first addGroup outright,
    // which is a different failure and would mask the one under test.
    faults.hollowReads = 99;
    faults.refuseGroups = 0;
    try {
      await insertDemoDeck([{ scene: midSizedScene(), tagData: '{"kind":"stacked"}' }]);
      const drawn = slides[slides.length - 1];
      expect(drawn.created.length, "nothing was drawn, so this proves nothing").toBeGreaterThan(10);
      expect(
        drawn.created.length,
        "24 shapes must fit one off-screen batch, or this tests something else",
      ).toBeLessThan(OFFSCREEN_BATCH);
      // The chart survived as a chart: grouped, and carrying its config.
      expect(
        drawn.created.some((s) => s.type === "group"),
        "the chart was left loose",
      ).toBe(true);
      expect(
        drawn.created.some((s) => s.tagStore.get(CHART_TAG) === '{"kind":"stacked"}'),
        "the chart lost its config tag",
      ).toBe(true);
    } finally {
      faults.hollowReads = 0;
    }
  });
});

/**
 * The settle's name search guarded "several charts on the slide" and not "the
 * one I found is somebody else's".
 *
 * Our chart is UNGROUPED — which is why it has no shape id, and why the name
 * search cannot see it. The group that IS on the slide belongs to a chart whose
 * own config landed perfectly well. The settle wrote our config over theirs and
 * reported a repair, so opening that chart hands back a different one's data.
 *
 * Asked of the rule directly, as `chooseGroupMembers` and
 * `targetWithNoTagResult` are: the end-to-end path resolves its shape by id and
 * never reaches this branch, so driving four simultaneous host failures through
 * the fake to arrive at it would test the fake's plumbing, not the rule.
 */
describe("the settle does not take a bystander chart's config", () => {
  it("writes through a shape it found BY ID, which is proof it is ours", () => {
    expect(mayTakeConfig({ foundById: true, hasConfig: undefined })).toBe(true);
    // Even one that already has a config: by id it IS our chart, and the config
    // it carries is the one this settle is replacing.
    expect(mayTakeConfig({ foundById: true, hasConfig: true })).toBe(true);
  });

  it("writes through a name match only when it carries no config", () => {
    expect(mayTakeConfig({ foundById: false, hasConfig: false })).toBe(true);
    expect(mayTakeConfig({ foundById: false, hasConfig: true })).toBe(false);
  });

  it("refuses a name match the host would not describe", () => {
    // An ungrouped chart carrying no config can be inserted again; a bystander
    // carrying the WRONG config means editing one chart rewrites another. So an
    // unreadable answer is refused rather than guessed.
    expect(mayTakeConfig({ foundById: false, hasConfig: undefined })).toBe(false);
  });
});

/**
 * A settle that lands puts the config back on the slide, so `no-config` should
 * go — and stripping it said nothing about whether the target it was clearing
 * still NAMED anything.
 *
 * For a chart whose new shapes were never read back, `shapeId` names the shape
 * that same call deleted. Handing it back bare is the trap the update path
 * documents in its own comments: the pane keeps it as the live edit target,
 * prints "Done." in green, and the next push resolves a dead id, is filtered out
 * as "the user deleted this chart", and tells them their chart is gone. The
 * settle SUCCEEDING made that worse — it was the one path that cleared the
 * marker protecting against it.
 */
describe("a settled chart whose shape was never named", () => {
  const target = { slideId: "s1", shapeId: "shape-25", left: 60, top: 90, lost: "no-config" as const };

  it("keeps a marker when the id it carries is the one this call deleted", () => {
    const back = afterSettle(target, { settled: true, untargeted: true });
    expect(back.lost, "handed back a live target pointing at a deleted shape").toBe("unknown-shape");
  });

  it("clears the marker when the target still names the new shape", () => {
    // The case the strip was written for, and it must keep working: the settle
    // landed and the target is real, so the chart is editable again.
    expect(afterSettle(target, { settled: true, untargeted: false }).lost).toBeUndefined();
  });

  it("leaves a chart the settle could not repair exactly as it was", () => {
    expect(afterSettle(target, { settled: false, untargeted: false }).lost).toBe("no-config");
    expect(afterSettle(target, { settled: false, untargeted: true }).lost).toBe("no-config");
  });

  it("does not touch a target that was never marked", () => {
    const fine = { slideId: "s1", shapeId: "shape-9", left: 0, top: 0 };
    expect(afterSettle(fine, { settled: true, untargeted: true })).toEqual(fine);
  });
});
