// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { installHost, makeSlide, applyWebProfile, faults, stallSyncOn } from "./helpers/office-host";
import { insertDemoDeck, DEMO_SLOT_TAG, CHART_TAG, _setBatchTimeoutForTest } from "../src/render/powerpoint";
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
