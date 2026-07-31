// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  _setBatchTimeoutForTest,
  _setBlankReReadDelayForTest,
  CHART_PARTS_TAG,
  CHART_TAG,
  getSelectionBounds,
  insertAgendaSlides,
  insertDemoDeck,
  insertSceneIntoSlide,
  isPowerPointHost,
  listChartsInDeck,
  listChartsInSelection,
  loadChartFromSelection,
  onLateSync,
  READBACK_PAGE,
  DEMO_SLOT_TAG,
  reconcileDeck,
  snapshotAddedSlides,
  updateChartInSlide,
  updateChartsInSlides,
} from "../src/render/powerpoint";
import { setTracing, traceLog } from "../src/core/trace";
import { buildChart, DEFAULT_SIZE } from "../src/core/chart";
import { sampleConfig } from "../src/core/samples";
import { buildAgendaScene } from "../src/core/agenda";
import type { ChartConfig, MarkerSymbol } from "../src/core/types";
import type { DemoReport } from "../src/render/powerpoint";

/** The indices a demo run did not render as a real chart (skipped or failed). */
const failedIndices = (r: DemoReport) =>
  r.results.map((x, i) => (x.status !== "rendered" ? i : -1)).filter((i) => i >= 0);

/**
 * Recording doubles for the PowerPoint JS proxy-object API: every shape the
 * renderer creates is captured with the geometry/format calls made on it, so
 * the whole scene→native-shapes mapping is testable without an Office host.
 */

let idSeq = 0;

function makeShape(
  type: string,
  geo: string | undefined,
  box: { left: number; top: number; width: number; height: number },
) {
  const tagStore = new Map<string, string>();
  const shape = {
    type,
    geo,
    box,
    fillColor: null as string | null,
    fillCleared: false,
    text: undefined as string | undefined,
    name: undefined as string | undefined,
    rotation: undefined as number | undefined,
    deleted: false,
    // The sync count at proxy creation — used to model the web host's
    // getItem(id) rewrite: a shape proxy is valid within the sync that queued
    // it plus the immediately following commit sync, and stale beyond that.
    // A hardened addGroup checks each member's age.
    syncCreated: trips.syncs,
    // A created shape's id exists only on the host: the renderer must load()
    // it back before it can write one down (see the parts tag).
    load() {},
    id: `shape-${++idSeq}`,
    left: box.left,
    top: box.top,
    width: box.width,
    height: box.height,
    tagStore,
    tags: {
      add: (k: string, v: string) => void tagStore.set(k, v),
      getItemOrNullObject: (k: string) => ({
        isNullObject: !tagStore.has(k),
        value: tagStore.get(k) ?? "",
        load() {},
        untrack() {
          untracked.tags++;
        },
      }),
    },
    // Set by fill.setImage — the bare base64 the host received. Records what the
    // renderer actually handed over, so a test can prove the data:/image-prefix
    // stripping happened rather than just that setImage was reached.
    imageBase64: undefined as string | undefined,
    fill: {
      setSolidColor(c: string) {
        shape.fillColor = c;
      },
      setImage(b64: string) {
        // Model the real API's one visible side effect: "This changes the fill
        // type to PictureAndTexture" (@types/office-js). `refusePictureFill`
        // makes the property access itself throw, which is what a pre-1.8 host
        // does — the method simply is not there.
        if (refusePictureFill) throw new Error("setImage is not available on this host");
        shape.imageBase64 = b64;
        shape.fillType = "PictureAndTexture";
      },
      clear() {
        shape.fillCleared = true;
      },
    },
    fillType: undefined as string | undefined,
    lineFormat: {} as Record<string, unknown>,
    textFrame: {
      textRange: { font: {} as Record<string, unknown>, paragraphFormat: {} as Record<string, unknown> },
    } as Record<string, unknown> & {
      textRange: { font: Record<string, unknown>; paragraphFormat: Record<string, unknown> };
    },
    grouped: undefined as unknown[] | undefined,
    // PowerPointApi 1.8's `Shape.group` — present only on an actual group, and
    // a throw otherwise, because that is what a repair pass has to survive: on
    // a host without 1.8 the property access poisons the sync it was queued in.
    get group() {
      if (!shape.grouped) throw new Error("shape is not a group");
      const live = (): { name?: string; deleted: boolean }[] =>
        (shape.grouped as { name?: string; deleted: boolean }[]).filter((c) => !c.deleted);
      return {
        shapes: {
          getCount: () => ({ value: live().length }),
          get items() {
            return live();
          },
          load() {},
        },
      };
    },
    delete() {
      shape.deleted = true;
      // Deleting a group deletes what it contains — otherwise an in-place update
      // leaves every child on the slide and the fake shows orphans the host
      // would never have.
      for (const child of (shape.grouped ?? []) as { deleted: boolean }[]) child.deleted = true;
    },
    untrack() {
      untracked.shapes++;
    },
  };
  return shape;
}

type FakeShape = ReturnType<typeof makeShape>;

/** Layout ids passed to slides.add() since the last installHost(). */
const addedWithLayout: (string | undefined)[] = [];

/**
 * The queued-command failure a stalled getItemAt handle produces, surfaced at
 * the NEXT sync — because Office.js reports queued-command errors only at sync,
 * never at the call that queued them. `installHost`'s sync() throws it. See
 * `freshWindowedHandle` below for why a reused fresh-slide handle poisons.
 */
let pendingHostError: Error | null = null;

/**
 * The slide count at the start of the current `PowerPoint.run` — every slide at
 * an index >= this was `add()`ed during this context and so is "fresh". A fresh
 * slide's getItemAt handle is only good within the sync it was acquired in (see
 * `freshWindowedHandle`); a pre-existing slide's handle is durable.
 */
let contextBaseCount = 0;

/** Set by installHost so a slide can splice itself out of the live deck. */
let deckRemove: ((s: { id: string }) => void) | null = null;

function makeSlide(id: string) {
  const created: FakeShape[] = [];
  // Shapes queued since the last successful sync. On success the pending list
  // is cleared (they're committed). On a failed sync the fake removes them
  // from `created` — the real host discards queued commands whose sync
  // rejected, and modeling that faithfully is what lets a mid-render
  // failure-mode test distinguish "40 batches committed, 41st batch dropped"
  // from "all 47 shapes on the slide" (the PR-8 partial-landed threshold).
  const pending: FakeShape[] = [];
  const slideTagStore = new Map<string, string>();
  const slide = {
    id,
    created,
    pending,
    isNullObject: false,
    load() {},
    // Slide.delete() — the repair pass removes duplicate slides, and a fake
    // that kept them would make every deletion assertion vacuous. installHost
    // owns the deck array, so removal goes through the hook it registers.
    delete() {
      deckRemove?.(slide);
    },
    tags: {
      add: (k: string, v: string) => void slideTagStore.set(k, v),
      getItemOrNullObject: (k: string) => ({
        isNullObject: !slideTagStore.has(k),
        value: slideTagStore.get(k) ?? "",
        load() {},
      }),
    },
    tagStore: slideTagStore,
    shapes: {
      // A deleted shape is gone from the host's collection; keeping it in `items`
      // let readbacks (listChartsInDeck) count a chart that no longer exists.
      // Accessing .items refreshes each shape's syncCreated: real Office.js
      // returns fresh proxies from shapes.items after a load+sync, valid in
      // the current sync — that's exactly what the PR 3 re-fetch relies on.
      get items() {
        const live = created.filter((s) => !s.deleted);
        for (const s of live) s.syncCreated = trips.syncs;
        return live;
      },
      load() {},
      addGeometricShape(geo: string, box: FakeShape["box"]) {
        const s = makeShape("geometric", geo, box);
        created.push(s);
        pending.push(s);
        return s;
      },
      addLine(_kind: string, box: FakeShape["box"]) {
        const s = makeShape("line", undefined, box);
        created.push(s);
        pending.push(s);
        return s;
      },
      addTextBox(text: string, box: FakeShape["box"]) {
        const s = makeShape("text", undefined, box);
        s.text = text;
        created.push(s);
        pending.push(s);
        return s;
      },
      addGroup(items: FakeShape[]) {
        // A host that refuses to group at all — the call throws where the API
        // exists but the host declines. Consumed per call, so a test can refuse
        // the render's own grouping and still let the fresh-context rescue's
        // addGroup work, which is the whole point of the rescue.
        if (refuseGroups > 0) {
          refuseGroups--;
          throw new Error("host refused addGroup");
        }
        // Model the web-host stale-proxy trap: a Shape proxy is valid within
        // the sync that queued it plus its immediately following commit sync,
        // and stale (getItem(id) rewrite, non-round-trippable id) beyond that.
        // addGroup(theseStaleProxies) silently loses grouping on real Office —
        // no group appears on the slide. Enable via strictGroup.
        if (strictGroup && items.some((s) => trips.syncs > s.syncCreated + 1)) {
          const g = makeShape("group", undefined, { left: 0, top: 0, width: 0, height: 0 });
          // Not pushed to `created` — no group lands on the slide, exactly
          // what the web host does for a stale-proxy addGroup call.
          return g;
        }
        // A real group's frame is the BOUNDING BOX of its children, not the
        // origin the caller drew at. Returning {0,0,0,0} made a group's position
        // meaningless and hid the update-drift bug entirely: the renderer reads
        // this back as an EditTarget and (before CHART_ORIGIN_TAG) fed it in as
        // the next render's frame origin, walking the chart down the slide.
        const box = items.length
          ? {
              left: Math.min(...items.map((s) => s.left)),
              top: Math.min(...items.map((s) => s.top)),
              width: Math.max(...items.map((s) => s.left + s.width)) - Math.min(...items.map((s) => s.left)),
              height: Math.max(...items.map((s) => s.top + s.height)) - Math.min(...items.map((s) => s.top)),
            }
          : { left: 0, top: 0, width: 0, height: 0 };
        const g = makeShape("group", undefined, box);
        g.grouped = items;
        created.push(g);
        pending.push(g);
        return g;
      },
      getItemOrNullObject(id: string) {
        // A null-object proxy still accepts load() in real Office.js — queuing a
        // load on it is the normal pattern, and only isNullObject tells you it
        // resolved to nothing. Without load() here the fake threw where the host
        // would not.
        return created.find((s) => s.id === id && !s.deleted) ?? { isNullObject: true, load() {}, delete() {} };
      },
      // Top-level shape count the host reports on readback — non-deleted shapes.
      getCount: () => {
        if (faultShapeGetCount) throw new Error("readback getCount faulted");
        return { value: created.filter((s) => !s.deleted).length };
      },
    },
  };
  return slide;
}

type FakeSlide = ReturnType<typeof makeSlide>;

/**
 * The reference `slides.getItemAt(i)` hands back for a FRESHLY-ADDED slide.
 *
 * It draws fine within the sync it was acquired in, but reusing it to draw AFTER
 * a later sync is the object-path rewrite trap: Office.js has by then rewritten
 * its path to `getItem(<web-non-round-trippable id>)`, so the next shape throws
 * "InvalidParam passed to GetItem(id)" (code 5010) at the following sync. That is
 * exactly why the fix re-acquires a brand-new getItemAt proxy every batch — a
 * fresh handle is always inside its own window — and why HOLDING one across
 * batches fails. When poisoned, nothing lands: the queued shapes are detached,
 * not pushed to the real slide's `created`.
 *
 * The old fake could not express this at all: it returned the live slide from
 * getItemAt, so a held handle was as good as a fresh one.
 */
function freshWindowedHandle(real: FakeSlide) {
  const acquiredSync = trips.syncs;
  // Valid only until the next sync moves past the window it was acquired in.
  const ok = () => {
    if (trips.syncs <= acquiredSync) return true;
    pendingHostError = new Error(
      'InvalidParam passed to GetItem(id) | code=5010 | debugInfo={"errorLocation":"SlideCollection.getItem"}',
    );
    return false;
  };
  return {
    id: real.id,
    isNullObject: false,
    load() {},
    tags: real.tags,
    shapes: {
      // Lazy — evaluating this at handle-creation would fire the .items
      // getter's syncCreated refresh, keeping every shape perpetually fresh
      // and hiding the stale-proxy trap the fake exists to model.
      get items() {
        return real.shapes.items;
      },
      load() {},
      addGeometricShape: (geo: string, box: FakeShape["box"]) =>
        ok() ? real.shapes.addGeometricShape(geo, box) : makeShape("geometric", geo, box),
      addLine: (kind: string, box: FakeShape["box"]) =>
        ok() ? real.shapes.addLine(kind, box) : makeShape("line", undefined, box),
      addTextBox: (text: string, box: FakeShape["box"]) => {
        if (ok()) return real.shapes.addTextBox(text, box);
        const s = makeShape("text", undefined, box);
        s.text = text;
        return s;
      },
      addGroup: (items: FakeShape[]) =>
        ok() ? real.shapes.addGroup(items) : makeShape("group", undefined, { left: 0, top: 0, width: 0, height: 0 }),
      getItemOrNullObject: (id: string) => real.shapes.getItemOrNullObject(id),
    },
  };
}

/**
 * Office round-trips since the last installHost(). Every context.sync() is a
 * trip to PowerPoint and dominates insert latency, so the count is a behaviour
 * worth asserting — see "round-trips do not scale with the chart count".
 */
const trips = { syncs: 0, contexts: 0 };

/** Proxy objects the renderer released via untrack(), by kind. */
const untracked = { shapes: 0, tags: 0 };

/**
 * Make the Nth context.sync() of the next run throw. Office.js queues commands
 * and only reports their errors at sync — so this, not a throwing addGroup, is
 * how a host actually refuses something. 0 = never.
 */
let failSyncOn = 0;

/** Like failSyncOn but a SET of sync indices — models a stall that persists across
 * the retry (fail both attempts' syncs), so a slide is truly lost, not recovered. */
const failSyncsOn = new Set<number>();

/** Make the next N slides.add() calls no-ops — models PowerPoint web silently
 * dropping a slide-add under load, the corruption the self-check must catch. */
let swallowAdds = 0;

/** Deck indices whose shapes.getCount() reports 0 on readback — models a slide
 * that committed but came back BLANK (its shapes detached), the silent partial
 * the on-slide readback must catch. */
const blankReadbackAt = new Set<number>();

/** When true, shapes.getCount() throws — models the blank readback faulting, so
 * the report must come back blanksRead:false rather than an empty "no blanks". */
let faultShapeGetCount = false;

/** Sync indices that STALL — the promise resolves after `stallSyncDelayMs`, but
 * withTimeout (shortened via _setBatchTimeoutForTest) fires first and abandons
 * it. Models the real-host bug: sync at 60s, shapes on the slide, no error to
 * blame. When the abandoned promise later settles, withTimeout records
 * `lastLateSync = "…SUCCEEDED after Ns"` — the signal PR 1 gates on. */
const stallSyncOn = new Set<number>();
let stallSyncDelayMs = 40;

/** When true, addGroup silently drops the group if any member proxy is more
 * than one sync stale — the exact web-host behavior that made multi-batch
 * charts land ungrouped. See PR 3's re-fetch fix.  */
let strictGroup = false;

/** Make the next N addGroup calls THROW — a host that declines to group even
 * though the API is there. Decremented per call, so refusing 1 leaves a chart
 * rendered-but-ungrouped and still lets the fresh-context rescue group it. */
let refuseGroups = 0;

/** When true, `fill.setImage` throws on access — a pre-1.8 host, where the
 * method does not exist at all. Drives the picture-insert fall-through. */
let refusePictureFill = false;

/** Install a fake PowerPoint global whose run() drives the mocked context.
 * `supported(version)` models the host's requirement-set support (default: all)
 * — pass a predicate to simulate e.g. PowerPoint on the web lacking grouping. */
function installHost(
  slides: FakeSlide[],
  selectedShapes: FakeShape[] = [],
  selectedSlide = slides[0],
  supported: (version: string) => boolean = () => true,
) {
  // The slide count as of the last COMMITTED sync. getCount() reports THIS, not
  // the live array — so an add() queued in the current batch is invisible to a
  // getCount() in the SAME batch, exactly as PowerPoint web behaves. A getCount
  // result resolves at the next sync to the count from before that sync's adds.
  let committedCount = slides.length;
  deckRemove = (s) => {
    const i = slides.findIndex((x) => x.id === s.id);
    if (i >= 0) slides.splice(i, 1);
  };
  const pendingCounts: { value: number }[] = [];
  const context = {
    presentation: {
      slides: {
        items: slides,
        load() {},
        getItem: (id: string) => slides.find((s) => s.id === id)!,
        // Real Office.js hands back a null OBJECT for an unknown id — it does
        // not throw and does not return undefined. A fake that returns
        // undefined would make `slide.isNullObject` a TypeError instead of the
        // false it should be, and hide the very case this models. By id, the
        // reference is always durable.
        getItemOrNullObject: (id: string) => slides.find((s) => s.id === id) ?? { isNullObject: true, load() {} },
        // A pre-existing slide's handle is durable; a freshly-added one's is only
        // good within the sync it was acquired in (see freshWindowedHandle), so
        // HOLDING one across the render's batches is the bug the fix avoids by
        // re-acquiring fresh each batch.
        getItemAt: (i: number) => {
          const s = slides[i];
          if (!s) return s;
          if (i >= contextBaseCount) return freshWindowedHandle(s);
          // A durable slide the host reports empty on readback (see blankReadbackAt).
          if (blankReadbackAt.has(i)) return { ...s, shapes: { ...s.shapes, getCount: () => ({ value: 0 }) } };
          return s;
        },
        // Resolves at the NEXT sync to the committed count (from before that
        // sync's adds), never to slides.length now — see committedCount.
        getCount: () => {
          const result = { value: committedCount };
          pendingCounts.push(result);
          return result;
        },
        add: (options?: { layoutId?: string }) => {
          addedWithLayout.push(options?.layoutId);
          if (swallowAdds > 0) {
            swallowAdds--; // the host dropped this add — no slide appears
            return;
          }
          slides.push(makeSlide(`slide-${slides.length + 1}`));
        },
      },
      // A real deck's master carries several layouts; only one is blank, and
      // its NAME is localised — which is why the renderer matches on type.
      slideMasters: {
        items: [
          {
            id: "master-1",
            layouts: {
              items: [
                { id: "layout-title", name: "Titeldias", type: "titleSlide" },
                { id: "layout-blank", name: "Tom", type: "blank" },
                { id: "layout-content", name: "Titel og indhold", type: "object" },
              ],
            },
          },
        ],
        load() {},
      },
      getSelectedSlides: () => ({ getItemAt: () => selectedSlide }),
      getSelectedShapes: () => ({ items: selectedShapes, load() {} }),
    },
    sync: async () => {
      trips.syncs++;
      // Commit or discard the shapes each slide has queued since the last
      // sync. On success, mark them "committed" (leave in `created`, drop
      // from `pending`). On failure, remove them from `created` — the real
      // host discards a batch of queued commands whose sync rejects, and the
      // PR-8 partial-landed gate reads back only the committed count.
      const commit = () => {
        for (const s of slides) s.pending.length = 0;
      };
      const discard = () => {
        for (const s of slides) {
          for (const p of s.pending) {
            const i = s.created.indexOf(p);
            if (i >= 0) s.created.splice(i, 1);
          }
          s.pending.length = 0;
        }
      };
      if (trips.syncs === failSyncOn || failSyncsOn.has(trips.syncs)) {
        discard();
        throw new Error("host refused a queued command");
      }
      if (stallSyncOn.has(trips.syncs)) {
        // Sleep past withTimeout's deadline, then settle successfully. The
        // queued shapes commit at settle time — same as real Office.js where
        // a slow sync eventually reports success and the shapes are on the
        // slide by then.
        await new Promise((r) => setTimeout(r, stallSyncDelayMs));
        for (const r of pendingCounts) r.value = committedCount;
        pendingCounts.length = 0;
        committedCount = slides.length;
        commit();
        return;
      }
      // A queued-command failure (e.g. drawing on a poisoned getItemAt handle)
      // surfaces here, at the sync, exactly as Office.js reports it.
      if (pendingHostError) {
        const err = pendingHostError;
        pendingHostError = null;
        discard();
        throw err;
      }
      // Each getCount from this batch resolves to the PRE-batch committed count,
      // then this batch's adds become visible to the NEXT sync's getCount.
      for (const r of pendingCounts) r.value = committedCount;
      pendingCounts.length = 0;
      committedCount = slides.length;
      commit();
    },
  };
  trips.syncs = 0;
  trips.contexts = 0;
  untracked.shapes = 0;
  untracked.tags = 0;
  pendingHostError = null;
  swallowAdds = 0;
  failSyncsOn.clear();
  stallSyncOn.clear();
  strictGroup = false;
  refuseGroups = 0;
  refusePictureFill = false;
  blankReadbackAt.clear();
  faultShapeGetCount = false;
  addedWithLayout.length = 0;
  vi.stubGlobal("PowerPoint", {
    run: async <T>(cb: (ctx: typeof context) => Promise<T>) => {
      trips.contexts++;
      // Slides present at the start of this context are "existing"; anything
      // add()ed past here is fresh, and its getItemAt handle is window-limited.
      contextBaseCount = slides.length;
      return cb(context);
    },
    // Real Office.js exposes all 177 presets. A plain object listing only the
    // ones in use today hands back `undefined` for any other name, and the
    // renderer then records a shape with no geometry while this suite still
    // passes green — a test that asserts nothing about the shape it drew.
    // The Proxy makes that impossible: reaching for a preset this stub has not
    // been told about throws instead of returning undefined.
    GeometricShapeType: new Proxy(
      {
        rectangle: "rectangle",
        ellipse: "ellipse",
        triangle: "triangle",
        chevron: "chevron",
        homePlate: "homePlate",
        lineInverse: "lineInverse",
        diamond: "diamond",
        plus: "plus",
      } as Record<string, string>,
      {
        get(target, prop: string) {
          if (!(prop in target)) throw new Error(`office stub: unknown GeometricShapeType "${String(prop)}"`);
          return target[prop];
        },
      },
    ),
    SlideLayoutType: { blank: "blank", titleSlide: "titleSlide", object: "object" },
    ConnectorType: { straight: "straight" },
    ShapeLineDashStyle: { dash: "dash", roundDot: "roundDot" },
    ShapeAutoSize: { autoSizeNone: "none" },
    TextVerticalAlignment: { top: "top", middle: "middle", bottom: "bottom" },
    ParagraphHorizontalAlignment: { left: "left", center: "center", right: "right" },
  });
  vi.stubGlobal("Office", {
    context: {
      host: "PowerPoint",
      requirements: { isSetSupported: (_set: string, version: string) => supported(version) },
    },
  });
  return context;
}

const config: ChartConfig = {
  kind: "stacked",
  ...DEFAULT_SIZE,
  data: {
    categories: ["A", "B"],
    series: [
      { name: "S1", values: [3, 4] },
      { name: "S2", values: [1, 2] },
    ],
  },
};

afterEach(() => vi.unstubAllGlobals());

describe("insertSceneIntoSlide", () => {
  it("creates native shapes at the requested offset, groups, and tags", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { left: 100, top: 50, tagData: JSON.stringify(config) });

    const rects = slide.created.filter((s) => s.geo === "rectangle");
    expect(rects.length).toBeGreaterThanOrEqual(4); // one per stacked segment
    for (const r of rects) expect(r.box.left).toBeGreaterThanOrEqual(100);
    const group = slide.created.find((s) => s.type === "group")!;
    expect(group.name).toBe("PowerChart");
    expect(group.grouped).toHaveLength(slide.created.length - 1);
    expect(group.tagStore.get(CHART_TAG)).toBe(JSON.stringify(config));
  });

  it("describes the chart group with accessible alt text", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: JSON.stringify(config) });
    const group = slide.created.find((s) => s.type === "group") as FakeShape & {
      altTextDescription?: string;
      altTextTitle?: string;
    };
    expect(group.altTextDescription).toBe(scene.desc);
    if (scene.title) expect(group.altTextTitle).toBe(scene.title);
  });

  it("describes the chart even when it is NOT grouped", async () => {
    // The alt text used to be assigned only on the group object, so every
    // ungrouped chart (group:false, a refused addGroup, a one-shape chart)
    // silently lost its text alternative. It belongs on whatever shape stands
    // for the chart — the same one the config tag lands on.
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: "cfg", group: false });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    const anchor = slide.created[0] as FakeShape & { altTextDescription?: string; altTextTitle?: string };
    expect(anchor.altTextDescription).toBe(scene.desc);
    expect(anchor.tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("skips alt text on a host below 1.10 instead of losing the group with it", async () => {
    // Shape.altTextDescription is PowerPointApi 1.10. Assigning it on an older
    // host is a queued command rejected at the next sync — the same sync that
    // carries the grouping, so an ungated assignment costs the group.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.10");
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group") as FakeShape & { altTextDescription?: string };
    expect(group).toBeDefined();
    expect(group.altTextDescription).toBeUndefined();
  });

  it("renders a pie as a rotated triangle fan", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = buildChart({
      ...config,
      kind: "pie",
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
    });
    await insertSceneIntoSlide(scene, {});
    const tris = slide.created.filter((s) => s.geo === "triangle" && s.name?.includes("-f"));
    expect(tris.length).toBeGreaterThan(10);
    for (const t of tris) expect(typeof t.rotation).toBe("number");
  });

  it("maps title font and alignment onto text boxes", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart({ ...config, title: "Hello" }), { fontFamily: "Arial" });
    const title = slide.created.find((s) => s.text === "Hello")!;
    expect(title.fillCleared).toBe(true);
    expect(title.textFrame.textRange.font).toMatchObject({ name: "Arial", bold: true });
  });

  it("draws value lines as dashed native connectors", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(
      buildChart({ ...config, decorations: { valueLines: [{ mode: "mean" }], segmentLabels: true } }),
      {},
    );
    const dashed = slide.created.filter((s) => s.type === "line" && s.lineFormat.dashStyle === "dash");
    expect(dashed.length).toBeGreaterThanOrEqual(1);
  });
});

describe("scene node mapping", () => {
  const insert = async (nodes: object[], opts = {}) => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide({ width: 200, height: 100, nodes } as never, opts);
    return slide;
  };

  it("maps ellipses with stroke or hidden outline", async () => {
    const slide = await insert([
      {
        kind: "ellipse",
        cx: 50,
        cy: 50,
        rx: 20,
        ry: 10,
        fill: "#ff0000",
        stroke: "#000000",
        strokeWidth: 2,
        name: "dot",
      },
      { kind: "ellipse", cx: 10, cy: 10, rx: 5, ry: 5, fill: "#00ff00" },
    ]);
    const [a, b] = slide.created.filter((s) => s.geo === "ellipse");
    // center − radius, plus the default 60/90pt insert offset
    expect(a.box).toEqual({ left: 90, top: 130, width: 40, height: 20 });
    expect(a.lineFormat).toMatchObject({ color: "#000000", weight: 2 });
    expect(b.lineFormat.visible).toBe(false);
  });

  it("honours an 8-digit #RRGGBBAA fill: 6-digit hue + transparency, never mis-parsed", async () => {
    // #RRGGBBAA is a valid hand-authored colour that the SVG preview and the
    // skill's pptx render translucent. Office.js setSolidColor validates 6-digit
    // hex only, so the alpha byte has to move to fill.transparency (1.4) or the
    // live add-in would mis-parse the value and lose the colour.
    const slide = await insert([{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#4e79a780", name: "band" }]);
    const rect = slide.created.find((s) => s.geo === "rectangle") as (typeof slide.created)[number] & {
      fill: { transparency?: number };
    };
    expect(rect.fillColor).toBe("#4e79a7"); // hue survives, no 8-digit string reaches Office.js
    expect(rect.fill.transparency).toBeCloseTo(1 - 128 / 255, 3);
  });

  it("leaves an opaque 6-digit fill untouched (no transparency set)", async () => {
    const slide = await insert([{ kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#4e79a7", name: "bar" }]);
    const rect = slide.created.find((s) => s.geo === "rectangle") as (typeof slide.created)[number] & {
      fill: { transparency?: number };
    };
    expect(rect.fillColor).toBe("#4e79a7");
    expect(rect.fill.transparency).toBeUndefined();
  });

  it("renders a fill:'none' rect as an outlined/hollow shape (IBCS plan columns)", async () => {
    const slide = await insert([
      { kind: "rect", x: 0, y: 0, w: 20, h: 40, fill: "none", stroke: "#3b6ea5", strokeWidth: 1.5, name: "pl" },
    ]);
    const rect = slide.created.find((s) => s.geo === "rectangle")!;
    expect(rect.fillCleared).toBe(true); // no fill
    expect(rect.fillColor).toBeNull(); // never mis-parsed as a colour
    expect(rect.lineFormat).toMatchObject({ color: "#3b6ea5", weight: 1.5 });
  });

  it("maps chevrons to chevron/homePlate geometry", async () => {
    const slide = await insert([
      { kind: "chevron", x: 0, y: 0, w: 40, h: 20, fill: "#123456", flatLeft: true },
      { kind: "chevron", x: 50, y: 0, w: 40, h: 20, fill: "#123456" },
    ]);
    expect(slide.created.filter((s) => s.type !== "group").map((s) => s.geo)).toEqual(["homePlate", "chevron"]);
  });

  it("draws axis-aligned lines with a clamped non-zero box (never a degenerate diagonal)", async () => {
    const slide = await insert([
      {
        kind: "line",
        x1: 10,
        y1: 50,
        x2: 200,
        y2: 50,
        stroke: "#333333",
        strokeWidth: 1,
        dash: [3, 2],
        name: "connector",
      },
    ]);
    const line = slide.created.find((s) => s.type === "line")!;
    // Horizontal line: width spans, height is clamped up from 0 so the web host
    // can't blow a zero-thickness box into a giant diagonal.
    expect(line.box.width).toBeGreaterThan(180);
    expect(line.box.height).toBeGreaterThanOrEqual(0.5);
    expect(line.lineFormat.dashStyle).toBe("dash");
  });

  it("honours a translucent stroke on BOTH the line and the rotated-rect branch", async () => {
    // The two branches of addSegment diverged: the axis-aligned/dashed line
    // dropped the alpha byte while the diagonal rotated-rect folded it into
    // transparency, so one series colour rendered at two opacities. Both must now
    // agree — a 50% alpha (#…80) → transparency ≈ 0.498, colour a bare 6-digit hex.
    const alpha = "#33333380";
    const horiz = await insert([{ kind: "line", x1: 10, y1: 50, x2: 200, y2: 50, stroke: alpha, name: "h" }]);
    const line = horiz.created.find((s) => s.type === "line")!;
    expect(line.lineFormat.color).toBe("#333333"); // alpha byte stripped off the colour
    expect(line.lineFormat.transparency).toBeCloseTo(1 - 0x80 / 255, 4); // …but carried here

    const diag = await insert([{ kind: "line", x1: 0, y1: 0, x2: 100, y2: 100, stroke: alpha, name: "d" }]);
    const rect = diag.created.find((s) => s.geo === "rectangle")!;
    expect(rect.fillColor).toBe("#333333");
    expect((rect.fill as unknown as { transparency?: number }).transparency).toBeCloseTo(1 - 0x80 / 255, 4);
  });

  it("draws diagonal lines as thin rotated rectangles (direction-correct on every host)", async () => {
    // Up-right and down-right diagonals a bounding box alone can't distinguish.
    const down = await insert([
      { kind: "line", x1: 0, y1: 0, x2: 100, y2: 100, stroke: "#a00000", strokeWidth: 2, name: "d" },
    ]);
    const dr = down.created.find((s) => s.geo === "rectangle")!;
    expect(dr).toBeTruthy();
    expect(dr.fillColor).toBe("#a00000");
    expect(dr.rotation).toBeCloseTo(45, 0); // down-right
    expect(down.created.some((s) => s.type === "line")).toBe(false);

    const up = await insert([
      { kind: "line", x1: 0, y1: 100, x2: 100, y2: 0, stroke: "#00a000", strokeWidth: 2, name: "u" },
    ]);
    const ur = up.created.find((s) => s.geo === "rectangle")!;
    expect(ur.rotation).toBeCloseTo(-45, 0); // up-right — the case a box would mirror
  });

  it("draws dashed diagonals as real line shapes, picking the geometry per direction", async () => {
    // A rotated rectangle carries its colour in its fill, which can't be
    // dashed — scatter trend lines and forecast segments came out solid.
    const down = await insert([
      {
        kind: "line",
        x1: 0,
        y1: 0,
        x2: 100,
        y2: 60,
        stroke: "#a00000",
        strokeWidth: 1.25,
        dash: [4, 2],
        name: "trend",
      },
    ]);
    const dl = down.created.find((s) => s.name === "trend")!;
    expect(dl.type).toBe("line"); // not a filled rectangle
    expect(dl.lineFormat.dashStyle).toBe("dash");
    expect(dl.box).toMatchObject({ width: 100, height: 60 });

    // Up-right: addLine only ever draws the box's top-left→bottom-right
    // diagonal, so this direction needs the lineInverse geometry.
    const up = await insert([
      {
        kind: "line",
        x1: 0,
        y1: 60,
        x2: 100,
        y2: 0,
        stroke: "#a00000",
        strokeWidth: 1.25,
        dash: [4, 2],
        name: "trend",
      },
    ]);
    const ul = up.created.find((s) => s.name === "trend")!;
    expect(ul.geo).toBe("lineInverse");
    expect(ul.lineFormat.dashStyle).toBe("dash");
  });

  it("renders a dotted array as roundDot, not a generic dash", async () => {
    // [1.5,1.5] is the dotted waterfall carry connector. It used to flatten to
    // the same dash enum as every other pattern, so the deck lost the dotted
    // look the SVG preview shows.
    const slide = await insert([
      {
        kind: "line",
        x1: 10,
        y1: 50,
        x2: 200,
        y2: 50,
        stroke: "#333",
        strokeWidth: 1,
        dash: [1.5, 1.5],
        name: "carry",
      },
    ]);
    const line = slide.created.find((s) => s.name === "carry")!;
    expect(line.lineFormat.dashStyle).toBe("roundDot");
  });

  it("draws polygon edges direction-correct, with no zero-thickness boxes", async () => {
    // A violin body: an up-right edge, a horizontal edge and a down-right edge.
    const slide = await insert([
      {
        kind: "polygon",
        points: [
          { x: 0, y: 40 },
          { x: 50, y: 0 },
          { x: 100, y: 40 },
          { x: 100, y: 40 },
        ],
        fill: "#eeeeee",
        stroke: "#3366cc",
        strokeWidth: 1,
        name: "violin-0",
      },
    ]);
    const edges = slide.created.filter((s) => s.name?.startsWith("violin-0-e"));
    expect(edges).toHaveLength(4);
    for (const e of edges) {
      // Every edge is a real segment: a bounding box collapsed to zero on one
      // axis let the web host blow it up into a giant diagonal.
      expect(e.box.width).toBeGreaterThanOrEqual(0.5);
      expect(e.box.height).toBeGreaterThanOrEqual(0.5);
    }
    // Edge 0 (0,40)->(50,0) rises to the right; edge 1 (50,0)->(100,40) falls.
    // Passing both bounding boxes to addLine drew them as the same diagonal.
    const [e0, e1] = edges;
    expect(e0.rotation).toBeLessThan(0);
    expect(e1.rotation).toBeGreaterThan(0);
  });

  it("maps arrowheads to rotated triangles anchored at the tip", async () => {
    const slide = await insert([{ kind: "arrowhead", x: 10, y: 10, size: 4, angle: 45, fill: "#000000", name: "ah" }]);
    const tri = slide.created[0];
    expect(tri.geo).toBe("triangle");
    expect(tri.rotation).toBe(135); // scene angle + 90
    // The triangle's tip (box top-centre, rotated θ about the box centre) must
    // land on the scene point (10,10) + the default 60/90pt insert offset.
    const s = 8; // size * 2
    const theta = (tri.rotation! * Math.PI) / 180;
    const cx = tri.box.left + s / 2;
    const cy = tri.box.top + s / 2;
    const tipX = cx + (s / 2) * Math.sin(theta);
    const tipY = cy - (s / 2) * Math.cos(theta);
    expect(tipX).toBeCloseTo(70, 4); // 60 + 10
    expect(tipY).toBeCloseTo(100, 4); // 90 + 10
  });

  it("renders an annular wedge (sunburst ring / gauge) as a rotated rectangle band", async () => {
    const slide = await insert([
      {
        kind: "wedge",
        cx: 50,
        cy: 50,
        r: 30,
        innerR: 15,
        startAngle: 0,
        endAngle: 90,
        fill: "#333333",
        stroke: "#ffffff",
        strokeWidth: 1,
        name: "ring",
      },
    ]);
    const band = slide.created.filter((s) => s.geo === "rectangle" && s.name?.includes("-f"));
    expect(band.length).toBeGreaterThan(2); // the annular band, not a triangle fan
    for (const b of band) {
      expect(b.fillColor).toBe("#333333");
      expect(typeof b.rotation).toBe("number");
    }
    // No triangles for an annular wedge (a triangle can't leave a hole).
    expect(slide.created.some((s) => s.geo === "triangle")).toBe(false);
    // Two radial separators in the stroke colour.
    const edges = slide.created.filter((s) => s.name === "ring-edge");
    expect(edges.length).toBe(2);
    for (const e of edges) expect(e.fillColor).toBe("#ffffff");
  });

  it("skips grouping when group:false or only one shape", async () => {
    const slide = await insert(
      [
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#111111" },
        { kind: "rect", x: 20, y: 0, w: 10, h: 10, fill: "#222222" },
      ],
      { group: false, tagData: "cfg" },
    );
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    // The tag falls back onto the first created shape.
    expect(slide.created[0].tagStore.get(CHART_TAG)).toBe("cfg");
  });

  // Shape.rotation is PowerPointApi 1.10 and the manifests admit hosts from 1.4.
  // A try/catch around the assignment catches NOTHING on a real host: Office.js
  // proxy setters do not throw synchronously — the host rejects the queued
  // command at the next context.sync(), which carries the whole batch. So the
  // fake here rejects at SYNC, the way PowerPoint does, not at the setter.
  it.each(["pie", "doughnut", "sunburst"])(
    "%s still inserts on a pre-1.10 host (rotation gated, not wrapped)",
    async (kind) => {
      const slide = makeSlide("s-old");
      const ctx = installHost([slide], [], slide, (v) => v !== "1.10");
      const realSync = ctx.sync;
      ctx.sync = async () => {
        // Any rotation assigned on a host without 1.10 poisons the whole batch.
        if (slide.created.some((sh) => sh.rotation !== undefined)) {
          throw new Error("PropertyNotSupported: Shape.rotation requires PowerPointApi 1.10");
        }
        return realSync();
      };
      const scene = buildChart(sampleConfig(kind as never));
      // Must not reject: the chart degrades (no wedges) instead of failing.
      let failure: unknown = null;
      try {
        await insertSceneIntoSlide(scene, { tagData: "{}" });
      } catch (e) {
        failure = e;
      }
      expect(failure, `insert rejected on a pre-1.10 host: ${String(failure)}`).toBeNull();
      // The rest of the chart — labels, leader lines — still lands on the slide.
      expect(slide.created.length).toBeGreaterThan(0);
      // …and nothing carries a rotation the host cannot accept.
      expect(slide.created.every((sh) => sh.rotation === undefined)).toBe(true);
    },
  );

  it("degrades gracefully when the host lacks grouping and rotation", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    // Break addGroup and rotation assignment the way an old host would.
    slide.shapes.addGroup = () => {
      throw new Error("addGroup requires PowerPointApi 1.8");
    };
    const scene = {
      width: 200,
      height: 100,
      nodes: [
        { kind: "rect", x: 0, y: 0, w: 10, h: 10, fill: "#111111" },
        { kind: "wedge", cx: 50, cy: 50, r: 30, innerR: 0, startAngle: 0, endAngle: 90, fill: "#333333", name: "w" },
      ],
    };
    const realAdd = slide.shapes.addGeometricShape.bind(slide.shapes);
    slide.shapes.addGeometricShape = (geo, box) => {
      const s = realAdd(geo, box);
      if (geo === "triangle") {
        Object.defineProperty(s, "rotation", {
          set() {
            throw new Error("rotation requires PowerPointApi 1.10");
          },
        });
      }
      return s;
    };
    await insertSceneIntoSlide(scene as never, { tagData: "cfg" });
    // No group, no fan triangles survive — but the rect is inserted and tagged.
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(slide.created[0].tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("still inserts (ungrouped) when the host lacks grouping — the web case", async () => {
    // PowerPoint on the web: grouping (1.8) unsupported, tags (1.3) supported.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.8");
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    // The shapes are committed and no grouping was attempted…
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(slide.created.filter((s) => s.geo === "rectangle").length).toBeGreaterThanOrEqual(4);
    // …and the config tag lands on the first shape, so the chart is re-editable.
    expect(slide.created[0].tagStore.get(CHART_TAG)).toBe("cfg");
  });

  it("skips tagging when the host lacks tags", async () => {
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => false); // nothing supported
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    expect(slide.created[0].tagStore.get(CHART_TAG)).toBeUndefined();
    expect(slide.created.length).toBeGreaterThan(0);
  });

  it("falls back to the first slide when nothing is selected", async () => {
    const slide = makeSlide("s1");
    const ctx = installHost([slide]);
    ctx.presentation.getSelectedSlides = () => {
      throw new Error("no selection");
    };
    await insertSceneIntoSlide(
      { width: 10, height: 10, nodes: [{ kind: "rect", x: 0, y: 0, w: 5, h: 5, fill: "#111111" }] } as never,
      {},
    );
    expect(slide.created).toHaveLength(1);
  });
});

describe("updateChartInSlide", () => {
  it("deletes the old group and re-renders at the same position", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "x" });
    const oldGroup = slide.created.find((s) => s.type === "group")!;
    const before = slide.created.length;
    await updateChartInSlide(buildChart(config), { slideId: "s1", shapeId: oldGroup.id, left: 33, top: 44 });
    expect(oldGroup.deleted).toBe(true);
    const fresh = slide.created.slice(before).filter((s) => s.type !== "group");
    expect(fresh.length).toBeGreaterThan(0);
    expect(Math.min(...fresh.map((s) => s.box.left))).toBeGreaterThanOrEqual(33);
  });

  it("deletes the WHOLE chart, not just its tagged shape, when it is ungrouped", async () => {
    // PowerPoint on the web: no grouping, so the config tag can only sit on ONE
    // of the chart's shapes. The update deleted exactly that shape and redrew
    // the chart, leaving the other twelve underneath — 13 shapes became 25, then
    // 37, on successive edits, as stacked misaligned duplicates.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, (v) => v !== "1.8");
    const scene = buildChart(config);
    await insertSceneIntoSlide(scene, { tagData: "cfg" });
    expect(slide.created.some((s) => s.type === "group")).toBe(false);
    const drawn = slide.created.length;
    expect(drawn).toBeGreaterThan(1);
    // The tagged shape carries the rest of the chart with it.
    expect(JSON.parse(slide.created[0].tagStore.get(CHART_PARTS_TAG)!)).toHaveLength(drawn - 1);

    const live = () => slide.created.filter((s) => !s.deleted);
    for (const edit of [1, 2]) {
      // Same Scale: read the deck back, re-render every chart it finds.
      const found = (await listChartsInDeck()).filter((c) => live().some((s) => s.id === c.target.shapeId));
      expect(found, `edit ${edit}`).toHaveLength(1);
      await updateChartsInSlides([{ scene, target: found[0].target, opts: { tagData: "cfg" } }]);
      expect(live(), `edit ${edit}`).toHaveLength(drawn);
    }
  });

  it("does not resurrect a chart whose shape the user deleted", async () => {
    // The pane still holds an editTarget for a chart the user has since removed
    // from the slide. A stale SLIDE id is already treated as nothing to do; a
    // stale SHAPE id redrew the chart at its old position instead.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "cfg" });
    const group = slide.created.find((s) => s.type === "group")!;
    for (const s of slide.created) s.delete(); // the user deletes the chart
    await updateChartInSlide(buildChart(config), { slideId: "s1", shapeId: group.id, left: 10, top: 20 });
    expect(slide.created.filter((s) => !s.deleted)).toHaveLength(0);
  });
});

describe("selection readers", () => {
  it("loadChartFromSelection returns the tagged config and target", async () => {
    const slide = makeSlide("s1");
    const chart = makeShape("group", undefined, { left: 10, top: 20, width: 300, height: 200 });
    chart.tagStore.set(CHART_TAG, '{"kind":"pie"}');
    const other = makeShape("geometric", "rectangle", { left: 0, top: 0, width: 5, height: 5 });
    installHost([slide], [other, chart]);
    const res = await loadChartFromSelection();
    expect(res?.configJson).toBe('{"kind":"pie"}');
    expect(res?.target).toMatchObject({ slideId: "s1", shapeId: chart.id, left: 10, top: 20 });
  });

  it("loadChartFromSelection returns null for untagged selections", async () => {
    const slide = makeSlide("s1");
    installHost([slide], [makeShape("geometric", "rectangle", { left: 0, top: 0, width: 5, height: 5 })]);
    expect(await loadChartFromSelection()).toBeNull();
  });

  it("getSelectionBounds returns plain shape bounds but skips charts and multi-selects", async () => {
    const slide = makeSlide("s1");
    const box = makeShape("geometric", "rectangle", { left: 7, top: 8, width: 100, height: 60 });
    installHost([slide], [box]);
    expect(await getSelectionBounds()).toEqual({ left: 7, top: 8, width: 100, height: 60 });

    box.tagStore.set(CHART_TAG, "{}");
    installHost([slide], [box]);
    expect(await getSelectionBounds()).toBeNull();

    installHost([slide], [box, makeShape("geometric", "rectangle", { left: 0, top: 0, width: 1, height: 1 })]);
    expect(await getSelectionBounds()).toBeNull();
  });

  it("getSelectionBounds swallows host errors", async () => {
    vi.stubGlobal("PowerPoint", {
      run: async () => {
        throw new Error("no selection");
      },
    });
    expect(await getSelectionBounds()).toBeNull();
  });

  it("listChartsInSelection filters to tagged shapes", async () => {
    const slide = makeSlide("s1");
    const a = makeShape("group", undefined, { left: 1, top: 1, width: 1, height: 1 });
    a.tagStore.set(CHART_TAG, "{}");
    const b = makeShape("geometric", "rectangle", { left: 2, top: 2, width: 1, height: 1 });
    installHost([slide], [a, b]);
    const res = await listChartsInSelection();
    expect(res).toHaveLength(1);
    expect(res[0].target.shapeId).toBe(a.id);
  });
});

describe("listChartsInDeck", () => {
  it("finds tagged charts across all slides", async () => {
    const s1 = makeSlide("s1");
    const s2 = makeSlide("s2");
    installHost([s1, s2]);
    await insertSceneIntoSlide(buildChart(config), { tagData: '{"a":1}' });
    const g = s2.shapes.addGroup([]);
    g.tagStore.set(CHART_TAG, '{"b":2}');
    s2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });

    const found = await listChartsInDeck();
    expect(found).toHaveLength(2);
    expect(found.map((f) => f.target.slideId).sort()).toEqual(["s1", "s2"]);
    // The deck-wide sweep releases its proxy objects (one tag + one shape per
    // shape on every slide) once their values are read.
    expect(untracked.tags).toBeGreaterThan(0);
    expect(untracked.shapes).toBeGreaterThan(0);
  });
});

describe("proxy lifecycle", () => {
  it("listChartsInSelection untracks the shape and tag proxies it scans", async () => {
    const slide = makeSlide("s1");
    const a = makeShape("group", undefined, { left: 1, top: 1, width: 1, height: 1 });
    a.tagStore.set(CHART_TAG, "{}");
    const b = makeShape("geometric", "rectangle", { left: 2, top: 2, width: 1, height: 1 });
    installHost([slide], [a, b]);
    await listChartsInSelection();
    expect(untracked.tags).toBe(6); // both selected shapes' config, parts AND origin tags
    expect(untracked.shapes).toBe(2);
  });
});

describe("in-place update keeps the chart where it is", () => {
  // The tagged shape's left/top is NOT the frame origin: grouped it is the
  // group's bounding box, ungrouped it is whatever created[0] happens to be.
  // Feeding it back as the next render's origin shifted the chart by the scene's
  // ink offset — and compounded, so Same Scale walked charts off the slide.
  const cfg: ChartConfig = {
    kind: "stacked",
    ...DEFAULT_SIZE,
    data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
  };

  const cycle = async (grouped: boolean) => {
    const slide = makeSlide("s1");
    installHost([slide], [], slide, grouped ? () => true : (v) => v !== "1.8");
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });
    const seen: string[] = [];
    for (let i = 0; i < 3; i++) {
      const charts = await listChartsInDeck();
      expect(charts).toHaveLength(1);
      const t = charts[0].target;
      seen.push(`${Math.round(t.left)},${Math.round(t.top)}`);
      await updateChartsInSlides([{ scene: buildChart(cfg), target: t, opts: { tagData: JSON.stringify(cfg) } }]);
    }
    return seen;
  };

  it("does not drift when the host groups (desktop)", async () => {
    const seen = await cycle(true);
    expect(new Set(seen).size, `chart moved across update cycles: ${seen.join(" -> ")}`).toBe(1);
  });

  it("follows a chart the user has DRAGGED, instead of teleporting it back", async () => {
    // Re-rendering at the tagged shape's corner drifts; re-rendering at the
    // recorded origin teleports a moved chart back to where it was first
    // inserted, silently undoing the user's drag. The origin tag therefore also
    // records the ANCHOR (where the tagged shape landed), so an update shifts the
    // origin by exactly how far the shape has moved since.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });

    const before = (await listChartsInDeck())[0].target;
    // The user drags the whole chart across the slide.
    const [dx, dy] = [240, 110];
    for (const sh of slide.created) {
      sh.left += dx;
      sh.top += dy;
    }
    const moved = (await listChartsInDeck())[0].target;
    expect(moved.left).toBeCloseTo(before.left + dx, 5);

    await updateChartsInSlides([{ scene: buildChart(cfg), target: moved, opts: { tagData: JSON.stringify(cfg) } }]);
    const after = (await listChartsInDeck())[0].target;
    expect(after.left, "update dragged the chart back to its insert position").toBeCloseTo(moved.left, 5);
    expect(after.top, "update dragged the chart back to its insert position").toBeCloseTo(moved.top, 5);
  });

  it("follows a drag even when the caller reuses a CACHED target (what the pane does)", async () => {
    // The pane captures state.editTarget once, when the chart is loaded, and
    // re-uses it for every subsequent "Update chart". Measuring the drag against
    // that snapshot reports no movement, so the update put the chart back where
    // it was — the same teleport, just reached through the pane's real flow
    // rather than a fresh deck read.
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 4] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: JSON.stringify(cfg), left: 60, top: 90 });
    const held = (await listChartsInDeck())[0].target; // captured once, then kept

    for (const sh of slide.created.filter((s) => !s.deleted)) {
      sh.left += 200;
      sh.top += 60;
    }
    const moved = (await listChartsInDeck())[0].target;

    // The pane updates with the STALE target it has been holding all along.
    await updateChartInSlide(buildChart(cfg), held, { tagData: JSON.stringify(cfg) });
    const after = (await listChartsInDeck())[0].target;
    expect(after.left, "update teleported the chart back").toBeCloseTo(moved.left, 5);
    expect(after.top, "update teleported the chart back").toBeCloseTo(moved.top, 5);
  });

  it("does not drift when the host cannot group (web)", async () => {
    const seen = await cycle(false);
    expect(new Set(seen).size, `chart moved across update cycles: ${seen.join(" -> ")}`).toBe(1);
  });
});

describe("repeated in-place updates keep landing", () => {
  // An update replaces every shape, so the caller's EditTarget is dead the
  // moment it returns. Before updateChartsInSlides handed back a fresh one, the
  // SECOND update named a shape that no longer existed, was filtered out as
  // "the user deleted this chart", and did nothing — with no error. Auto-update
  // (which fires 900ms after every control change) died the same way after one
  // push, so the pane looked like "Update stopped working".
  it("a second update against the returned target still lands", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const cfg: ChartConfig = {
      kind: "stacked",
      ...DEFAULT_SIZE,
      data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 2] }] },
    };
    await insertSceneIntoSlide(buildChart(cfg), { tagData: '{"v":0}' });
    let target = (await listChartsInDeck())[0].target;

    for (const v of [1, 2, 3]) {
      const next = await updateChartInSlide(buildChart(cfg), target, { tagData: `{"v":${v}}` });
      expect(next, `update ${v} returned no target`).toBeTruthy();
      target = next!;
      // The edit actually reached the slide, every time.
      const live = await listChartsInDeck();
      expect(live, `update ${v} lost the chart`).toHaveLength(1);
      expect(JSON.parse(live[0].configJson).v, `update ${v} silently did nothing`).toBe(v);
    }
  });
});

describe("insertAgendaSlides", () => {
  it("appends one slide per chapter and renders ungrouped", async () => {
    const s1 = makeSlide("s1");
    const slides = [s1];
    installHost(slides);
    const chapters = ["Intro", "Findings", "Next steps"];
    const scenes = chapters.map((_, i) => buildAgendaScene(chapters, { highlight: i }));
    await insertAgendaSlides(scenes);
    expect(slides).toHaveLength(4);
    for (let i = 1; i < 4; i++) {
      expect(slides[i].created.length).toBeGreaterThan(0);
      expect(slides[i].created.some((s) => s.type === "group")).toBe(false);
    }
  });
});

describe("insertDemoDeck", () => {
  it("appends one slide per item and tags the charts with their config", async () => {
    const s1 = makeSlide("s1");
    const slides = [s1];
    installHost(slides);
    const items = [
      {
        scene: buildChart({
          ...DEFAULT_SIZE,
          kind: "pie" as const,
          data: { categories: ["A", "B"], series: [{ name: "S", values: [3, 1] }] },
        }),
        tagData: '{"kind":"pie"}',
      },
      {
        scene: buildChart({
          ...DEFAULT_SIZE,
          kind: "clustered" as const,
          data: { categories: ["A"], series: [{ name: "S", values: [5] }] },
        }),
        tagData: '{"kind":"clustered"}',
      },
      {
        scene: {
          width: 100,
          height: 40,
          nodes: [{ kind: "rect" as const, x: 0, y: 0, w: 10, h: 10, fill: "#111111" }],
        },
      }, // untagged element
    ];
    await insertDemoDeck(items);
    // Three slides appended after the original.
    expect(slides).toHaveLength(4);
    for (let i = 1; i < 4; i++) expect(slides[i].created.length).toBeGreaterThan(0);
    // The two chart slides carry their config tag; the element slide does not.
    expect(slides[1].created.some((s) => s.tagStore.get(CHART_TAG) === '{"kind":"pie"}')).toBe(true);
    expect(slides[3].created.every((s) => !s.tagStore.has(CHART_TAG))).toBe(true);
  });
});

describe("isPowerPointHost", () => {
  it("is false outside an Office host and true inside", () => {
    expect(isPowerPointHost()).toBe(false);
    vi.stubGlobal("PowerPoint", {});
    vi.stubGlobal("Office", { context: { host: "PowerPoint" } });
    expect(isPowerPointHost()).toBe(true);
  });
});

describe("marker symbols in the live add-in", () => {
  const markerScene = (markers: MarkerSymbol[]) =>
    buildChart({
      kind: "scatter",
      width: 480,
      height: 300,
      data: {
        categories: ["a", "b", "c"],
        series: [
          { name: "X", values: [1, 2, 3] },
          { name: "Y", values: [2, 4, 3] },
          { name: "Group", values: [1, 2, 3] },
        ],
      },
      scatter: { markers },
    });

  it("draws each symbol as native preset geometry, filled", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(markerScene(["diamond", "plus", "triangle"]), { left: 0, top: 0 });

    // Filled preset geometry is the whole reason a symbol is not a polygon:
    // PowerPoint can only outline a freeform, so a polygon marker would be
    // hollow here while the SVG preview showed it solid.
    for (const preset of ["diamond", "plus", "triangle"]) {
      const shapes = slide.created.filter((s) => s.geo === preset);
      expect(shapes.length, preset).toBeGreaterThan(0);
      for (const s of shapes) {
        expect(s.fillColor, preset).toMatch(/^#[0-9a-f]{6}$/i);
        expect(s.fillCleared, preset).toBe(false);
        expect(s.box.width, preset).toBeGreaterThan(0);
        expect(s.box.width, preset).toBeCloseTo(s.box.height, 9);
      }
    }
  });

  it("needs no rotation, so it works on a bare 1.4 host", async () => {
    // Arrowheads and pie fans need Shape.rotation (1.10+) and degrade without
    // it. The marker set is deliberately rotation-free: nothing here may set
    // rotation, so a 1.4 host draws the same shapes as a current one.
    const slide = makeSlide("s1");
    installHost([slide], [], slide, () => false);
    await insertSceneIntoSlide(markerScene(["diamond", "triangle", "plus"]), { left: 0, top: 0 });
    const presets = slide.created.filter((s) => ["diamond", "triangle", "plus"].includes(s.geo!));
    expect(presets.length).toBeGreaterThan(0);
    for (const s of presets) expect(s.rotation).toBeUndefined();
  });

  it("places the symbol's box centred on the point", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const scene = markerScene(["diamond", "diamond", "diamond"]);
    await insertSceneIntoSlide(scene, { left: 100, top: 50 });
    const nodes = scene.nodes.filter((n) => n.kind === "symbol");
    expect(nodes.length).toBeGreaterThan(0);
    for (const n of nodes) {
      if (n.kind !== "symbol") continue;
      const s = slide.created.find((c) => c.geo === "diamond" && Math.abs(c.box.left - (100 + n.cx - n.size)) < 1e-6);
      expect(s, `no shape at cx=${n.cx}`).toBeTruthy();
      expect(s!.box.top).toBeCloseTo(50 + n.cy - n.size, 9);
      expect(s!.box.width).toBeCloseTo(n.size * 2, 9);
    }
  });
});

describe("Office round-trips do not scale with the chart count", () => {
  const cfgFor = (v: number): ChartConfig => ({
    ...config,
    data: { categories: ["A", "B"], series: [{ name: "S1", values: [v, v + 1] }] },
  });
  const targetsOn = (slide: FakeSlide, n: number) =>
    Array.from({ length: n }, (_, i) => {
      // A real target names a shape that exists on the slide; make one per chart.
      const s = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
      return {
        scene: buildChart(cfgFor(i)),
        target: { slideId: slide.id, shapeId: s.id, left: 10, top: 20 },
        opts: { tagData: `{"i":${i}}` },
      };
    });

  it("re-renders N charts in ONE context, whatever N is", async () => {
    // The defect this guards: doSameScale awaited the single-chart update in a
    // loop, so each chart opened its own PowerPoint.run — 20 contexts across a
    // 20-chart deck. That is the property worth pinning. The SYNC count is no
    // longer flat and must not be: shapes commit in batches, because a live
    // canvas will not take a whole chart at once (SHAPES_PER_SYNC).
    for (const n of [1, 2, 10, 20]) {
      const slide = makeSlide("s1");
      installHost([slide]);
      await updateChartsInSlides(targetsOn(slide, n));
      expect(trips.contexts, `${n} charts`).toBe(1);
    }
  });

  it("costs syncs per BATCH OF SHAPES, not a fixed toll per chart", async () => {
    // The two failure modes this sits between: a per-chart context (the old
    // N+1, 80 round-trips for 20 charts), and a per-chart mega-batch that the
    // host silently refuses. Syncs must track the shapes, and nothing else.
    const slide = makeSlide("s1");
    installHost([slide]);
    await updateChartsInSlides(targetsOn(slide, 1));
    const one = trips.syncs;
    installHost([makeSlide("s2")]);
    const slide2 = makeSlide("s2");
    installHost([slide2]);
    await updateChartsInSlides(targetsOn(slide2, 2));
    const two = trips.syncs;
    // Doubling the charts doubles the drawing, not a fixed per-chart overhead:
    // the growth is the extra shapes' batches, so it stays well under 2x.
    expect(two).toBeGreaterThan(one);
    expect(two).toBeLessThan(one * 2 + 2);
  });

  it("still draws every chart it batches, tagged and grouped", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const items = targetsOn(slide, 3);
    await updateChartsInSlides(items);
    const groups = slide.created.filter((s) => s.type === "group");
    expect(groups).toHaveLength(3);
    // Each group carries its OWN config, not the last one written.
    expect(groups.map((g) => g.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":1}', '{"i":2}']);
    // The old shape each target named is gone.
    for (const it of items) expect(slide.created.find((s) => s.id === it.target.shapeId)!.deleted).toBe(true);
    // Charts land at their target's position, not the default offset.
    for (const r of slide.created.filter((s) => s.geo === "rectangle" && s.box.width > 1)) {
      expect(r.box.left).toBeGreaterThanOrEqual(10);
    }
  });

  it("keeps the single-chart paths to ONE context each", async () => {
    // updateChartInSlide is now updateChartsInSlides([one]); the Insert button
    // opens its own. Neither may open more than one, however many shapes the
    // chart has.
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(buildChart(config), { tagData: "{}" });
    expect(trips.contexts).toBe(1);

    const slide2 = makeSlide("s2");
    installHost([slide2]);
    const s = slide2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    await updateChartInSlide(buildChart(config), { slideId: "s2", shapeId: s.id, left: 0, top: 0 }, { tagData: "{}" });
    expect(trips.contexts).toBe(1);
  });

  it("does nothing, and opens no context, for an empty batch", async () => {
    installHost([makeSlide("s1")]);
    await updateChartsInSlides([]);
    expect([trips.syncs, trips.contexts]).toEqual([0, 0]);
  });

  it("keeps every chart re-editable when the grouping sync is refused", async () => {
    // Batching costs granularity: a refused grouping now loses grouping for the
    // whole batch, not just one chart. What must NOT be lost is the config tag —
    // the charts are already on the slide (their shapes committed a phase
    // earlier), so each must fall back to tagging its own first shape or it
    // silently stops being re-editable.
    //
    // The failure has to come from the SYNC, not from addGroup: Office.js only
    // reports queued commands there, which means every tag target has already
    // been pointed at a group that turned out not to exist. A test that throws
    // from addGroup instead never overwrites them and proves nothing.
    const slide = makeSlide("s1");
    installHost([slide]);
    // The group sync is no longer a fixed number: the shapes commit in batches
    // first, so its index depends on the chart's size. Find it rather than
    // hardcode it — a wrong number here silently tests nothing.
    const batches = Math.ceil(buildChart(cfgFor(0)).nodes.length / 10);
    // 1 resolve slides, 1 resolve old shapes, 1 delete, then each chart's
    // batches, then GROUP.
    failSyncOn = 3 + batches * 3 /* 3 charts */ + 1;
    try {
      const items = targetsOn(slide, 3);
      // One refreshed target per chart — the caller needs them to stay live.
      await expect(updateChartsInSlides(items)).resolves.toHaveLength(items.length);
      // Each chart's OWN config, back on each chart's OWN first shape.
      const tagged = slide.created.filter((s) => s.tagStore.has(CHART_TAG));
      expect(tagged.map((s) => s.tagStore.get(CHART_TAG))).toEqual(['{"i":0}', '{"i":1}', '{"i":2}']);
      expect(tagged.every((s) => s.type !== "group")).toBe(true);
    } finally {
      failSyncOn = 0;
    }
  });

  it("renders one slide per context and reports progress per slide", async () => {
    // One PowerPoint.run per slide isolates a chart the host can't finish and
    // keeps each context light (one chart's shapes, not a chunk's four). Progress
    // is per slide, so a slow host shows slides landing instead of freezing.
    for (const n of [2, 12, 35] as const) {
      installHost([makeSlide("s1")]);
      const seen: string[] = [];
      const report = await insertDemoDeck(
        Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
        (done, total) => seen.push(`${done}/${total}`),
      );
      expect(failedIndices(report), `${n} slides`).toEqual([]);
      // Self-check: the deck grew by exactly one slide per item, nothing lost.
      expect(report.slidesAdded, `${n} slides`).toBe(n);
      // One context per SLIDE, plus one more per slide for addSlides' settled
      // fresh-context verify (nothing lost, so no retry context), the two
      // settled slideCount reads (before/after), and the paged on-slide
      // readback (nothing lost, so it runs).
      expect(trips.contexts, `${n} slides`).toBe(2 * n + 2 + Math.ceil(n / READBACK_PAGE));
      expect(seen, `${n} slides`).toHaveLength(n);
      expect(seen.at(-1)).toBe(`${n}/${n}`);
      // Monotonic, never over-counting.
      expect(seen.map((x) => Number(x.split("/")[0]))).toEqual(
        [...seen.map((x) => Number(x.split("/")[0]))].sort((a, b) => a - b),
      );
    }
  });

  it("appends every demo slide, each tagged with its own config", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 35;
    const report = await insertDemoDeck(
      Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
    );
    expect(failedIndices(report)).toEqual([]);
    expect(report.slidesAdded).toBe(n);
    // The fake appends a slide per add(); the original + n new ones.
    expect(deck.length).toBe(1 + n);
    const tags = deck.slice(1).map((s) => s.created.map((c) => c.tagStore.get(CHART_TAG)).find(Boolean));
    expect(tags).toEqual(Array.from({ length: n }, (_, i) => `{"i":${i}}`));
  });

  it("records per-item and total wall-clock so a run's duration is on the record", async () => {
    installHost([makeSlide("s1")]);
    const n = 5;
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // Every item is timed; the total is present and never negative.
    expect(report.results.every((r) => typeof r.ms === "number" && r.ms >= 0)).toBe(true);
    expect(typeof report.totalMs).toBe("number");
    expect(report.totalMs).toBeGreaterThanOrEqual(0);
    // The whole run is at least as long as its slowest single item.
    expect(report.totalMs).toBeGreaterThanOrEqual(Math.max(...report.results.map((r) => r.ms)));
  });

  it("finds no blank slots and completes the readback on a clean run", async () => {
    installHost([makeSlide("s1")]);
    const n = 4;
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    expect(report.blankSlides).toEqual([]); // every added slide read back with shapes
    expect(report.blanksRead).toBe(true); // the readback finished
    expect(report.addsIssued).toBe(n); // one add per item, no retries/fails
  });

  it("reports a host-blanked slide by DECK POSITION, not by item name", async () => {
    const deck = [makeSlide("s1")]; // pre-existing at index 0; added slides take indices 1..n
    installHost(deck);
    const n = 4;
    blankReadbackAt.add(2); // the added slide at deck index 2 reads back empty on readback
    const report = await insertDemoDeck(Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // A blank slide has no content/tag to name it — reported as the 1-based deck position (index 2 → slide 3).
    expect(report.blankSlides).toEqual([3]);
    expect(report.blanksRead).toBe(true);
  });

  it("names a blank slide from its slot tag when the item carries a title", async () => {
    // A blank readback used to say only "slide 3". Every demo slide now gets a
    // POWERCHART_DEMO_SLOT tag on creation, so the readback can name the missing
    // chart by title. blankReadbackAt makes index 2 report 0 shapes; the slot
    // tag survives (we never emptied the tag store) and gives us the item name.
    _setBlankReReadDelayForTest(0); // no wall-clock sleep in the test
    try {
      const deck = [makeSlide("s1")];
      installHost(deck);
      const n = 3;
      blankReadbackAt.add(2);
      const report = await insertDemoDeck(
        Array.from({ length: n }, (_, i) => ({
          scene: buildChart(cfgFor(i)),
          title: `chart-${i}`,
        })),
      );
      expect(report.blankSlides).toEqual([3]);
      // Index 2 corresponds to item 1 (item 0 is index 1, item 1 is index 2, ...).
      expect(report.blankItems).toEqual([{ position: 3, title: "chart-1" }]);
    } finally {
      _setBlankReReadDelayForTest(200);
    }
  });

  it("un-masks a lost slide that a retry stray hid from the deck-growth count", async () => {
    // The exact real-host coincidence: the host loses one item's slide while a
    // retry leaves a stray, so net growth == items.length and the naive
    // (items.length − slidesAdded) reads 0. addsIssued − slidesAdded surfaces it.
    const deck = [makeSlide("s1")];
    installHost(deck);
    // Neither item carries tagData, so groupAndTagAll's tag/origin syncs never
    // fire — each item's steady-state cost is getCount, add, addSlides' settled
    // fresh-context verify, render, group (5 syncs; item 0 pays one more for
    // blankLayoutId). item0 strays: attempt-1 render (#6) fails after its add
    // lands, retry recovers → 2 slides. item1 is lost: both attempts fail at
    // their getCount (#14/#16), before any add → 0 slides.
    // (Sync numbers include catch-path readbacks: slideCount + slideShapeCount
    // per attempt-1 catch, plus one slideCount per addSlides-refused catch —
    // the softer PR-8 gate reads the slide back on every non-tooDense throw.)
    failSyncsOn.add(6);
    failSyncsOn.add(14);
    failSyncsOn.add(16);
    const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)) }, { scene: buildChart(cfgFor(1)) }]);
    expect(report.results[0].retried).toBe(true); // the stray item recovered
    expect(report.results[1].status).toBe("failed"); // the lost item
    expect(report.slidesAdded).toBe(2); // net growth == items.length → naive lost would read 0
    expect(report.addsIssued).toBe(4); // 2 items + 1 retried + 1 failed
    expect(report.addsIssued - report.slidesAdded).toBe(2); // the hardened lost surfaces the loss
  });

  it("marks the blank readback incomplete when it faults, not falsely clean", async () => {
    installHost([makeSlide("s1")]);
    faultShapeGetCount = true; // every readback getCount throws
    const report = await insertDemoDeck(Array.from({ length: 3 }, (_, i) => ({ scene: buildChart(cfgFor(i)) })));
    // An empty list must NOT read as "no blanks" when we could not measure.
    expect(report.blanksRead).toBe(false);
    expect(report.blankSlides).toEqual([]);
    // The run itself still succeeded — a readback fault is not a render failure.
    expect(report.results.every((r) => r.status === "rendered")).toBe(true);
  });

  it("treats a timed-out sync as rendered when the readback shows the shapes actually landed", async () => {
    // The real-host bug PR 1 exists for: withTimeout rejected at 45s, but the
    // sync had queued every shape and the host settled a moment later. Marking
    // it "failed" both wastes a retry (adds a duplicate slide) and stamps a
    // real chart NOT COMPLETE. The fix waits for the abandoned promise to
    // report its outcome, reads the slide back, and — when every shape landed
    // — records it as rendered with lateSettled=true.
    _setBatchTimeoutForTest(5);
    stallSyncDelayMs = 40; // safely past the 5ms timeout
    try {
      const deck: FakeSlide[] = [makeSlide("s1")];
      installHost(deck);
      // Sync map for the first item: #1=slideCount, #2=blankLayoutId,
      // #3=addSlides.getCount, #4=addSlides.add, #5=addSlides' settled
      // fresh-context verify (nothing lost, no retry), #6=renderShapesChunked's
      // only batch (cfgFor's scene fits in one). Stall #6 — the render sync —
      // so the shapes are already on the fake slide when withTimeout gives up.
      stallSyncOn.add(6);
      const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)), tagData: `{"i":0}` }]);
      // Late-settled path: rendered without retry, with lateOutcome captured.
      expect(report.results[0].status).toBe("rendered");
      expect(report.results[0].lateSettled).toBe(true);
      expect(report.results[0].retried).toBeFalsy();
      expect(report.results[0].lateOutcome).toMatch(/eventually SUCCEEDED/);
      expect(failedIndices(report)).toEqual([]);
      // No duplicate slide from a bogus retry.
      expect(report.slidesAdded).toBe(1);
      expect(report.addsIssued).toBe(1);
      // And the slide is NOT stamped — stamping a complete chart NOT COMPLETE
      // is the false-negative the fix eliminates.
      expect(deck[1].created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncDelayMs = 40;
    }
  });

  it("rescues an ungrouped rendered slide by regrouping in a fresh context", async () => {
    // Presentation_3.pptx: PR 3's in-context regroup fires only if the render
    // itself finishes. A lateSettled render — sync timed out but shapes landed
    // — dies with the context and never groups; the chart ends up loose,
    // untagged, not re-editable. The rescue reopens a fresh context, loads
    // slide.shapes.items, and addGroups them so the chart is one shape again.
    _setBatchTimeoutForTest(5);
    stallSyncDelayMs = 40;
    try {
      const deck: FakeSlide[] = [makeSlide("s1")];
      installHost(deck);
      // Stall the render sync (see the lateSettled test above) so the
      // addAndRenderItem context dies before groupAndTagAll runs.
      stallSyncOn.add(6);
      const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)), tagData: `{"i":0}` }]);
      expect(report.results[0].status).toBe("rendered");
      expect(report.results[0].lateSettled).toBe(true);
      // The rescue ran and the chart is now one group carrying its config tag.
      expect(report.results[0].grouped).toBe(true);
      const groups = deck[1].created.filter((s) => s.type === "group");
      expect(groups).toHaveLength(1);
      expect(groups[0].tagStore.get(CHART_TAG)).toBe('{"i":0}');
    } finally {
      _setBatchTimeoutForTest(45_000);
      stallSyncDelayMs = 40;
    }
  });

  it("treats a mid-render failure with ≥85% shapes on the slide as rendered-partial, no retry, no stamp", async () => {
    // Presentation_3.pptx: Line landed 31/36 shapes (86%) — PR 1's strict
    // `readback >= shapeCount` gate missed it, retried into a duplicate
    // slide, stamped both NOT COMPLETE. The softer PR-8 gate treats a
    // readback ≥ ceil(shapeCount * 0.85) as rendered-partial: no retry, no
    // stamp, rescue groups whatever landed.
    //
    // Multi-batch scene at the off-screen batch size (40): 47 rect nodes
    // means batch 1 = 40 shapes (sync OK), batch 2 = 7 shapes (sync FAIL).
    // With the fake's discard-on-throw fidelity, readback = 40 = ceil(47 * 0.85).
    const NODES = 47;
    const scene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    // Sync map: #1 outer slideCount, #2 blankLayoutId, #3 addSlides getCount,
    // #4 addSlides add, #5 addSlides' settled fresh-context verify (nothing
    // lost, no retry), #6 render batch 1 (40 shapes), #7 render batch 2 (FAIL).
    failSyncOn = 7;
    try {
      const report = await insertDemoDeck([{ scene, tagData: '{"i":0}', title: "big" }]);
      expect(report.results[0].status).toBe("rendered");
      expect(report.results[0].partialLanded).toBe(true);
      expect(report.results[0].retried).toBeFalsy();
      expect(report.results[0].lateSettled).toBeFalsy();
      // No duplicate slide from a bogus retry.
      expect(report.slidesAdded).toBe(1);
      expect(report.addsIssued).toBe(1);
      // Slide has the 40 shapes that committed — no stamp banner.
      const stamps = deck[1].created.filter((s) => s.name === "PowerChart:not-complete");
      expect(stamps).toHaveLength(0);
      // Rescue grouped whatever landed → chart is re-editable.
      expect(report.results[0].grouped).toBe(true);
      const groups = deck[1].created.filter((s) => s.type === "group");
      expect(groups).toHaveLength(1);
      expect(groups[0].tagStore.get(CHART_TAG)).toBe('{"i":0}');
    } finally {
      failSyncOn = 0;
    }
  });

  it("unstamps and rescues a genuinely-failed item whose last attempt still landed real shapes", async () => {
    // Both attempt 1 AND its retry fail below PARTIAL_RENDER_THRESHOLD (85%),
    // so the item would normally end status:"failed" with a NOT-COMPLETE
    // stamp over whatever landed. But the retry's failure still left real
    // chart shapes on the slide (just not enough to qualify as
    // rendered-partial) — the unstamp-and-rescue path should delete that
    // banner and group those shapes instead of leaving them hidden under it.
    //
    // 60 rect nodes, multi-batch at the off-screen batch size (40): batch 1 =
    // 40 shapes, batch 2 = 20 shapes.
    //
    // Sync map (addSlides opens ITS OWN verify sync — see the lost-adds
    // self-heal): #1 outer slideCount, #2 blankLayoutId, #3 addSlides
    // getCount, #4 addSlides add, #5 addSlides verify, #6 attempt-1 render
    // batch 1 (FAIL — 0 shapes commit, batch discarded). Attempt-1 catch's
    // readback: #7 slideCount, #8 slideShapeCount (reads 0) — below the
    // partial threshold, falls through to the retry. Retry: #9 addSlides
    // getCount, #10 addSlides add, #11 addSlides verify, #12 render batch 1
    // (40 shapes, OK), #13 render batch 2 (FAIL — 20 shapes discarded).
    // Retry catch's readback: #14 slideCount, #15 slideShapeCount (reads 40
    // = 66% of 60, under the 85% gate) — still not enough to call
    // rendered-partial, so status ends "failed" and stampLastSlide stamps.
    const NODES = 60;
    const scene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    failSyncsOn.add(6);
    failSyncsOn.add(13);
    try {
      const report = await insertDemoDeck([{ scene, tagData: '{"i":0}', title: "big" }]);
      expect(report.results[0].status).toBe("rendered");
      expect(report.results[0].partialLanded).toBe(true);
      expect(report.results[0].grouped).toBe(true);
      expect(report.results[0].created).toBe(40);
      // The retry's stray slide is where the 40 shapes and the stamp landed —
      // the LAST slide in the deck (index 2: preexisting + attempt-1's empty
      // stray + the retry's slide).
      // The fake models a real Office.js delete: the shape is flagged deleted,
      // not spliced out of `created` (see `getCount`/`items` on FakeSlide,
      // which both filter `!deleted` — the same lens a real readback uses).
      const rescuedSlide = deck[deck.length - 1];
      const stamps = rescuedSlide.created.filter((s) => s.name === "PowerChart:not-complete" && !s.deleted);
      expect(stamps).toHaveLength(0);
      const groups = rescuedSlide.created.filter((s) => s.type === "group" && !s.deleted);
      expect(groups).toHaveLength(1);
      expect(groups[0].tagStore.get(CHART_TAG)).toBe('{"i":0}');
    } finally {
      failSyncsOn.clear();
    }
  });

  it("bypassBudget lets a text-heavy scene render even when its shape count is over the budget", async () => {
    // The results/contents slide bug: 32 failures pushed the results scene to
    // 135 shapes — over DEMO_SHAPE_BUDGET (90) — and the run's own summary
    // came back as a red "NOT COMPLETE" stamp. Text-only scenes don't hit the
    // wedge/polygon flood the budget guards against; they should render.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const denseTextScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "text" as const,
        x: k,
        y: 0,
        w: 40,
        h: 20,
        text: `row ${k}`,
        fontSize: 12,
        color: "#000000",
        align: "left" as const,
        valign: "top" as const,
      })),
    };
    // With bypassBudget: the scene renders as a real chart, no stamp.
    const withBypass = await insertDemoDeck([{ scene: denseTextScene, title: "Results", bypassBudget: true }]);
    expect(withBypass.results[0].status).toBe("rendered");
    expect(deck[1].created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    expect(deck[1].created.filter((s) => s.type === "text").length).toBeGreaterThanOrEqual(120);
    // Without bypassBudget: the scene is stamped instead of drawn.
    installHost([makeSlide("s1")]);
    const withoutBypass = await insertDemoDeck([{ scene: denseTextScene, title: "Results" }]);
    expect(withoutBypass.results[0].status).toBe("skipped");
  });

  it("retries a slide the host stalls on once, and the transient stall recovers", async () => {
    // A single refused sync (a transient host stall) must not lose the slide: the
    // item is retried once in a fresh context, its later syncs land, and the deck
    // completes whole. This is the everyday web flake the retry exists for.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 6;
    // Fail a single sync partway in — item 1's render sync. Steady-state per
    // item (layout resolved once, on item 0) is 7 syncs: addSlides getCount,
    // add, its settled fresh-context verify, render, group, tag, origin-tag.
    // Item 0 additionally pays for blankLayoutId (+1, syncs 2-9); item 1 then
    // runs syncs 10 (getCount) .. 16 (origin), so its render sync is #13.
    failSyncOn = 13;
    try {
      const report = await insertDemoDeck(
        Array.from({ length: n }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
      );
      // Nothing ends failed — the one stall was recovered on retry.
      expect(failedIndices(report)).toEqual([]);
      // Exactly the stalled item is flagged retried.
      expect(report.results.filter((r) => r.retried).length).toBe(1);
    } finally {
      failSyncOn = 0;
    }
  });

  it("gives up after a single retry when the stall persists, and finishes the rest", async () => {
    // A chart the host genuinely cannot draw stalls BOTH the first attempt and the
    // retry. It must then be given up (status failed) without aborting the deck —
    // the remaining slides still render. failSyncsOn hits each attempt's first
    // propagating sync (the addSlides getCount): #3 = item 0 attempt 1, #5 = its
    // retry. (Sync #4 is the catch-path slideCount readback the PR-8 gate adds.)
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    failSyncsOn.add(3);
    failSyncsOn.add(5);
    const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)) }, { scene: buildChart(cfgFor(1)) }]);
    expect(report.results[0].status).toBe("failed");
    expect(report.results[0].retried).toBeFalsy(); // a retry that also failed is not a recovery
    expect(report.results[1].status).toBe("rendered"); // the deck kept going
  });

  it("skips a chart too dense for the host, keeps the slide, and stamps it NOT COMPLETE", async () => {
    // The heavy charts (area ~208 shapes) will not land on web and burn the
    // timeout trying. They are skipped up-front — the slide is kept (coverage),
    // its chart NOT drawn, and a stamp makes the placeholder unmistakable.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 1,
        h: 1,
        fill: "#111111",
      })),
    };
    const light = () => buildChart(cfgFor(0)); // a handful of shapes, well under budget
    const report = await insertDemoDeck([{ scene: light() }, { scene: dense }, { scene: light() }]);
    // Only the dense one is reported (as "skipped"); the deck still has all three.
    expect(failedIndices(report)).toEqual([1]);
    expect(report.results[1].status).toBe("skipped");
    expect(report.slidesAdded).toBe(3);
    expect(deck.length).toBe(1 + 3);
    // The dense slide (deck[2]) carries a stamp, NOT the 120 chart shapes.
    const denseSlide = deck[2];
    const stamp = denseSlide.created.find((s) => s.name === "PowerChart:not-complete");
    expect(stamp, "dense slide is stamped").toBeTruthy();
    expect(stamp!.text).toContain("NOT COMPLETE");
    // A top strip, not a slab over the middle — a mis-targeted stamp must not
    // obliterate a real chart under it (a 540pt-tall slide).
    expect(stamp!.top, "stamp sits at the top").toBeLessThan(80);
    expect(stamp!.height, "stamp is a strip").toBeLessThan(120);
    expect(denseSlide.created.length, "chart not drawn").toBeLessThan(120);
    // The light neighbours rendered as real charts (no stamp).
    expect(deck[1].created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    expect(deck[3].created.length).toBeGreaterThan(1);
  });

  it("keeps the rescue aimed at the right slide after a failure whose slide still landed", async () => {
    // The drift: `runningCount` is only re-synced on the failure paths that ALSO
    // read the failed slide's shapes (`!tooDense && shapeCount > 0`). A too-dense
    // item whose stamp sync rejects lands its slide but takes neither branch — so
    // the count stays one behind FOR THE REST OF THE RUN, and every later rescue
    // (which indexes `runningCount - 1`) reaches the PREVIOUS slide. Best case it
    // finds the 1-shape NOT-COMPLETE placeholder and bails, leaving a real chart
    // ungrouped; worst case it groups a neighbour's chart and writes THIS chart's
    // POWERCHART_CONFIG onto it, so editing one chart silently opens another.
    //
    // Sync map (deck of 1, item 0 pays for blankLayoutId): 1 = the run's opening
    // slideCount, 2 = blankLayoutId, 3-5 = addSlides (getCount, add, settled
    // verify), 6 = the too-dense item's stampSlide. Failing 6 leaves its slide on
    // the deck with only the fresh-context NOT-COMPLETE banner.
    //
    // Item 1 is then made to land ungrouped by refusing ONE addGroup — a knob,
    // not a sync index, because the fix itself adds a sync on the failure path
    // and any index pinned past it would move.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 1,
        h: 1,
        fill: "#111111",
      })),
    };
    failSyncsOn.add(6);
    refuseGroups = 1; // item 1's own grouping only — the rescue's addGroup still works
    const report = await insertDemoDeck([
      { scene: dense, title: "too dense" },
      { scene: buildChart(cfgFor(1)), tagData: '{"i":1}', title: "ungrouped" },
    ]);
    // Premise: the dense item's slide is on the deck, stamped, chart-less.
    expect(deck.length).toBe(3);
    expect(report.results[0].status).toBe("failed");
    expect(deck[1].created.map((s) => s.name)).toEqual(["PowerChart:not-complete"]);
    // The rescue found item 1's OWN slide: its chart is one group carrying its
    // own config tag, and nothing was grouped onto the placeholder next door.
    expect(report.results[1].status).toBe("rendered");
    expect(report.results[1].grouped, "item 1 was rescued into a group").toBe(true);
    const groups = deck[2].created.filter((s) => s.type === "group");
    expect(groups).toHaveLength(1);
    expect(groups[0].tagStore.get(CHART_TAG)).toBe('{"i":1}');
    expect(
      deck[1].created.some((s) => s.type === "group"),
      "placeholder untouched",
    ).toBe(false);
  });

  it("does not report a phantom lost slide for a too-dense item whose stamp was refused", async () => {
    // `addsIssued` used to be inferred as "one per item, plus one more for every
    // retried or failed item". A too-dense item never re-renders — there is
    // nothing to re-render, only a placeholder to stamp — so when its stamp sync
    // is refused it ends "failed" having issued exactly ONE add. The inference
    // then charged it a second, and `addsIssued − slidesAdded` accused the host
    // of losing a slide it had actually kept. This harness exists to stop
    // inventing failures, so a false ⚠ is the bug.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const dense = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: 120 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 1,
        h: 1,
        fill: "#111111",
      })),
    };
    failSyncsOn.add(6); // the too-dense item's stampSlide — see the test above
    const report = await insertDemoDeck([
      { scene: dense, title: "too dense" },
      { scene: buildChart(cfgFor(1)), tagData: '{"i":1}', title: "fine" },
    ]);
    expect(report.results[0].status).toBe("failed"); // the premise: stamped, refused
    expect(report.results[0].retried, "a too-dense item never re-renders").toBeFalsy();
    expect(report.results[0].attempts, "one add issued, not two").toBe(1);
    // Both slides are on the deck, so nothing was lost — and the report agrees.
    expect(report.slidesAdded).toBe(2);
    expect(report.addsIssued).toBe(2);
    expect(report.addsIssued - report.slidesAdded, "no phantom loss").toBe(0);
  });

  it("self-check catches a slide the host silently drops (deck grew by less than asked)", async () => {
    // The corruption a visual scan misses and today cost us 3 lost slides: an
    // add() that never lands leaves the deck one slide short with no error the
    // user sees. The report's slidesAdded is read back from the host, so the
    // shortfall is caught — and the dropped item shows up as not rendered.
    //
    // addSlides now self-heals a SINGLE dropped add with its own fresh-context
    // retry (see the addSlides retry/verify tests below), so a genuinely lost
    // slide here needs BOTH addSlides attempts (the item's own outer retry
    // makes a second full attempt) to exhaust their own add-then-retry pairs:
    // attempt 1's add + its in-addSlides retry (2), attempt 2's add + its
    // in-addSlides retry (2 more) — 4 drops before the deck is truly one short.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    swallowAdds = 4; // both outer attempts' addSlides calls fully exhausted → gone for good
    try {
      const report = await insertDemoDeck(
        Array.from({ length: 4 }, (_, i) => ({ scene: buildChart(cfgFor(i)), tagData: `{"i":${i}}` })),
      );
      // 4 items asked for, but the deck only grew by 3 — the lost-slide signal.
      expect(report.slidesAdded).toBe(3);
      expect(report.results).toHaveLength(4); // every item is still accounted for
      expect(failedIndices(report).length).toBeGreaterThanOrEqual(1); // the dropped one is flagged
    } finally {
      swallowAdds = 0;
    }
  });

  it("addSlides self-heals one dropped add via its own fresh-context retry, and surfaces it when the retry also fails", async () => {
    // addSlides now verifies its own adds landed (a settled getCount() in a
    // FRESH context, after the existing 2 syncs) and gets ONE retry round
    // before giving up — the fix for the Presentation_3.pptx bug where
    // PowerPoint web silently dropped ~half of 20 issued add()s. A single
    // dropped add should never even reach insertDemoDeck's own item-level
    // retry: it should be invisible, recovered inside addSlides itself.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    swallowAdds = 1; // the FIRST add() call anywhere is dropped, none after
    try {
      const report = await insertDemoDeck([{ scene: buildChart(cfgFor(0)), tagData: '{"i":0}' }]);
      // The retry landed: the deck grew by exactly the one slide asked for.
      expect(report.slidesAdded).toBe(1);
      // Recovered INSIDE addSlides — insertDemoDeck's own item-level retry
      // never had to fire.
      expect(report.results[0].retried).toBeFalsy();
      expect(report.results[0].status).toBe("rendered");
      // Nothing was lost AT COMMIT: the one drop was fully recovered.
      expect(report.addsLostAtCommit).toBe(0);
    } finally {
      swallowAdds = 0;
    }

    // Second sub-case: the drop persists through addSlides' one retry round
    // too. A single insertDemoDeck attempt only issues 2 adds per addSlides
    // call (the original + the in-addSlides retry), and a failed attempt
    // itself triggers insertDemoDeck's OWN item-level retry — a second full
    // addAndRenderItem attempt, i.e. a second addSlides call with its own
    // add + retry pair. So defeating item 0 entirely takes 4 dropped adds:
    // attempt 1's add + retry (2), attempt 2's add + retry (2 more). A second
    // item follows so the run is not a TOTAL loss (which insertDemoDeck itself
    // would throw on, per "A whole deck lost to HOST errors" below) — item 1
    // renders normally once swallowAdds is exhausted.
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const deck2: FakeSlide[] = [makeSlide("s1")];
    installHost(deck2);
    swallowAdds = 4;
    try {
      const report = await insertDemoDeck([
        { scene: buildChart(cfgFor(0)), tagData: '{"i":0}' },
        { scene: buildChart(cfgFor(1)), tagData: '{"i":1}' },
      ]);
      // Item 0 is a genuine total loss; item 1 landed once swallowAdds ran out.
      expect(report.slidesAdded).toBe(1);
      expect(report.results[0].status).toBe("failed");
      expect(report.results[1].status).toBe("rendered");
      // addSlides confirmed the loss (not recovered by its own retry) on
      // BOTH item 0's outer attempt and its own item-level retry.
      expect(report.addsLostAtCommit).toBeGreaterThanOrEqual(1);
      expect(warnSpy).toHaveBeenCalled();
    } finally {
      swallowAdds = 0;
      warnSpy.mockRestore();
    }
  });

  it('render:"image" draws ONE picture shape instead of the scene nodes', async () => {
    // The whole point of image mode: a dense chart becomes one shape, so the
    // PowerPoint-web dense-shape wall (office-js #4272 / #5022 / #6498) never
    // gets hit. A violin-sized scene is ~250 native shapes; here it must be 1.
    const NODES = 25;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA", left: 60, top: 90 });
    // Exactly one shape, geometry rectangle, sized to the FRAME (not the nodes).
    const live = slide.created.filter((s) => !s.deleted);
    expect(live).toHaveLength(1);
    expect(live[0].geo).toBe("rectangle");
    expect(live[0].box).toEqual({ left: 60, top: 90, width: 480, height: 300 });
    // The fill became a picture, carrying the payload, and the outline is off.
    expect(live[0].imageBase64).toBe("AAAA");
    expect(live[0].fillType).toBe("PictureAndTexture");
    expect(live[0].lineFormat.visible).toBe(false);
    // Named like a chart group so the Selection Pane reads the same either way,
    // and it carries the config tag — the picture IS the re-editable chart.
    expect(live[0].name).toBe("PowerChart");
    expect(live[0].tagStore.get(CHART_TAG)).toBe("{}");
  });

  it("strips every base64 spelling down to the bare payload the host wants", async () => {
    // Three forms circulate: a browser toDataURL (`data:image/png;base64,…`),
    // what render-pptx.mjs hands pptxgen (`image/png;base64,…` — NO data:
    // scheme), and already-bare. A `startsWith("data:")` guard would pass the
    // middle one through with `image/png;base64,` still glued on, and the host
    // would get a corrupt payload. Splitting on the last comma handles all three.
    const scene = { width: 100, height: 100, nodes: [{ kind: "rect" as const, x: 0, y: 0, w: 4, h: 4, fill: "#111" }] };
    for (const [input, expected] of [
      ["data:image/png;base64,PAYLOAD", "PAYLOAD"],
      ["image/png;base64,PAYLOAD", "PAYLOAD"],
      ["PAYLOAD", "PAYLOAD"],
    ] as const) {
      const slide = makeSlide("s1");
      installHost([slide]);
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: input });
      expect(slide.created.filter((s) => !s.deleted)[0].imageBase64, input).toBe(expected);
    }
  });

  it("falls back to native shapes on a host without PowerPointApi 1.8", async () => {
    // setImage is 1.8 and the manifests admit hosts from 1.4, so this must be
    // GATED, not attempted-and-caught: a queued command the host rejects takes
    // the whole sync with it. The chart must still land, as shapes.
    const NODES = 12;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    // Everything except 1.8 — so grouping is off too, exactly like the web host.
    installHost([slide], [], slide, (v) => v !== "1.8");
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
    const live = slide.created.filter((s) => !s.deleted);
    expect(live.length).toBeGreaterThanOrEqual(NODES); // the nodes, not a picture
    expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
    // Still re-editable: the config tag landed on the first shape (no group at 1.4).
    expect(live.some((s) => s.tagStore.get(CHART_TAG) === "{}")).toBe(true);
  });

  it("falls back to native shapes when the host refuses the picture fill", async () => {
    // A 1.8-advertising host that still rejects setImage (wrong payload format,
    // host quirk). renderPictureShape catches, best-effort deletes the rect, and
    // renderShapesChunked draws the nodes in the SAME request context.
    const NODES = 12;
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    refusePictureFill = true;
    try {
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
      const live = slide.created.filter((s) => !s.deleted);
      // The nodes landed, and no picture-filled shape survived.
      expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
      expect(live.filter((s) => s.geo === "rectangle").length).toBeGreaterThanOrEqual(NODES);
      expect(live.some((s) => s.tagStore.get(CHART_TAG) === "{}")).toBe(true);
    } finally {
      refusePictureFill = false;
    }
  });

  it("refuses an over-budget payload rather than burning the batch timeout", async () => {
    // MAX_PICTURE_BASE64 is a guard against a pathological custom frame size,
    // not a measured host limit — real payloads are 20-133 KB. Crossing it
    // degrades to shapes; the code that says so is PC-IMG-TOOBIG on the console,
    // because the chart still appears and the fallback is otherwise invisible.
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: 12 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const slide = makeSlide("s1");
    installHost([slide]);
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "A".repeat(4_000_001) });
      const live = slide.created.filter((s) => !s.deleted);
      expect(live.some((s) => s.imageBase64 !== undefined)).toBe(false);
      expect(live.length).toBeGreaterThanOrEqual(12);
      // A specific, greppable code — so a real bug report teaches us the number.
      expect(warn.mock.calls.flat().join(" ")).toContain("PC-IMG-TOOBIG");
    } finally {
      warn.mockRestore();
    }
  });

  it("costs ONE sync for the drawing, however dense the scene", async () => {
    // The reliability claim: image mode replaces N/batchSize render syncs with
    // exactly one, so a chart that could never commit as shapes lands in a
    // single round trip. Compare a 60-node scene both ways.
    const scene = {
      width: 480,
      height: 300,
      nodes: Array.from({ length: 60 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    installHost([makeSlide("shapes")]);
    await insertSceneIntoSlide(scene, { tagData: "{}" });
    const shapesSyncs = trips.syncs;

    installHost([makeSlide("picture")]);
    await insertSceneIntoSlide(scene, { tagData: "{}", pictureBase64: "AAAA" });
    const pictureSyncs = trips.syncs;

    // 60 nodes at SHAPES_PER_SYNC=10 is 6 render syncs; the picture is 1. Both
    // pay the same tag/group syncs afterwards, so the picture must be strictly
    // and substantially cheaper.
    expect(pictureSyncs).toBeLessThan(shapesSyncs);
    expect(shapesSyncs - pictureSyncs).toBeGreaterThanOrEqual(4);
  });

  it("re-fetches the slide's shape collection before addGroup on a multi-batch chart", async () => {
    // The real-host bug this guards: a >10-shape chart commits in multiple
    // batches, and the Shape proxies returned by earlier batches have their
    // object paths rewritten to getItem(id) by the time the group sync runs.
    // The web host silently drops addGroup(theseStaleProxies), leaving the
    // chart loose and unable to carry its POWERCHART_CONFIG tag. In the run
    // behind this fix, agenda / KPI+flow / table all landed ungrouped and
    // therefore un-re-editable. Fix: re-load slide.shapes.items right before
    // addGroup and pass those fresh proxies to addGroup.
    const NODES = 25; // 3 batches at SHAPES_PER_SYNC=10
    const bigScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    strictGroup = true; // web-host stale-proxy semantics — without the re-fetch, no group appears
    try {
      const report = await insertDemoDeck([{ scene: bigScene, tagData: '{"i":0}', title: "multi-batch" }]);
      expect(report.results[0].status).toBe("rendered");
      // Every appended chart is grouped — the new field flowed through.
      expect(report.results[0].grouped).toBe(true);
      // The slide holds ONE native group carrying the CHART_TAG — proving the
      // group survived and the tag landed on it (not on a stray first shape).
      const groups = deck[1].created.filter((s) => s.type === "group");
      expect(groups).toHaveLength(1);
      expect(groups[0].tagStore.get(CHART_TAG)).toBe('{"i":0}');
    } finally {
      strictGroup = false;
    }
  });

  it("re-acquires each freshly-added slide per batch, so a rewritten getItemAt cannot 5010 mid-deck", async () => {
    // The real regression: HOLD one getItemAt handle to a new slide and reuse it
    // across the render's batched syncs, and once Office.js rewrites its path to
    // getItem(<web-non-round-trippable id>) the next shape throws "InvalidParam
    // passed to GetItem(id)", code 5010 — the deck dies partway through, as it did
    // on the real host. The fix re-acquires a fresh proxy each batch; the fake
    // window-limits a held one.
    //
    // Load-bearing: each slide must span MORE than one batch, because a held
    // handle only goes stale on the batch AFTER a sync. SHAPES_PER_SYNC is 10, so
    // a 25-node scene is 3 batches — a single-batch chart (e.g. cfgFor) can hold
    // its handle and never notice, which is exactly how a weaker version of this
    // test passed against the very bug it meant to guard.
    const NODES = 25;
    const bigScene = {
      width: 100,
      height: 100,
      nodes: Array.from({ length: NODES }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    };
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const n = 6;
    const report = await insertDemoDeck(Array.from({ length: n }, () => ({ scene: bigScene })));
    expect(failedIndices(report)).toEqual([]);
    // Every appended slide got all its shapes (plus the group) — nothing stranded
    // by a mid-batch 5010.
    expect(deck.length).toBe(1 + n);
    for (let i = 1; i <= n; i++) expect(deck[i].created.length, `slide ${i}`).toBeGreaterThanOrEqual(NODES);
  });
});

describe("a stalled host is legible, and does not hang the pane", () => {
  it("reports every phase, in order, with the shape count", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const seen: string[] = [];
    await insertSceneIntoSlide(buildChart(config), { tagData: "{}" }, (p, d) => seen.push(d ? `${p}:${d}` : p));
    expect(seen[0]).toBe("context");
    expect(seen.at(-1)).toBe("done");
    // "commit" now repeats — once per batch — because shapes land in batches.
    expect(seen.filter((s) => s.startsWith("commit:")).length).toBeGreaterThan(1);
    expect([...new Set(seen.map((s) => s.split(":")[0]))]).toEqual(["context", "queue", "commit", "group", "done"]);
    expect(seen.find((s) => s.startsWith("queue:"))).toMatch(/^queue:\d+ nodes$/);
    // Real progress: "10 of 40 shapes", ending at the total.
    const commits = seen.filter((s) => s.startsWith("commit:"));
    expect(commits[0]).toMatch(/^commit:\d+ of \d+ shapes$/);
    const [done, total] = commits
      .at(-1)!
      .match(/(\d+) of (\d+)/)!
      .slice(1);
    expect(done).toBe(total);
  });

  it("gives up on a host that never answers, naming the phase it died in", async () => {
    // The real failure mode: Office.js does not throw when the host stops
    // answering — the sync promise simply never settles, so the pane spins for
    // ever with nothing to report. This is the only way out.
    vi.useFakeTimers();
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      // A sync that never settles, exactly like a stalled PowerPoint.ashx.
      (slide as unknown as { id: string }).id = "s1";
      const ctxSync = () => new Promise<void>(() => {});
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: ctxSync,
          }),
      });
      const seen: string[] = [];
      const p = insertSceneIntoSlide(buildChart(config), {}, (ph) => seen.push(ph));
      const assertion = expect(p).rejects.toThrow(/did not respond while drawing shapes \d+-\d+ of \d+/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      // And it says where it stopped — "commit" is the last thing reached.
      expect(seen.at(-1)?.startsWith("commit")).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("a host that answers late still gets heard", () => {
  it("reports the real Office error when an abandoned sync finally rejects", async () => {
    // The evidence problem: racing a timeout throws the answer away. The
    // abandoned sync keeps running, and Office.js reports queued-command
    // failures THERE and nowhere else — so whatever it says next is the only
    // description of the bug we will ever get. Without this it is lost.
    vi.useFakeTimers();
    const heard: string[] = [];
    onLateSync((m) => heard.push(m));
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      let rejectSync!: (e: unknown) => void;
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: () => new Promise<void>((_, rej) => (rejectSync = rej)),
          }),
      });
      const p = insertSceneIntoSlide(buildChart(config), {});
      const assertion = expect(p).rejects.toThrow(/did not respond/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      expect(heard, "nothing heard before the host answers").toHaveLength(0);

      // Now the host finally answers — with a real RichApi-shaped error.
      rejectSync({
        message: "An internal error has occurred.",
        code: "GeneralException",
        debugInfo: { errorLocation: "Shape.name" },
      });
      await vi.advanceTimersByTimeAsync(1);
      expect(heard).toHaveLength(1);
      // The generic message alone is useless; code + debugInfo name the bug.
      expect(heard[0]).toContain("the host eventually FAILED");
      expect(heard[0]).toContain("code=GeneralException");
      expect(heard[0]).toContain("Shape.name");
    } finally {
      vi.useRealTimers();
    }
  });

  it("says so when the host was merely slow, not broken", async () => {
    vi.useFakeTimers();
    const heard: string[] = [];
    onLateSync((m) => heard.push(m));
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      let finish!: () => void;
      vi.stubGlobal("PowerPoint", {
        ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
        run: async (cb: (ctx: unknown) => Promise<unknown>) =>
          cb({
            presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
            sync: () => new Promise<void>((res) => (finish = res)),
          }),
      });
      const p = insertSceneIntoSlide(buildChart(config), {});
      const assertion = expect(p).rejects.toThrow(/did not respond/);
      await vi.advanceTimersByTimeAsync(400_000); // past max(45s, shapes*3s)
      await assertion;
      finish();
      await vi.advanceTimersByTimeAsync(1);
      // "SUCCEEDED late" means the timeout is too short — a different bug from
      // a host that is actually broken, and the note has to distinguish them.
      expect(heard[0]).toContain("the host eventually SUCCEEDED");
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("added slides use the blank layout", () => {
  it("asks for the blank layout by TYPE, not by its localised name", async () => {
    // A slide added with no layout inherits the PREVIOUS slide's — on a fresh
    // deck that is the title slide, so an agenda lands on top of "Click to add
    // title" with the placeholder showing through. We draw everything
    // ourselves and want no placeholders.
    // The master here is Danish ("Tom"), which is the point: matching the name
    // "Blank" would silently do nothing for most of the world.
    installHost([makeSlide("s1")]);
    await insertAgendaSlides([
      buildAgendaScene(["Intro", "Body"], { highlight: 0 }),
      buildAgendaScene(["Intro", "Body"], { highlight: 1 }),
    ]);
    expect(addedWithLayout).toEqual(["layout-blank", "layout-blank"]);
  });

  it("uses it for the demo deck too", async () => {
    installHost([makeSlide("s1")]);
    await insertDemoDeck([{ scene: buildChart(config), tagData: "{}" }, { scene: buildChart(config) }]);
    expect(addedWithLayout).toEqual(["layout-blank", "layout-blank"]);
  });

  it("still adds slides on a host that exposes no masters", async () => {
    // Layout choice is a nicety; inserting is not. If the host will not tell us
    // its layouts, fall back to the inherited one rather than failing.
    const ctx = installHost([makeSlide("s1")]);
    (ctx.presentation as unknown as { slideMasters: unknown }).slideMasters = {
      load() {},
      get items(): never {
        throw new Error("masters unavailable on this host");
      },
    };
    await insertAgendaSlides([buildAgendaScene(["Intro"], { highlight: 0 })]);
    expect(addedWithLayout).toEqual([undefined]);
  });
});

describe("the wait budget scales with the work", () => {
  /** Park the sync so we can watch the clock without the host ever answering. */
  const parkedHost = (slide: FakeSlide) =>
    vi.stubGlobal("PowerPoint", {
      ...(globalThis as unknown as { PowerPoint: Record<string, unknown> }).PowerPoint,
      run: async (cb: (ctx: unknown) => Promise<unknown>) =>
        cb({
          presentation: { slides: { getItemAt: () => slide }, getSelectedSlides: () => ({ getItemAt: () => slide }) },
          sync: () => new Promise<void>(() => {}),
        }),
    });

  /** A scene of `n` trivial shapes — the budget is a function of the count. */
  const sceneOf = (n: number) => ({
    width: 400,
    height: 300,
    nodes: Array.from({ length: n }, (_, i) => ({ kind: "rect" as const, x: i, y: 0, w: 4, h: 4, fill: "#111111" })),
  });

  it("never hands the host more than a batch at once — THE bug", async () => {
    // Measured against real PowerPoint on the web: ~10 shapes insert instantly,
    // the 18-shape table element works, a 30-shape butterfly NEVER commits —
    // the sync simply stops answering and nothing lands. The same shapes go
    // onto off-screen slides by the hundred, because those are not painted.
    // So the fix was never a bigger timeout; it was a smaller batch.
    const slide = makeSlide("s1");
    installHost([slide]);
    const perSync: number[] = [];
    let last = 0;
    const ctx = installHost([slide]);
    ctx.sync = async () => {
      trips.syncs++;
      perSync.push(slide.created.length - last);
      last = slide.created.length;
    };
    const scene = buildChart(config);
    expect(scene.nodes.length).toBeGreaterThan(10); // must actually span batches
    await insertSceneIntoSlide(scene, { tagData: "{}" });
    expect(Math.max(...perSync), `handed over at once: ${perSync.join(",")}`).toBeLessThanOrEqual(10);
  });

  // The batching above counts SHAPES. A node is not a shape: a wedge fans into
  // triangles and a polygon becomes one line per edge, so the kinds that flood
  // the host are exactly the ones the all-rect `stacked` config cannot exercise.
  it.each(["pie", "doughnut", "sunburst", "radar", "violin"])(
    "batches %s by shapes, not nodes — the wedge/polygon flood",
    async (kind) => {
      const slide = makeSlide(`s-${kind}`);
      const perSync: number[] = [];
      let last = 0;
      const ctx = installHost([slide]);
      ctx.sync = async () => {
        trips.syncs++;
        perSync.push(slide.created.length - last);
        last = slide.created.length;
      };
      const scene = buildChart(sampleConfig(kind as never));
      await insertSceneIntoSlide(scene, { tagData: "{}" });
      // A single indivisible node (a wedge fan) may exceed the budget on its own;
      // nothing may exceed the host's measured breaking point of ~18.
      expect(Math.max(...perSync), `${kind} handed over at once: ${perSync.join(",")}`).toBeLessThanOrEqual(18);
    },
  );

  it("still bounds a trivial insert — the floor, not zero", async () => {
    vi.useFakeTimers();
    try {
      const slide = makeSlide("s1");
      installHost([slide]);
      parkedHost(slide);
      let settled = false;
      // 1 shape: the per-shape budget is tiny, so the 45s floor is what holds —
      // and 30s is past the old flat 20s, which is the thing that broke.
      const p = insertSceneIntoSlide(sceneOf(1), {}).catch(() => void (settled = true));
      await vi.advanceTimersByTimeAsync(30_000);
      expect(settled, "the floor keeps a small insert waiting past 30s").toBe(false);
      await vi.advanceTimersByTimeAsync(200_000);
      await p;
      expect(settled).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("EVERY insert path batches its shapes", () => {
  /** Max shapes handed to the host in any single sync of `run`. */
  async function maxPerSync(run: () => Promise<unknown>, slides: FakeSlide[]) {
    const worst = { n: 0 };
    let last = 0;
    const ctx = installHost(slides);
    // Wrap, don't replace: addSlides now verifies its adds via a settled
    // getCount() in a fresh context (see the addSlides retry/verify test),
    // which needs the real sync's committedCount/pendingCounts bookkeeping.
    // A bare replacement (as this used to be) freezes getCount() at its
    // initial value forever, which reads as "every add() was lost" and
    // starves insertAgendaSlides/insertDemoDeck of slides to render onto.
    const realSync = ctx.sync;
    const count = () => slides.reduce((a, s) => a + s.created.length, 0);
    ctx.sync = async () => {
      worst.n = Math.max(worst.n, count() - last);
      last = count();
      await realSync();
    };
    await run();
    return worst.n;
  }

  it("insert, update, agenda AND demo deck — none may send a whole scene", async () => {
    // The omission this exists for: I chunked insertSceneIntoSlide and
    // updateChartsInSlides and forgot insertAgendaSlides and insertDemoDeck.
    // The demo deck kept handing over ~200 shapes (4 slides at once) and sat at
    // "Working… 845s" having added nothing — and reported no progress, because
    // progress only fires when a chunk COMPLETES and the first never did.
    //
    // Live-canvas paths stay at ≤10 (repaints choke past that). Off-screen
    // append paths use a larger batch — the host tolerates far more when it
    // isn't repainting — but still bounded; they must NOT hand over the whole
    // scene, so a value at or under SHAPES_PER_SYNC_OFFSCREEN (40) is the
    // invariant, not the old flat 10.
    const scene = () => buildChart(config);
    expect(scene().nodes.length).toBeGreaterThan(10); // must span batches

    const s1 = makeSlide("s1");
    expect(
      await maxPerSync(() => insertSceneIntoSlide(scene(), { tagData: "{}" }), [s1]),
      "insert",
    ).toBeLessThanOrEqual(10);

    const s2 = makeSlide("s2");
    const old = s2.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    expect(
      await maxPerSync(
        () => updateChartInSlide(scene(), { slideId: "s2", shapeId: old.id, left: 0, top: 0 }, {}),
        [s2],
      ),
      "update",
    ).toBeLessThanOrEqual(11); // +1: the pre-existing shape this test planted

    const s3 = makeSlide("s3");
    expect(await maxPerSync(() => insertAgendaSlides([scene(), scene()]), [s3]), "agenda").toBeLessThanOrEqual(40);

    const s4 = makeSlide("s4");
    expect(
      await maxPerSync(
        () =>
          insertDemoDeck([
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
            { scene: scene() },
          ]),
        [s4],
      ),
      "demo deck",
    ).toBeLessThanOrEqual(40);
  });

  it("off-screen demo/agenda batches larger than the live canvas — cuts syncs per chart", async () => {
    // The live canvas caps at 10 shapes per batch (repaint mid-render kills the
    // host past that). Off-screen slides don't repaint, so demo/agenda push a
    // larger batch — 40 — cutting ~4x round-trips per chart. This asserts the
    // demo path ACTUALLY sends more than the live-canvas ceiling for a scene
    // that could fit in one 40-batch.
    const scene = () => ({
      width: 100,
      height: 100,
      nodes: Array.from({ length: 25 }, (_, k) => ({
        kind: "rect" as const,
        x: k,
        y: 0,
        w: 4,
        h: 4,
        fill: "#111111",
      })),
    });
    installHost([makeSlide("s-live")]);
    // Live canvas: caps at 10.
    const live = await maxPerSync(() => insertSceneIntoSlide(scene(), { tagData: "{}" }), [makeSlide("s-live")]);
    expect(live, "live canvas ≤10").toBeLessThanOrEqual(10);
    // Off-screen demo: uses ≥15 in some batch — proves the flat 10 cap is gone.
    installHost([makeSlide("s-off")]);
    const off = await maxPerSync(() => insertDemoDeck([{ scene: scene() }]), [makeSlide("s-off")]);
    expect(off, "off-screen batches larger than live").toBeGreaterThan(10);
  });
});

describe("a target whose slide is gone is nothing to do, not a crash", () => {
  it("skips a stale slideId instead of throwing InvalidParam", async () => {
    // The real error, from the real host:
    //   InvalidParam passed to GetItem(id) | code=5010
    //   errorLocation: SlideCollection.getItem
    // An EditTarget outlives the slide it names — delete the slide, undo, or
    // reopen the deck and the id is stale. getItem THROWS on that; it is a
    // normal condition wearing a crash's clothes. Same Scale over a deck would
    // take one deleted chart and lose every OTHER chart's rescale with it.
    const live = makeSlide("s-live");
    installHost([live]);
    const s = live.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 1, height: 1 });
    await expect(
      updateChartsInSlides([
        {
          scene: buildChart(config),
          target: { slideId: "s-deleted", shapeId: "gone", left: 0, top: 0 },
          opts: { tagData: "{}" },
        },
        {
          scene: buildChart(config),
          target: { slideId: "s-live", shapeId: s.id, left: 10, top: 20 },
          opts: { tagData: '{"ok":1}' },
        },
      ]),
      // One refreshed target back: the live chart's. The dead one contributes
      // nothing, and must not take the live one down with it.
    ).resolves.toHaveLength(1);
    // The live chart still got drawn and tagged — one dead target must not take
    // the others down.
    const group = live.created.find((c) => c.type === "group");
    expect(group, "the live chart was skipped too").toBeTruthy();
    expect(group!.tagStore.get(CHART_TAG)).toBe('{"ok":1}');
  });

  it("does nothing at all when every target is stale", async () => {
    const slide = makeSlide("s1");
    installHost([slide]);
    const before = slide.created.length;
    await expect(
      updateChartsInSlides([
        { scene: buildChart(config), target: { slideId: "nope", shapeId: "nope", left: 0, top: 0 }, opts: {} },
      ]),
    ).resolves.toEqual([]);
    expect(slide.created.length).toBe(before);
  });
});

/**
 * The repair pass, against the host fake — `src/core/reconcile.ts` decides
 * what to do, this is the half that talks to PowerPoint and has to survive it.
 */
describe("reading a demo deck back and repairing it", () => {
  /** A slide as a damaged run leaves it: some shapes, maybe a banner, maybe a group. */
  function demoSlide(
    id: string,
    opts: {
      slot?: { i: number; title: string };
      shapes?: number;
      stamped?: boolean;
      tagged?: boolean;
      grouped?: boolean;
    },
  ): FakeSlide {
    const slide = makeSlide(id);
    if (opts.slot) slide.tags.add(DEMO_SLOT_TAG, JSON.stringify(opts.slot));
    const parts: FakeShape[] = [];
    for (let i = 0; i < (opts.shapes ?? 0); i++) {
      const shape = slide.shapes.addTextBox(`n${i}`, { left: 0, top: 0, width: 10, height: 10 });
      shape.name = `part-${i}`;
      parts.push(shape);
    }
    if (opts.grouped && parts.length) {
      const group = slide.shapes.addGroup(parts);
      group.name = "PowerChart";
      if (opts.tagged) group.tags.add(CHART_TAG, `{"kind":"line"}`);
    } else if (opts.tagged && parts.length) {
      parts[parts.length - 1].tags.add(CHART_TAG, `{"kind":"line"}`);
    }
    if (opts.stamped) {
      const banner = slide.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 100, height: 20 });
      banner.name = "PowerChart:not-complete";
    }
    return slide;
  }

  const expect3 = (slot: number, title: string, chart = true) => ({ slot, title, shapes: 3, chart });

  /** A two-point line chart — small enough that one batch draws the whole thing. */
  const tinyChart = (): ChartConfig => ({
    ...sampleConfig("line"),
    title: "Line",
    data: { categories: ["A", "B"], series: [{ name: "S1", values: [1, 2] }] },
  });

  it("reads shape counts, banners, slot tags and group children off the deck", async () => {
    const deck = [
      demoSlide("s0", { shapes: 1 }),
      demoSlide("s1", { slot: { i: 0, title: "Line" }, shapes: 3, tagged: true }),
      demoSlide("s2", { slot: { i: 1, title: "Gantt" }, shapes: 3, stamped: true }),
      demoSlide("s3", { slot: { i: 2, title: "Pie" }, shapes: 3, grouped: true, tagged: true }),
    ];
    installHost(deck);
    const snaps = await snapshotAddedSlides(0, 4);
    expect(snaps.map((s) => s.slot)).toEqual([null, 0, 1, 2]);
    expect(snaps.map((s) => s.title)).toEqual([null, "Line", "Gantt", "Pie"]);
    expect(snaps[1]).toMatchObject({ shapes: 3, stamped: false, tagged: true });
    expect(snaps[2]).toMatchObject({ shapes: 4, stamped: true, tagged: false });
    // The group is one top-level shape; without its child count a complete
    // chart would read as 1 of 3 and be condemned as wreckage.
    expect(snaps[3]).toMatchObject({ grouped: true, groupChildren: 3, tagged: true });
  });

  it("deletes a duplicate slide, clears a stale banner, and re-groups a loose chart", async () => {
    // The shape of Presentation_4.pptx: a clean chart, the same chart again
    // under a NOT COMPLETE banner, and an empty slide the host left behind.
    const deck = [
      demoSlide("title", { slot: { i: 0, title: "Title" }, shapes: 3, grouped: true }),
      demoSlide("line", { slot: { i: 1, title: "Line" }, shapes: 3 }),
      demoSlide("line-dup", { slot: { i: 1, title: "Line" }, shapes: 3, stamped: true }),
      demoSlide("stray", {}),
    ];
    installHost(deck);
    const outcome = await reconcileDeck(
      [expect3(0, "Title", false), expect3(1, "Line")],
      { before: 0, after: 4 },
      () => `{"kind":"line"}`,
      { dropOrphanBlanks: true },
    );
    expect(outcome.applied).toEqual({ unstamped: 0, regrouped: 1, deleted: 2 });
    expect(outcome.refused).toBe(0);
    expect(deck.map((s) => s.id)).toEqual(["title", "line"]);
    // The surviving chart is now a group carrying the config — re-editable.
    const group = deck[1].created.filter((s) => !s.deleted).find((s) => s.name === "PowerChart");
    expect(group?.tagStore.get(CHART_TAG)).toBe(`{"kind":"line"}`);
  });

  it("pulls the banner off a chart that is in fact complete", async () => {
    const deck = [demoSlide("agenda", { slot: { i: 0, title: "Agenda" }, shapes: 3, stamped: true })];
    installHost(deck);
    const outcome = await reconcileDeck([expect3(0, "Agenda", false)], { before: 0, after: 1 }, () => undefined, {});
    expect(outcome.applied.unstamped).toBe(1);
    expect(deck[0].created.filter((s) => !s.deleted).some((s) => s.name === "PowerChart:not-complete")).toBe(false);
  });

  it("deletes from the end, so an earlier delete cannot renumber a later one", async () => {
    // Ascending deletes would remove index 1, shifting B's duplicate into the
    // slot the plan meant for something else — and take a good slide with it.
    const deck = [
      demoSlide("a", { slot: { i: 0, title: "A" }, shapes: 3 }),
      demoSlide("a-dup", { slot: { i: 0, title: "A" }, shapes: 3, stamped: true }),
      demoSlide("b", { slot: { i: 1, title: "B" }, shapes: 3 }),
      demoSlide("b-dup", { slot: { i: 1, title: "B" }, shapes: 3, stamped: true }),
    ];
    installHost(deck);
    const outcome = await reconcileDeck(
      [expect3(0, "A"), expect3(1, "B")],
      { before: 0, after: 4 },
      () => undefined,
      {},
    );
    expect(outcome.applied.deleted).toBe(2);
    expect(deck.map((s) => s.id)).toEqual(["a", "b"]);
  });

  it("leaves a slide from an unrelated deck untouched", async () => {
    const deck = [demoSlide("theirs", { shapes: 2 }), demoSlide("ours", { slot: { i: 0, title: "Line" }, shapes: 3 })];
    installHost(deck);
    const outcome = await reconcileDeck([expect3(0, "Line")], { before: 0, after: 2 }, () => undefined, {});
    expect(outcome.plan.orphans.map((o) => o.index)).toEqual([0]);
    expect(deck.map((s) => s.id)).toEqual(["theirs", "ours"]);
  });

  it("never groups the NOT COMPLETE banner in with the chart", async () => {
    // A real run shipped a Line chart whose group held 37 shapes: 36 of them
    // the chart, one a red NOT COMPLETE stripe. Once inside, the banner is
    // invisible to every later repair — a snapshot reads top-level names — so
    // it rides along with the chart forever.
    // 18 of 20 shapes: enough to re-group, not enough to count as complete,
    // so the plan asks for a regroup with the banner still on the slide.
    const deck = [demoSlide("partial", { slot: { i: 0, title: "Line" }, shapes: 18, stamped: true })];
    installHost(deck);
    await reconcileDeck(
      [{ slot: 0, title: "Line", shapes: 20, chart: true }],
      { before: 0, after: 1 },
      () => undefined,
      {},
    );
    const live = deck[0].created.filter((s) => !s.deleted);
    const group = live.find((s) => s.name === "PowerChart")!;
    expect(group.grouped).toHaveLength(18);
    expect((group.grouped as { name?: string }[]).some((c) => c.name === "PowerChart:not-complete")).toBe(false);
  });

  it("clears the banner on a slide whose chart is already grouped", async () => {
    // NOTE ON WHAT THIS DOES AND DOES NOT PROVE: in this fake a grouped shape
    // stays visible at the top level, so the repair finds the banner there.
    // On a real host a group hides its children, and the deck from the
    // 2026-07-31 run has a banner buried INSIDE a chart's group — `deleteStamp`
    // reaches into the group for exactly that case, and that reach is NOT
    // exercised here. `test/reconcile.test.ts` covers the planning half
    // (a grouped, stamped slide still gets an unstamp action).
    const deck = [demoSlide("buried", { slot: { i: 0, title: "Line" }, shapes: 3, stamped: true })];
    const shapes = deck[0].created;
    const group = deck[0].shapes.addGroup(shapes.slice());
    group.name = "PowerChart";
    installHost(deck);
    const outcome = await reconcileDeck(
      [{ slot: 0, title: "Line", shapes: 3, chart: true }],
      { before: 0, after: 1 },
      () => undefined,
      {},
    );
    expect(outcome.applied.unstamped).toBe(1);
    expect(deck[0].created.filter((s) => !s.deleted).some((s) => s.name === "PowerChart:not-complete")).toBe(false);
  });

  it("stamps nothing when the item's own slide never landed", async () => {
    // `stampLastSlide` used to brand whatever was last in the deck. When the
    // host swallowed the add, that was the PREVIOUS item's slide — a real run
    // defaced a KPI tile that had rendered perfectly, because a results page
    // that never landed stamped it.
    const existing = makeSlide("theirs");
    existing.shapes.addTextBox("someone else's work", { left: 0, top: 0, width: 10, height: 10 });
    const deck: FakeSlide[] = [existing];
    installHost(deck);
    swallowAdds = 4; // both the add and its retry vanish
    _setBatchTimeoutForTest(5); // no slide to draw on — do not wait 45s for it
    try {
      // Nothing rendered, so insertDemoDeck rethrows the host's own error —
      // the point here is what it did NOT do on the way out.
      await expect(insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }])).rejects.toThrow();
      expect(existing.created.some((s) => s.name === "PowerChart:not-complete")).toBe(false);
    } finally {
      _setBatchTimeoutForTest(45_000);
    }
    // Generous: an earlier test in this file can leave an abandoned sync
    // outstanding, and the failure path waits up to 5s for it to report.
  }, 20_000);

  it("stops drawing shapes and inserts pictures once the host falls behind", async () => {
    // We cannot catch the crash — the tab dies, there is no rejected promise —
    // so the run watches what comes BEFORE it. A budget of zero shapes stands
    // in for a host that has already spent its allowance.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const scene = buildChart(tinyChart());
    const report = await insertDemoDeck(
      [
        { scene, title: "One", tagData: `{"i":0}` },
        { scene, title: "Two", tagData: `{"i":1}` },
        { scene, title: "Three", tagData: `{"i":2}` },
      ],
      undefined,
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 0 },
    );
    expect(report.degradedAt).toBe(1);
    expect(report.degradeReason).toMatch(/shapes drawn/);
    // Item 1 drew its shapes; items 2 and 3 are one picture each.
    expect(report.results[0].created).toBeGreaterThan(1);
    expect(report.results[1].created).toBe(1);
    expect(report.results[2].created).toBe(1);
    // And the picture carries the config tag, so the chart is still editable.
    const last = deck[deck.length - 1].created.filter((s) => !s.deleted);
    expect(last.some((s) => s.imageBase64 === "iVBORw0KGgo=")).toBe(true);
  }, 20_000);

  it("keeps drawing shapes when the host is keeping up", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck(
      [{ scene: buildChart(tinyChart()), title: "One", tagData: `{"i":0}` }],
      undefined,
      { pictureFor: async () => "iVBORw0KGgo=", shapeBudget: 10_000 },
    );
    expect(report.degradedAt).toBeUndefined();
    expect(report.results[0].created).toBeGreaterThan(1);
  }, 20_000);

  it("never degrades when the caller offers no picture to fall back to", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "One" }], undefined, {
      shapeBudget: 0,
    });
    expect(report.degradedAt).toBeUndefined();
  }, 20_000);

  it("traces the run in enough detail to diagnose it afterwards", async () => {
    // Every hard thing in this project was diagnosed after the fact, from a
    // deck and a one-line summary. This is the record that was missing: which
    // item, what landed, what the host refused, and when it was given up on.
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    setTracing(true);
    try {
      await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line", tagData: `{"i":0}` }]);
      const { entries } = traceLog();
      const byScope = (scope: string) => entries.filter((e) => e.scope === scope);
      // The per-item verdict, with the numbers a reader needs to judge it.
      const item = byScope("demo").find((e) => e.message === "item finished");
      expect(item?.data).toMatchObject({ i: 0, title: "Line", status: "rendered" });
      expect(item?.data?.created).toBeGreaterThan(1);
      // And the drawing itself, batch by batch — "died at batch 1 of 4" was
      // the entire diagnosis of the update stall.
      expect(byScope("draw").length).toBeGreaterThan(0);
      expect(byScope("draw")[0].data).toHaveProperty("total");
    } finally {
      setTracing(false);
    }
  }, 20_000);

  it("costs a run nothing while it is off", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    setTracing(false);
    // Switching off deliberately KEEPS the log readable, so measure growth
    // rather than emptiness — the claim is that a disabled trace records
    // nothing new, not that it forgets what it already had.
    const before = traceLog().entries.length;
    await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }]);
    expect(traceLog().entries).toHaveLength(before);
  }, 20_000);

  it("closes a demo run with the settled truth when asked for it", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck(
      [{ scene: buildChart(tinyChart()), title: "Line", tagData: `{"i":0}` }],
      undefined,
      { reconcile: true },
    );
    expect(report.reconcile).toBeDefined();
    expect(report.reconcile?.plan.verdicts[0]).toMatchObject({ title: "Line", status: "rendered" });
    // A clean run needs no repair, and says so instead of inventing work.
    expect(report.reconcile?.plan.actions).toEqual([]);
    expect(report.blankSlides).toEqual([]);
  });

  it("does not reconcile unless the caller opts in", async () => {
    const deck: FakeSlide[] = [makeSlide("s1")];
    installHost(deck);
    const report = await insertDemoDeck([{ scene: buildChart(tinyChart()), title: "Line" }]);
    expect(report.reconcile).toBeUndefined();
  });
});
