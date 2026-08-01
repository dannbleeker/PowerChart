/**
 * A fake PowerPoint host: recording doubles for the Office.js proxy-object API.
 *
 * Extracted from `office-render.test.ts`, which owned it for the whole of this
 * project's history and could not share it. Every host bug found against a
 * real PowerPoint has been paid for with a deploy, a run someone had to sit
 * through, and an uploaded .pptx — and each fix then added one more knob here
 * so the same bug could never come back. Those knobs are worth more than one
 * test file: pinned together they describe a host far nastier than any suite
 * had been running against, which is why the next bug class should be findable
 * without leaving CI.
 */
import { vi } from "vitest";
// @ts-expect-error — the .mjs auditor has no types. Deliberately the SAME
// decoder `npm run verify-deck` uses: a fake that parsed the generated deck
// its own way could agree with the renderer while both disagreed with the file.
import { readDeckBytes } from "../../scripts/verify-deck.mjs";

interface DeckRow {
  slot: number | null;
  title: string | null;
  run: string | null;
  configJson: string | null;
  chartShapes: number;
}

const DEMO_SLOT_TAG_KEY = "POWERCHART_DEMO_SLOT";
const CHART_TAG_KEY = "POWERCHART_CONFIG";

/**
 * The host's misbehaviours, in one mutable object.
 *
 * These were module-scope `let`s inside the test file that owned this harness,
 * which is exactly why the harness could not move: an importing module cannot
 * assign to another module's binding. Gathering them here is what lets any
 * suite drive the fake host — and lets a profile turn a set of them on at once
 * (see `WEB_PROFILE`).
 *
 * Every one of them was added AFTER a real PowerPoint taught us the behaviour
 * it models. That is the pattern this file exists to break: the next such
 * lesson should be learnable in CI.
 */
export const faults = {
  failSyncOn: 0,
  swallowAdds: 0,
  faultShapeGetCount: false,
  strictGroup: false,
  strictTags: false,
  hollowReads: 0,
  refusePictureFill: false,
  refuseGroups: 0,
  /**
   * Whole decks the host accepts and never lands — the file path's version of
   * `swallowAdds`, and observed for real: a 12-item file insert came back
   * "11 of 12 complete · 1 lost" with nothing thrown. Counted down per call.
   */
  swallowDecks: 0,
};

/**
 * PowerPoint on the web, at its worst — every observed misbehaviour at once.
 *
 * Not the default: applied to every existing test it would fail hundreds of
 * them for the wrong reason, asserting nothing about the code under test. It
 * is a named profile, and the demo path is pinned to it (`web-host.test.ts`)
 * because the demo path is where all of this was found.
 *
 * `swallowAdds` and `failSyncOn` are deliberately NOT here: they are one-shot
 * counters aimed at a particular sync, and a blanket value would decide which
 * call fails rather than letting the test say.
 *
 * Call AFTER `installHost` — installing a host resets every fault to its
 * well-behaved default, so applying the profile first would be undone.
 */
export function applyWebProfile(): void {
  // A shape proxy older than one sync is rejected — the web host rewrites
  // proxies as `shapes.getItem(id)` and answers InvalidParam for a stale one.
  faults.strictGroup = true;
  faults.strictTags = true;
  // Shape collections read back shorter than they are, without throwing.
  faults.hollowReads = 3;
  // The first grouping attempt is refused outright.
  faults.refuseGroups = 1;
  // NOT refusePictureFill. The web host supports PowerPointApi 1.8 and takes
  // pictures — that support is the whole reason the degraded-picture path
  // exists there. Refusing it models a DIFFERENT host, and it silently
  // switches off the one path where the worst of the stale-proxy bugs lived:
  // a degraded chart is a single shape, and a single shape is what the tag
  // write had no fresh proxy for.
}

/**
 * Recording doubles for the PowerPoint JS proxy-object API: every shape the
 * renderer creates is captured with the geometry/format calls made on it, so
 * the whole scene→native-shapes mapping is testable without an Office host.
 */

let idSeq = 0;

export function makeShape(
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
      add: (k: string, v: string) => {
        // The web host rewrites a proxy as `shapes.getItem(id)` and rejects it
        // once stale — the SAME trap addGroup falls into, seen in the wild as
        // `InvalidParam passed to GetItem(id)`, code 5010, at
        // `ShapeCollection.getItem`, 28 times in one 38-item run. Tagging
        // crosses a sync boundary just like grouping does.
        if (faults.strictTags && trips.syncs > shape.syncCreated + 1) {
          throw new Error("InvalidParam passed to GetItem(id) | code=5010");
        }
        tagStore.set(k, v);
      },
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
        // type to PictureAndTexture" (@types/office-js). `faults.refusePictureFill`
        // makes the property access itself throw, which is what a pre-1.8 host
        // does — the method simply is not there.
        if (faults.refusePictureFill) throw new Error("setImage is not available on this host");
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

export type FakeShape = ReturnType<typeof makeShape>;

/** Layout ids passed to slides.add() since the last installHost(). */
export const addedWithLayout: (string | undefined)[] = [];

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

/**
 * A fresh proxy onto a shape, as `shapes.items` hands back after a load+sync.
 *
 * Everything reads and writes through to the one real shape — a tag added via
 * a fresh handle is on the shape, because it is the same shape. What the
 * handle does NOT share is `syncCreated`: this one was minted now, so it is
 * current, while whatever handle the caller was holding before stays as old as
 * it was. That is the whole distinction, and modelling it is what lets a test
 * tell "re-fetched the collection and used the new handle" apart from
 * "re-fetched the collection and went on using the old one".
 */
function freshHandle(shape: FakeShape): FakeShape {
  let own = trips.syncs;
  return new Proxy(shape, {
    get(target, prop, recv) {
      if (prop === "syncCreated") return own;
      // `tags` has to be rebound too. The shape's own tags object closes over
      // the shape's age, so a fresh handle reaching for `.tags` would get a
      // writer that still consults the STALE number and refuses — the handle
      // would look fresh and behave stale. Everything else (the tag store, the
      // geometry, the delete) is deliberately shared: it is one shape.
      if (prop === "tags") {
        return {
          add: (k: string, v: string) => {
            if (faults.strictTags && trips.syncs > own + 1) {
              throw new Error("InvalidParam passed to GetItem(id) | code=5010");
            }
            target.tagStore.set(k, v);
          },
          getItemOrNullObject: target.tags.getItemOrNullObject,
        };
      }
      return Reflect.get(target, prop, prop === "syncCreated" ? recv : target);
    },
    set(target, prop, value) {
      if (prop === "syncCreated") {
        own = value as number;
        return true;
      }
      return Reflect.set(target, prop, value);
    },
  });
}

export function makeSlide(id: string) {
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
      //
      // Reading the collection hands back FRESH HANDLES, and leaves every
      // handle taken earlier exactly as stale as it was. This is the one place
      // the fake used to be kinder than the host it models: it refreshed
      // `syncCreated` on the shape objects themselves, so a re-fetch anywhere
      // healed a stale proxy held anywhere else. Real Office.js returns new
      // proxy objects — your old variable still points at the old proxy, and
      // the old proxy is still refused.
      //
      // That difference is not academic. It is precisely why a whole class of
      // stale-proxy bug was invisible here and had to be found by a human
      // running the add-in in a real PowerPoint: the code could re-fetch a
      // collection, tag the WRONG (old) handle, and this fake would let it
      // through. Restoring the distinction is what makes `web-host.test.ts`
      // able to catch it.
      get items() {
        const live = created.filter((s) => !s.deleted).map(freshHandle);
        // The web host has been observed answering a shape-collection read
        // with FAR fewer shapes than it holds: one readback page asked about
        // 19 slides carrying 19 shapes and got 3 back. `faults.hollowReads` models
        // that — the collection is short, nothing throws, and the caller has
        // no way to know unless it compares against a count it took earlier.
        if (faults.hollowReads > 0 && lastShapeLoad === "items/id") {
          faults.hollowReads--;
          return [];
        }
        return live;
      },
      // The tag pass is the only reader that asks for `items/id`; pass A and
      // the group-child count both ask for `items/name`. Recording which lets
      // a test make the TAG read hollow without also blinding the count it is
      // supposed to be checked against.
      load(p?: string) {
        if (p) lastShapeLoad = p;
      },
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
        if (faults.refuseGroups > 0) {
          faults.refuseGroups--;
          throw new Error("host refused addGroup");
        }
        // Model the web-host stale-proxy trap: a Shape proxy is valid within
        // the sync that queued it plus its immediately following commit sync,
        // and stale (getItem(id) rewrite, non-round-trippable id) beyond that.
        // addGroup(theseStaleProxies) silently loses grouping on real Office —
        // no group appears on the slide. Enable via faults.strictGroup.
        if (faults.strictGroup && items.some((s) => trips.syncs > s.syncCreated + 1)) {
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
        if (faults.faultShapeGetCount) throw new Error("readback getCount faulted");
        return { value: created.filter((s) => !s.deleted).length };
      },
    },
  };
  return slide;
}

export type FakeSlide = ReturnType<typeof makeSlide>;

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
export const trips = { syncs: 0, contexts: 0 };

/** Proxy objects the renderer released via untrack(), by kind. */
export const untracked = { shapes: 0, tags: 0 };

/**
 * Make the Nth context.sync() of the next run throw. Office.js queues commands
 * and only reports their errors at sync — so this, not a throwing addGroup, is
 * how a host actually refuses something. 0 = never.
 */

/** Like faults.failSyncOn but a SET of sync indices — models a stall that persists across
 * the retry (fail both attempts' syncs), so a slide is truly lost, not recovered. */
export const failSyncsOn = new Set<number>();

/** Make the next N slides.add() calls no-ops — models PowerPoint web silently
 * dropping a slide-add under load, the corruption the self-check must catch. */

/** Deck indices whose shapes.getCount() reports 0 on readback — models a slide
 * that committed but came back BLANK (its shapes detached), the silent partial
 * the on-slide readback must catch. */
export const blankReadbackAt = new Set<number>();

/** When true, shapes.getCount() throws — models the blank readback faulting, so
 * the report must come back blanksRead:false rather than an empty "no blanks". */

/** Sync indices that STALL — the promise resolves after `stallSyncDelayMs`, but
 * withTimeout (shortened via _setBatchTimeoutForTest) fires first and abandons
 * it. Models the real-host bug: sync at 60s, shapes on the slide, no error to
 * blame. When the abandoned promise later settles, withTimeout records
 * `lastLateSync = "…SUCCEEDED after Ns"` — the signal PR 1 gates on. */
export const stallSyncOn = new Set<number>();
const stallSyncDelayMs = 40;

/** When true, addGroup silently drops the group if any member proxy is more
 * than one sync stale — the exact web-host behavior that made multi-batch
 * charts land ungrouped. See PR 3's re-fetch fix.  */

/** When true, `tags.add` rejects a proxy more than one sync old — see makeShape. */

/** Make the next N addGroup calls THROW — a host that declines to group even
 * though the API is there. Decremented per call, so refusing 1 leaves a chart
 * rendered-but-ungrouped and still lets the fresh-context rescue group it. */

/** Shape-collection reads that come back EMPTY without throwing — see `items`. */
let lastShapeLoad = "";

/** When true, `fill.setImage` throws on access — a pre-1.8 host, where the
 * method does not exist at all. Drives the picture-insert fall-through. */

/** Install a fake PowerPoint global whose run() drives the mocked context.
 * `supported(version)` models the host's requirement-set support (default: all)
 * — pass a predicate to simulate e.g. PowerPoint on the web lacking grouping. */
export function installHost(
  slides: FakeSlide[],
  selectedShapes: FakeShape[] = [],
  selectedSlideArg = slides[0],
  supported: (version: string) => boolean = () => true,
) {
  // The slide count as of the last COMMITTED sync. getCount() reports THIS, not
  // the live array — so an add() queued in the current batch is invisible to a
  // getCount() in the SAME batch, exactly as PowerPoint web behaves. A getCount
  // result resolves at the next sync to the count from before that sync's adds.
  let selectedSlide = selectedSlideArg;
  let committedCount = slides.length;
  /** Decks handed to insertSlidesFromBase64 and not yet resolved by a sync. */
  const pendingDecks: string[] = [];
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
          if (faults.swallowAdds > 0) {
            faults.swallowAdds--; // the host dropped this add — no slide appears
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
      setSelectedSlides: (ids: string[]) => {
        const found = slides.find((sl) => ids.includes(sl.id));
        if (found) selectedSlide = found;
      },
      /**
       * The whole-deck insert — and the path a real run actually takes.
       *
       * The fake could not model it at all, so every test of the default path
       * either skipped or ran against the shape-by-shape one instead. Here the
       * bytes are really decoded (by the same auditor `npm run verify-deck`
       * uses, so the fake cannot disagree with the tool) and each slide in the
       * file becomes a slide in the deck, carrying its slot tag and a single
       * `PowerChart` shape holding the config — which is what the generator
       * genuinely writes.
       *
       * Queued, not immediate: Office.js resolves this at the next sync, and a
       * fake that appended straight away would hide every ordering bug the
       * queue creates.
       */
      insertSlidesFromBase64: (b64: string) => {
        if (faults.swallowDecks > 0) {
          faults.swallowDecks--; // taken, acknowledged, and never landed
          return;
        }
        pendingDecks.push(b64);
      },
    },
    sync: async () => {
      trips.syncs++;
      for (const b64 of pendingDecks.splice(0)) {
        const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
        const { rows } = (await readDeckBytes(bytes)) as { rows: DeckRow[] };
        for (const row of rows) {
          const added = makeSlide(`slide-${slides.length + 1}`);
          if (row.slot !== null) {
            added.tagStore.set(DEMO_SLOT_TAG_KEY, JSON.stringify({ i: row.slot, title: row.title, run: row.run }));
          }
          // One group named PowerChart, tagged, holding as many children as
          // the file really holds — what the generator writes.
          //
          // The child COUNT is the load-bearing part. A readback measures a
          // chart by how many shapes are inside its group, and a fake that put
          // one shape there made every generated chart read as wreckage: the
          // repair pass then had no healthy copy to prefer and queued no
          // duplicates, so the duplicate-slot scenario reported a failure that
          // was the fake's, not the code's.
          const shape = added.shapes.addGeometricShape("rectangle", { left: 0, top: 0, width: 10, height: 10 });
          shape.name = "PowerChart";
          shape.grouped = Array.from({ length: Math.max(1, row.chartShapes) }, (_, k) =>
            makeShape("geometric", "rectangle", { left: k, top: 0, width: 1, height: 1 }),
          );
          if (row.configJson) shape.tagStore.set(CHART_TAG_KEY, row.configJson);
          added.pending.length = 0;
          slides.push(added);
        }
      }
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
      if (trips.syncs === faults.failSyncOn || failSyncsOn.has(trips.syncs)) {
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
  faults.swallowAdds = 0;
  failSyncsOn.clear();
  stallSyncOn.clear();
  faults.strictGroup = false;
  faults.strictTags = false;
  faults.refuseGroups = 0;
  faults.hollowReads = 0;
  lastShapeLoad = "";
  faults.refusePictureFill = false;
  blankReadbackAt.clear();
  faults.faultShapeGetCount = false;
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
