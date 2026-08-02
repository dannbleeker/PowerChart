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
  /**
   * `shapes.load(...)` throws a TypeError instead of queueing.
   *
   * Observed on the web as "e.load is not a function": the host handed back
   * something without the method, and because the QUEUE step sat outside the
   * try/catch that was written to cover it, an update whose shapes had already
   * committed reported total failure. Charts on the slide, one TypeError, no
   * result — which is what Same Scale across the deck did.
   */
  faultShapeCollectionLoad: false,
  /**
   * A shape's `id`/`left`/`top` are unreadable until a `load()` on it has been
   * ANSWERED by a sync — which is simply how Office.js works, and how this fake
   * has never worked.
   *
   * Off by default, and only because the fake's shape objects double as the
   * surface tests assert geometry against: turning it on globally would fail
   * hundreds of tests on their own reads rather than on the code under test.
   * Where it IS on, it catches the class of bug that a lenient shape hides
   * completely — reading a position off a proxy whose load went down with a
   * failed sync, which is `PropertyNotLoaded` at `Shape.left` on a real host
   * and a plain number here. That is how a deck-wide rescale came to report a
   * crash for charts that were, on the slide, correctly redrawn.
   *
   * A test may switch it back off before inspecting the fake's own state: the
   * gate models a HOST read, not the bookkeeping a test does afterwards.
   */
  strictShapeReads: false,
  /**
   * Tag loads the host takes and never answers — the next N `load()`s on a tag
   * proxy leave it unreadable after their sync.
   *
   * The tag twin of `unansweredShapeReads`, and the one with the widest blast
   * radius: `isNullObject`/`value` on a config tag is read on every path that
   * asks "is this shape a chart" — the selection readers, the deck scan, the
   * repair pass. Each of those used to read both properties raw, so a tag the
   * host stayed quiet about did not mean "not a chart", it meant the call
   * threw. Counted down per load, so a test can silence one read and watch the
   * rest of a scan carry on.
   */
  unansweredTagLoads: 0,
  /**
   * `setSelectedShapes([])` does not clear the selection — PowerPoint on the
   * web, office-js#3083. Desktop clears it.
   *
   * Worth modelling because of what it costs downstream rather than what it
   * costs at the call: on the web a picture cannot be inserted while a shape is
   * selected (office-js#3698), so a battery whose scenario leaves a chart
   * selected fails the NEXT scenario instead of its own. That is the hardest
   * kind of failure to read in a run log, and it is only reachable here if the
   * fake refuses the clear the way the web host does.
   */
  webIgnoresDeselect: false,
  /**
   * `setSelectedShapes(ids)` selects SOMETHING, but not what was asked for.
   *
   * A host that takes the call and gets the shape wrong is worse than one that
   * refuses: the pane then edits a chart the user did not click. The scenario
   * that covers the selection round trip has to be able to see that, and it
   * can only see it if the fake is capable of doing it.
   */
  selectionIgnoresIds: false,
  /**
   * A programmatic `setSelectedShapes(ids)` **wedges the selection subsystem**.
   *
   * Measured on PowerPoint on the web: the call itself is taken — no error, no
   * refusal, the sync resolves — and then every selection call after it goes
   * silent. `getSelectedShapes` ran out a full 90-second budget, and so did the
   * `setSelectedSlides` behind it. Nothing throws. The host simply stops
   * answering, which is the one failure shape a fake that only ever throws or
   * only ever lies cannot produce.
   *
   * The third of the web host's selection bugs, after #3083 and #3698, and the
   * reason the battery's selection scenario reports *skipped* rather than red
   * there. That gate is only worth having if CI can watch it work, and it can
   * only watch it work against a host that behaves this way.
   *
   * Deliberately NOT in `applyWebProfile()`, though it belongs to the same real
   * host: it is the one fault whose cost is measured in seconds rather than in
   * assertions. Every test using the profile would pay two full budgets to
   * reach a verdict it was not asking about.
   */
  selectionWedgesHost: false,
  /**
   * `Slide.getImageAsBase64` answers with the same bytes whatever is on the
   * slide — a host that hands back a blank render, or a cached one.
   *
   * The visibility scenario exists to catch a chart that is structurally
   * perfect and invisible. Against a host like this its comparison is
   * meaningless, and it must report that rather than pass.
   */
  constantSlideImage: false,
  /**
   * `Slide.delete()` is taken and does nothing.
   *
   * The self-test's visibility scenario is the only one that borrows a slide
   * and gives it back, so it is the only one that can leave litter. A host that
   * refuses the delete without saying so leaves a tagged chart in the deck under
   * a verdict that reads clean — reporting success while leaving a mess, which
   * is the pattern this project keeps finding. The scenario has to notice.
   */
  refuseSlideDelete: false,
  /**
   * The next N shapes to be tagged answer `.tags` as **undefined**.
   *
   * Observed on a real host, four times in one run, each on a chart whose
   * grouping had just been refused with InvalidParam 5010. Reading `.add` off
   * it throws SYNCHRONOUSLY, which is what made it so expensive: it escaped
   * the tagging loop rather than failing one chart, so every chart after it in
   * the batch lost its config without being attempted. Counted down per read.
   */
  tagsUndefinedOn: 0,
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
  // setSelectedShapes([]) is ignored — office-js#3083.
  faults.webIgnoresDeselect = true;
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

/**
 * The slide size this fake deck declares, in EMU. Defaults to PowerPoint's
 * canonical 16:9; a test sets it to 9144000×6858000 for 4:3.
 *
 * Drives BOTH reads the renderer can make — `pageSetup` and the `<p:sldSz>` in
 * an exported slide — so a test cannot accidentally assert against a size only
 * one of the two rungs would report.
 */
export const hostSlideSize = { cx: 12192000, cy: 6858000 };

/** Exports whose base64 is built during the sync that was asked for it. */
const pendingExports: { result: { value: string }; build: () => Promise<string> }[] = [];

/**
 * The host's current SHAPE selection, shared by everything that touches it.
 *
 * One array, mutated in place, because `Slide.setSelectedShapes` (the writer)
 * and `Presentation.getSelectedShapes` (the reader) are on different classes
 * and a test needs the round trip between them to be real. `installHost` seeds
 * it; the slide method writes it; the presentation getter reads it.
 */
const selectionRef: FakeShape[] = [];

/**
 * Set once a programmatic select has wedged the host — see
 * `faults.selectionWedgesHost`. Sticky for the rest of the run, because on the
 * real host it was: nothing un-wedged it short of a reload.
 */
let selectionWedged = false;
/** A selection call joined the sync now being built, and it will never land. */
let wedgeThisSync = false;

/** Called by every selection entry point once the subsystem is wedged. */
function noteSelectionCall(): void {
  if (selectionWedged) wedgeThisSync = true;
}

/**
 * A minimal .pptx carrying nothing but a `ppt/presentation.xml` with the
 * deck's `<p:sldSz>` — which is the only part `slideSize` reads.
 */
async function exportedDeckBase64(): Promise<string> {
  const { default: JSZip } = await import("jszip");
  const zip = new JSZip();
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:sldSz cx="${hostSlideSize.cx}" cy="${hostSlideSize.cy}"/>` +
      `</p:presentation>`,
  );
  return zip.generateAsync({ type: "base64" });
}

/**
 * Tag proxies whose `load()` has been queued but whose sync has not yet run.
 *
 * Flushed by a successful sync, dropped by a failed one — a load whose sync
 * rejects never populates anything, same as the real host.
 */
const pendingTagLoads: (() => void)[] = [];

/**
 * Shape `load()`s awaiting their sync — the same contract as `pendingTagLoads`,
 * for the properties a chart's tagged shape is identified by. Only read when
 * `faults.strictShapeReads` is on; queued always, so arming the fault
 * mid-render cannot grant a load that never landed.
 */
const pendingShapeLoads: (() => void)[] = [];

/**
 * A tag proxy that answers like Office.js does: resolving one gives you an
 * object whose properties are NOT readable until `load()` AND a sync.
 *
 * This fake used to populate `isNullObject` and `value` the instant
 * `getItemOrNullObject` was called, which made a whole class of bug invisible
 * here. Reading an unloaded proxy is `PropertyNotLoaded` on a real host, and
 * that is precisely how editing an ungrouped chart came to be impossible on
 * PowerPoint web while every test in this repo stayed green: `partIds` proxies
 * were resolved and never loaded, so the fake answered and the host threw.
 *
 * One honest proxy found the next one immediately (`deleteSlide`'s slot-tag
 * read). Being kinder than the host you model is not a neutral simplification —
 * it is a promise the tests cannot keep.
 */
function makeTagProxy(store: Map<string, string>, key: string, onUntrack?: () => void) {
  let readable = false;
  const need = (prop: string) => {
    if (!readable) {
      throw new Error(
        `PropertyNotLoaded: The property '${prop}' is not available. ` +
          "Before reading the property's value, call the load method on the containing object " +
          "and call context.sync() on the associated request context.",
      );
    }
  };
  return {
    load(_p?: string) {
      // Taken and never answered — see `faults.unansweredTagLoads`. The load
      // is still queued, so the sync carries it; nothing comes back.
      if (faults.unansweredTagLoads > 0) {
        faults.unansweredTagLoads--;
        return;
      }
      pendingTagLoads.push(() => {
        readable = true;
      });
    },
    untrack() {
      onUntrack?.();
    },
    get isNullObject() {
      need("isNullObject");
      return !store.has(key);
    },
    get value() {
      need("value");
      return store.get(key) ?? "";
    },
  };
}

export function makeShape(
  type: string,
  geo: string | undefined,
  box: { left: number; top: number; width: number; height: number },
) {
  const tagStore = new Map<string, string>();
  let ownId = `shape-${++idSeq}`;
  let ownLeft = box.left;
  let ownTop = box.top;
  /** Whether an answered `load()` has made this shape's properties readable. */
  let loadedProps = false;
  const needLoaded = (prop: string) => {
    if (!faults.strictShapeReads || loadedProps) return;
    throw new Error(
      `The property '${prop}' is not available. Before reading the property's value, call the load ` +
        'method on the containing object and call "context.sync()" on the associated request context. ' +
        `| code=PropertyNotLoaded | errorLocation=Shape.${prop}`,
    );
  };
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
    load() {
      // Queued, not granted. A load only takes effect on the sync that carries
      // it, and a sync that rejects carries nothing — same contract as
      // `pendingTagLoads`, and the reason `faults.strictShapeReads` can tell
      // "loaded" from "asked for".
      pendingShapeLoads.push(() => {
        loadedProps = true;
      });
    },
    get id() {
      needLoaded("id");
      return ownId;
    },
    set id(v: string) {
      ownId = v;
    },
    get left() {
      needLoaded("left");
      return ownLeft;
    },
    set left(v: number) {
      ownLeft = v;
    },
    get top() {
      needLoaded("top");
      return ownTop;
    },
    set top(v: number) {
      ownTop = v;
    },
    width: box.width,
    height: box.height,
    tagStore,
    get tags() {
      // See faults.tagsUndefinedOn — a host that hands back a shape with no
      // tags collection at all.
      if (faults.tagsUndefinedOn > 0) {
        faults.tagsUndefinedOn--;
        return undefined as unknown as FakeShape["tagsImpl"];
      }
      return shape.tagsImpl;
    },
    tagsImpl: {
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
      getItemOrNullObject: (k: string) =>
        makeTagProxy(tagStore, k, () => {
          untracked.tags++;
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

/**
 * Wrap a `getItemOrNullObject` result the way Office.js hands one over.
 *
 * `isNullObject` is not a property you can just read: it is populated on the
 * sync the proxy participated in, and a proxy nobody called `load()` on never
 * participates. Reading it then throws PropertyNotLoaded. The fake answered it
 * unconditionally, so a whole class of bug — resolve a shape, never load it,
 * read `isNullObject` — was invisible here and fatal on a real host.
 *
 * The load must name a REAL property. `load("isNullObject")` looks like the
 * obvious way to ask and is not one: `isNullObject` is not a property the host
 * holds, it is a flag Office.js sets from the response to a load of real
 * properties. Selecting it alone asks the host for nothing, the proxy takes no
 * part in the sync, and the flag stays unreadable — `PropertyNotLoaded`, with
 * no `errorLocation`, because the getter is on Office.js's base class.
 *
 * The fake accepted it, so five resolves in `powerpoint.ts` were written that
 * way and every test agreed with them. On PowerPoint on the web the one on the
 * in-place update path meant editing a chart threw before it did anything —
 * "edit a chart on the visible slide", the first self-test scenario to fail on
 * a real host.
 */
function nullObjectProxy<T extends object>(found: T | undefined) {
  let loaded = false;
  const base = found ?? ({ delete() {} } as unknown as T);
  // Read once, off the raw object, before any strictness gate can refuse it —
  // the fault is addressed by id and must not depend on the fault it models.
  const ownId = (() => {
    try {
      return (found as { id?: string } | undefined)?.id;
    } catch {
      return undefined;
    }
  })();
  return new Proxy(base, {
    get(target, prop, recv) {
      if (prop === "load")
        return (p?: string | string[]) => {
          const asked = (Array.isArray(p) ? p : (p ?? "").split(",")).map((s) => s.trim()).filter(Boolean);
          if (asked.length && asked.every((a) => a === "isNullObject")) return;
          // The host took the load and answered nothing — see
          // `unansweredNullChecks`. One-shot, so a retry can find it.
          if (ownId && unansweredNullChecks.has(ownId)) {
            unansweredNullChecks.delete(ownId);
            return;
          }
          loaded = true;
        };
      if (prop === "isNullObject") {
        if (!loaded)
          throw new Error(
            "The property 'isNullObject' is not available. Before reading the property's value, call the load " +
              'method on the containing object and call "context.sync()" on the associated request context.',
          );
        return !found;
      }
      return Reflect.get(target, prop, recv);
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
      // Taken and not performed — see faults.refuseSlideDelete.
      if (faults.refuseSlideDelete) return;
      deckRemove?.(slide);
    },
    /**
     * PowerPointApi 1.8's single-slide export.
     *
     * Returns a real .pptx (base64) carrying the deck's `<p:sldSz>` — the part
     * `slideSize`'s middle rung parses. A ClientResult, so `.value` is empty
     * until the sync it was queued in lands, exactly like the real one.
     */
    exportAsBase64() {
      const result = { value: "" };
      pendingExports.push({ result, build: exportedDeckBase64 });
      return result;
    },
    /**
     * PowerPointApi 1.5's shape selection — on **Slide**, which is the point.
     *
     * `Presentation` has `getSelectedShapes` and `setSelectedSlides` but no
     * `setSelectedShapes`; the shape half is one class down. Believing
     * otherwise is why the in-host battery went months without covering the
     * pane's selection-driven paths at all.
     *
     * Writes through to the same array `getSelectedShapes()` reads, so a
     * select-then-read round trip means something here.
     */
    setSelectedShapes(ids: string[]) {
      noteSelectionCall();
      // Taken, and poisonous. The select itself still happens below — the web
      // host does perform it — and only what comes AFTER stops answering.
      if (faults.selectionWedgesHost && ids.length) selectionWedged = true;
      if (!ids.length) {
        // Cleared on desktop, IGNORED on PowerPoint on the web
        // (office-js#3083). Modelled as the web does it when the fault is
        // armed, because a scenario that relies on the clear working is one
        // that will pass here and strand the next scenario in the field.
        if (!faults.webIgnoresDeselect) selectionRef.length = 0;
        return;
      }
      const found = faults.selectionIgnoresIds
        ? created.filter((s) => !s.deleted).slice(0, 1)
        : created.filter((s) => !s.deleted && ids.includes(s.id));
      selectionRef.length = 0;
      selectionRef.push(...found);
    },
    /**
     * PowerPointApi 1.8's slide raster. A ClientResult like `exportAsBase64`.
     *
     * The fake cannot draw, so it answers with a PNG whose payload encodes what
     * is ON the slide — live shape count and total ink area. That is enough for
     * the only assertion worth making here: an EMPTY slide and a slide with a
     * chart on it must not produce the same image. A fake that returned a fixed
     * string would let a scenario claiming "the chart is visible" pass over a
     * blank slide, which is precisely the failure the scenario exists to catch.
     */
    getImageAsBase64(_options?: { width?: number }) {
      const result = { value: "" };
      pendingExports.push({
        result,
        build: async () => {
          if (faults.constantSlideImage) return btoa("PNG:blank");
          const live = created.filter((s) => !s.deleted);
          const ink = live.reduce((n, s) => n + Math.max(0, s.width) * Math.max(0, s.height), 0);
          return btoa(`PNG:${slide.id}:shapes=${live.length}:ink=${ink}`);
        },
      });
      return result;
    },
    tags: {
      add: (k: string, v: string) => void slideTagStore.set(k, v),
      getItemOrNullObject: (k: string) => makeTagProxy(slideTagStore, k),
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
        // The host never answered the load queued for this collection — see
        // `unansweredShapeReads`. Nothing to hand back, and nothing to say
        // about what is on the slide.
        if (unansweredShapeReads.has(id)) {
          unansweredShapeReads.delete(id);
          throw new Error(
            "The property 'items' is not available. Before reading the property's value, call the load " +
              'method on the containing object and call "context.sync()" on the associated request context. ' +
              "| code=PropertyNotLoaded | errorLocation=ShapeCollection.items",
          );
        }
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
        if (faults.faultShapeCollectionLoad) throw new TypeError("e.load is not a function");
        if (p) lastShapeLoad = p;
        // `items/id,items/left,items/top` loads the SHAPES, not just the
        // collection — that is what the selector means, and it is how the real
        // host populates them. Without this the fake would be stricter than
        // Office.js rather than equal to it, and `strictShapeReads` would fail
        // correct code, which is the mirror image of the sin it exists to
        // prevent.
        if (p?.includes("items/")) for (const s of created) s.load();
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
        //
        // And `isNullObject` is only READABLE once the proxy has been loaded.
        // This fake used to hand it over unconditionally, which is a fiction:
        // Office.js populates it on the sync the proxy took part in, and a
        // proxy nobody loaded takes part in nothing. Reading it then throws
        // PropertyNotLoaded — which is exactly what editing an ungrouped chart
        // did on a real host while every test here passed, because the fake
        // answered a question the host refuses. Model the refusal.
        const found = created.find((s) => s.id === id && !s.deleted);
        return nullObjectProxy(found);
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

/**
 * Slide ids whose next shape-collection read finds the load UNANSWERED.
 *
 * One step past `hollowReads`, and observed as its limit: the web host answers
 * a collection read short, and sometimes not at all. Office.js then leaves the
 * collection unpopulated, and `.items` throws `PropertyNotLoaded` at
 * `ShapeCollection.items` rather than coming back empty — which is a very
 * different thing for a caller to be told, and the reason a deck-wide chart
 * scan used to die on one silent slide out of thirty-eight.
 *
 * Per slide id, and consumed on use, so a test can silence exactly one slide
 * without touching the fake's own internal reads of the same collection.
 */
export const unansweredShapeReads = new Set<string>();

/**
 * Shape/slide ids whose `getItemOrNullObject` proxy is never populated, however
 * it is loaded — so `isNullObject` stays unreadable.
 *
 * The single-object twin of `unansweredShapeReads`, and the state a caller has
 * the hardest time reasoning about: the object was asked for, the sync landed,
 * and the host said nothing either way. It is neither "there" nor "gone", and
 * code that treats it as one of the two either deletes something it cannot see
 * or refuses to touch something that is plainly there. Only the first of those
 * is unrecoverable, which is why `isLive` answers false.
 */
export const unansweredNullChecks = new Set<string>();

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
        //
        // Through `nullObjectProxy` for the same reason the shape collection
        // goes through it: `isNullObject` is populated by the sync a `load()`
        // enrolled the proxy in, so reading it off an unloaded proxy throws.
        // Answering it unconditionally is what let the tag version of this
        // mistake ship — every caller here happens to load first today, and
        // this is what keeps that true.
        getItemOrNullObject: (id: string) => nullObjectProxy(slides.find((s) => s.id === id)),
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
      /**
       * PowerPointApi 1.10's direct slide-size read.
       *
       * Load-gated like everything else here: `slideWidth` off an unloaded
       * PageSetup is PropertyNotLoaded on a real host, and a fake that answered
       * anyway would accept a `slideSize()` that forgot to load it — which is
       * the one mistake this whole class of proxy keeps producing.
       */
      pageSetup: (() => {
        let readable = false;
        const need = (prop: string) => {
          if (!readable) throw new Error(`PropertyNotLoaded: The property '${prop}' is not available.`);
        };
        return {
          load(_p?: string) {
            pendingTagLoads.push(() => {
              readable = true;
            });
          },
          get slideWidth() {
            need("slideWidth");
            return hostSlideSize.cx / 12700;
          },
          get slideHeight() {
            need("slideHeight");
            return hostSlideSize.cy / 12700;
          },
        };
      })(),
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
      getSelectedShapes: () => ({
        // The LIVE selection, not the array installHost was handed — otherwise
        // `setSelectedShapes` would write somewhere nothing reads.
        get items() {
          return selectionRef;
        },
        // Same contract as a slide's own collection: an `items/…` selector
        // populates the shapes it names. See ShapeCollection.load.
        load(p?: string) {
          noteSelectionCall();
          if (p?.includes("items/")) for (const s of selectionRef) s.load();
        },
      }),
      setSelectedSlides: (ids: string[]) => {
        noteSelectionCall();
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
        // Queued tag loads become readable here and nowhere else — see
        // makeTagProxy. A load only takes effect when its sync lands.
        for (const apply of pendingTagLoads) apply();
        pendingTagLoads.length = 0;
        for (const apply of pendingShapeLoads) apply();
        pendingShapeLoads.length = 0;
      };
      // Exports resolve on the sync that asked for them, same as a real
      // ClientResult. Awaited before `commit()` so the value is there the
      // moment the caller's `await context.sync()` returns.
      if (pendingExports.length) {
        for (const e of pendingExports) e.result.value = await e.build();
        pendingExports.length = 0;
      }
      const discard = () => {
        for (const s of slides) {
          for (const p of s.pending) {
            const i = s.created.indexOf(p);
            if (i >= 0) s.created.splice(i, 1);
          }
          s.pending.length = 0;
        }
        // A sync that rejects populates nothing: the loads it carried are lost,
        // and reading those proxies still throws.
        pendingTagLoads.length = 0;
        pendingShapeLoads.length = 0;
      };
      if (trips.syncs === faults.failSyncOn || failSyncsOn.has(trips.syncs)) {
        discard();
        throw new Error("host refused a queued command");
      }
      // A sync carrying a selection call on a wedged host. Never settles —
      // neither resolve nor reject, which is what the web host did and is why
      // the wait had to be bounded rather than caught. The caller's
      // `withTimeout` is the only thing that ends it.
      if (wedgeThisSync) {
        wedgeThisSync = false;
        await new Promise(() => {});
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
  // Proxies from a previous test's host must not become readable inside this
  // one's first sync.
  pendingTagLoads.length = 0;
  pendingShapeLoads.length = 0;
  pendingExports.length = 0;
  // Back to 16:9 unless a test says otherwise — a leaked 4:3 would silently
  // change placement for every test after it.
  hostSlideSize.cx = 12192000;
  hostSlideSize.cy = 6858000;
  pendingHostError = null;
  // failSyncOn was the one fault installHost did not reset, so a test that set
  // it leaked into every later test in the file and every test in every file
  // that ran after. It only ever looked safe because each user reset it by
  // hand — a convention, one `return` away from being forgotten.
  faults.failSyncOn = 0;
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
  faults.faultShapeCollectionLoad = false;
  faults.strictShapeReads = false;
  faults.unansweredTagLoads = 0;
  faults.webIgnoresDeselect = false;
  faults.selectionIgnoresIds = false;
  faults.selectionWedgesHost = false;
  selectionWedged = false;
  wedgeThisSync = false;
  faults.constantSlideImage = false;
  faults.refuseSlideDelete = false;
  faults.tagsUndefinedOn = 0;
  // The live shape selection starts as installHost was told, and is mutated
  // from there by Slide.setSelectedShapes.
  selectionRef.length = 0;
  selectionRef.push(...selectedShapes);
  unansweredShapeReads.clear();
  unansweredNullChecks.clear();
  faults.swallowDecks = 0;
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
