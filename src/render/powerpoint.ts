/**
 * Office.js renderer: draws a scene as native, individually editable
 * PowerPoint shapes on the current slide, then groups them.
 *
 * This is the same output strategy as think-cell/UpSlide-style tools —
 * charts stay fully editable in PowerPoint (every bar and label is a shape),
 * rather than being pasted as pictures or opaque OLE charts.
 *
 * Requires PowerPointApi 1.4+ (ShapeCollection.addGeometricShape / addLine /
 * addTextBox) — marker symbols are preset geometry, so they need only 1.4 too.
 * Grouping (1.8+) and shape rotation (1.10+) degrade gracefully on older hosts.
 */
import { polar, arrowheadBox, wedgeFanSteps, wedgeFanChord, SYMBOL_PRESET, dashKind } from "../core/geometry";
import { estimateOfficeShapes } from "../core/scene";
import { toHex6, alphaOf } from "../core/color";
import type { PolygonNode, Scene, SceneNode, TextNode, WedgeNode } from "../core/scene";

/* global PowerPoint, Office */

export interface InsertOptions {
  /** Top-left of the chart frame on the slide, in points. */
  left?: number;
  top?: number;
  /** Group the shapes after insertion (default true). */
  group?: boolean;
  fontFamily?: string;
  /**
   * Serialized chart model stored as a tag on the inserted group
   * (PowerPointApi 1.3+), so a future version can re-open and re-edit
   * the chart — the think-cell "live chart" pattern.
   */
  tagData?: string;
  /**
   * Accessible description set on the chart group (Shape.altTextDescription), so
   * a screen reader announces the chart in PowerPoint instead of reading a pile
   * of unnamed shapes. Comes from the scene's `desc` (see describeChart).
   */
  altText?: string;
  /** Accessible title (Shape.altTextTitle); comes from the chart title. */
  altTitle?: string;
  /**
   * Batch size for renderShapesChunked. Defaults to SHAPES_PER_SYNC (10),
   * calibrated for the LIVE canvas — PowerPoint web repaints as shapes arrive
   * and stops answering past roughly twenty in one batch. Off-screen slides
   * (demo, agenda) don't repaint mid-render and swallow far larger batches,
   * so the demo/agenda paths pass a larger value to cut ~4-7 syncs per chart.
   */
  shapesPerSync?: number;
  /**
   * A PNG of the WHOLE scene, base64. Accepts a bare payload, a `data:` URI, or
   * the prefix-without-scheme form the skill's pptxgen path emits
   * (`image/png;base64,…`) — `barePng` normalises all three, because Office.js
   * wants the bare payload ("A string that is a Base64 encoding of the image
   * data", @types/office-js).
   *
   * When set, AND the host advertises PowerPointApi 1.8 (`ShapeFill.setImage`),
   * AND the payload is under MAX_PICTURE_BASE64, the chart is inserted as ONE
   * picture shape instead of the scene's nodes. Otherwise it is ignored and the
   * native-shape path runs unchanged — see `wantsPicture`.
   *
   * The raster MUST be of the same scene that sizes the box: the rect is
   * `scene.width` x `scene.height` points, so a caller that resizes between
   * rasterising and inserting distorts the chart. Never stored in ChartConfig
   * and therefore never in a shape tag — it is derived from the config at
   * insert time and thrown away.
   */
  pictureBase64?: string;
}

/** Tag key under which the chart's serialized config is persisted. */
export const CHART_TAG = "POWERCHART_CONFIG";

/**
 * Tag key under which an UNGROUPED chart records its other shapes' ids.
 *
 * A grouped chart is one shape, so deleting the tagged shape deletes the chart.
 * Without grouping (PowerPoint on the web, or an addGroup the host refuses) the
 * config tag can only sit on ONE of the chart's shapes — and an in-place update
 * that deletes just that one leaves the other twelve on the slide, underneath
 * the redraw, so the slide gains a whole chart's worth of orphans on every edit.
 * The tagged shape therefore also carries its siblings' ids: the group the host
 * would not make, written down. Absent on a grouped chart, and absent on charts
 * inserted before this existed — both fall back to deleting the tagged shape.
 */
export const CHART_PARTS_TAG = "POWERCHART_PARTS";

/**
 * Tag key under which a chart records the FRAME ORIGIN it was drawn at.
 *
 * An in-place update re-renders the scene at an origin, and the only origin it
 * had was the tagged shape's `left`/`top` — which is not the frame origin. For a
 * grouped chart that is the group's bounding box, i.e. the frame origin plus the
 * scene's ink offset; ungrouped it is whatever `created[0]` happens to be. Either
 * way, feeding it back as the frame origin shifts the chart by that offset, and
 * the shift COMPOUNDS: every Same Scale run or select-and-update cycle moved the
 * chart again (measured: +8pt per cycle grouped, +285pt ungrouped — off the slide
 * in one run on the web host).
 *
 * Writing the origin down makes the update idempotent: re-rendering lands the
 * chart exactly where it already was. Absent on charts inserted before this
 * existed, which fall back to the old shape-position behaviour.
 */
export const CHART_ORIGIN_TAG = "POWERCHART_ORIGIN";

/**
 * Slide-level tag written on every demo-deck slide at creation time — a JSON
 * envelope carrying the item's title so a blank readback can NAME the missing
 * chart instead of just its 1-based deck position. Requires PowerPointApi 1.3;
 * a host without slide tags silently skips it and blanks stay position-only.
 * Not read by the update path — this exists purely for regression self-check.
 */
export const DEMO_SLOT_TAG = "POWERCHART_DEMO_SLOT";

/**
 * Release a proxy object's client-side memory once its loaded values have been
 * read. Office.js keeps every proxy touched in a `run` alive until the context
 * is disposed, and the docs call out a "noticeable performance benefit when
 * using large numbers of proxy objects" from untracking — a deck-wide scan
 * creates one shape proxy and one tag proxy PER shape ON EVERY SLIDE, which is
 * exactly that case. Best-effort: a null-object proxy may not expose untrack.
 */
function untrack(obj: unknown): void {
  try {
    (obj as { untrack?: () => void })?.untrack?.();
  } catch {
    /* not a tracked proxy on this host */
  }
}

const DEFAULT_FONT = "Segoe UI";

/** Where an existing PowerChart lives on the deck, for in-place update. */
export interface EditTarget {
  slideId: string;
  shapeId: string;
  left: number;
  top: number;
  /**
   * The rest of the chart's shapes when it is UNGROUPED — read back from the
   * parts tag, so an in-place update deletes the whole chart. See
   * CHART_PARTS_TAG.
   */
  partIds?: string[];
  /**
   * The frame origin the chart was drawn at, from CHART_ORIGIN_TAG. Present on
   * charts inserted since that tag existed; an in-place update re-renders here
   * so it lands where it already is instead of walking across the slide.
   */
  origin?: { left: number; top: number; anchorLeft: number; anchorTop: number };
}

/**
 * Where an insert has got to. A host that stops answering does not throw — the
 * sync promise simply never settles — so without this a stall is
 * indistinguishable from slow work, and there is nothing to report but a
 * spinner. Every phase is named so the pane can say which one it died in.
 */
export type InsertPhase = "context" | "queue" | "commit" | "group" | "done";

/**
 * Reported when a timed-out call finally settles — see `withTimeout`.
 * `null` while nothing has been abandoned.
 */
export let lastLateSync: string | null = null;

/**
 * Monotonically incremented each time `describe` fires — an abandoned sync
 * reported success or a RichApi error. Callers snapshot it before an item and
 * compare after, so an identical `lastLateSync` string across two consecutive
 * items (same `what` + same rounded seconds) doesn't read as "no late-sync
 * happened". Two identical stalls back-to-back would silently miss lateFired
 * otherwise — a real ordering bug the pane's test fleet surfaced.
 */
export let lastLateSyncSeq = 0;

let lateSubscriber: ((msg: string) => void) | null = null;

/** Be told when a call we already gave up on finally settles. */
export function onLateSync(cb: (msg: string) => void): void {
  lateSubscriber = cb;
}

/**
 * The most recent abandoned sync's settle promise — resolves when the sync
 * eventually reports success OR a RichApi error. Callers that catch a timeout
 * can `await waitForLateSync()` before deciding what to do, so `lastLateSync`
 * carries the host's actual verdict rather than a stale null.
 */
let pendingLateSyncSettle: Promise<void> | null = null;

/**
 * Wait up to `maxMs` for the last abandoned sync (see `withTimeout`) to report
 * its outcome, so `lastLateSync` reflects reality before the caller reads it.
 * No-op if nothing is outstanding. Returns whether the settle actually landed.
 */
export async function waitForLateSync(maxMs = 5_000): Promise<boolean> {
  if (!pendingLateSyncSettle) return false;
  const p = pendingLateSyncSettle;
  let settled = false;
  await Promise.race([
    p.then(() => {
      settled = true;
    }),
    new Promise<void>((r) => setTimeout(r, maxMs)),
  ]);
  return settled;
}

/**
 * Reject if `p` has not settled within `ms` — a hung host must not hang the
 * pane.
 *
 * Racing alone throws away the answer, and the answer is the whole point: the
 * abandoned promise keeps running, and whatever it does NEXT is the only
 * evidence we get about a host that went quiet. If it resolves at 45s the host
 * is merely slow and the timeout is wrong; if it rejects with a RichApi error,
 * that error names the real bug and would otherwise be lost forever, because
 * Office.js reports queued-command failures HERE and nowhere else. So keep
 * listening after giving up, and record what arrives.
 *
 * On timeout, `pendingLateSyncSettle` tracks the abandoned promise's eventual
 * outcome so a caller can `await waitForLateSync()` before reading
 * `lastLateSync` — the difference between "sync at 60s, chart on the slide"
 * (rendered late) and "sync never returned" (a real host death).
 */
function withTimeout<T>(p: Promise<T>, ms: number, what: string): Promise<T> {
  const started = Date.now();
  const describe = (outcome: string): void => {
    lastLateSync = `${what}: ${outcome} after ${Math.round((Date.now() - started) / 1000)}s`;
    lastLateSyncSeq += 1;
    lateSubscriber?.(lastLateSync);
  };
  return new Promise<T>((resolve, reject) => {
    let done = false;
    p.then(
      (v) => {
        if (!done) {
          done = true;
          clearTimeout(timer);
          resolve(v);
        } else {
          describe("the host eventually SUCCEEDED");
        }
      },
      (err: unknown) => {
        if (!done) {
          done = true;
          clearTimeout(timer);
          reject(err);
        } else {
          describe(`the host eventually FAILED — ${errorText(err)}`);
        }
      },
    );
    const timer = setTimeout(() => {
      if (done) return;
      done = true;
      // Track the abandoned promise's eventual settle so waitForLateSync can
      // await it before the caller reads lastLateSync.
      pendingLateSyncSettle = p.then(
        () => undefined,
        () => undefined,
      );
      reject(new Error(`PowerPoint did not respond while ${what} (${ms / 1000}s)`));
    }, ms);
  });
}

/**
 * Everything an Office.js error knows. A RichApi.Error's `message` is usually
 * generic ("An internal error has occurred"); the useful part — the failing
 * command and why — lives in `code` and `debugInfo`, which a plain String(err)
 * silently drops.
 */
export function errorText(err: unknown): string {
  if (!err || typeof err !== "object") return String(err);
  const e = err as { message?: string; code?: string; debugInfo?: unknown };
  const bits = [e.message ?? String(err)];
  if (e.code) bits.push(`code=${e.code}`);
  if (e.debugInfo) {
    try {
      bits.push(`debugInfo=${JSON.stringify(e.debugInfo)}`);
    } catch {
      /* not serialisable — the message and code still carry */
    }
  }
  return bits.join(" | ");
}

/**
 * How long one batch may take before we call the host stalled.
 *
 * A constant, because the batch is a constant: SHAPES_PER_SYNC caps what we
 * hand over at a size the live canvas is known to swallow, so there is nothing
 * left for the budget to scale WITH. (This was briefly a function of the shape
 * count — a flat 20s had killed a 40-shape chart at almost exactly the moment
 * it would have finished. Chunking made that whole question moot: the fix was
 * never a bigger number, it was a smaller batch.)
 *
 * Generous, because its only job is to stop an infinite spinner. Being late
 * costs a user nothing; being wrong costs them their chart.
 */
let BATCH_TIMEOUT_MS = 45_000;

/**
 * Test-only: shorten the batch timeout so a stalled-sync scenario is testable
 * inside a normal vitest run. Production callers never touch it; the default is
 * the 45s ceiling above. Restore by passing 45_000 back in.
 */
export function _setBatchTimeoutForTest(ms: number): void {
  BATCH_TIMEOUT_MS = ms;
}

/**
 * The blank layout of the presentation's first master, or undefined if the host
 * has no opinion.
 *
 * A slide added with no layout inherits the previous slide's — which on a fresh
 * deck is the title slide, so an agenda lands on top of "Click to add title"
 * with the placeholder showing through. We draw every element ourselves and
 * want no placeholders at all. Matched on `type`, not on the name: the name is
 * localised ("Tom" on a Danish master) and matching English would silently do
 * nothing for most of the world.
 */
async function blankLayoutId(context: PowerPoint.RequestContext): Promise<string | undefined> {
  try {
    const masters = context.presentation.slideMasters;
    masters.load("items/id,items/layouts/items/id,items/layouts/items/type");
    await context.sync();
    for (const master of masters.items) {
      const blank = master.layouts.items.find((l) => l.type === PowerPoint.SlideLayoutType.blank);
      if (blank) return blank.id;
    }
  } catch {
    /* no master/layout access on this host — fall back to the inherited layout */
  }
  return undefined;
}

export async function insertSceneIntoSlide(
  scene: Scene,
  opts: InsertOptions = {},
  onPhase?: (phase: InsertPhase, detail?: string) => void,
): Promise<void> {
  onPhase?.("context");
  await PowerPoint.run(async (context) => {
    // The current slide already exists, so its proxy IS stable across syncs (its
    // id round-trips) — hold one and reuse it. Only a freshly-added slide needs a
    // per-batch fresh proxy; see SlideThunk. Resolving once also pins the target
    // to the slide selected at the start, immune to any selection drift mid-draw.
    const slide = getTargetSlide(context);
    const getSlide: SlideThunk = () => slide;
    onPhase?.("queue", `${scene.nodes.length} nodes`);
    // Committed in batches: the whole scene in one sync is what a live canvas
    // will not take. Each batch reports, so progress here is measured, not
    // guessed — see renderShapesChunked.
    const created = await renderShapesChunked(context, getSlide, scene, opts, (done, total) =>
      onPhase?.("commit", `${done} of ${total} shapes`),
    );
    // Shapes are committed by now, so grouping/tagging (which some hosts,
    // notably PowerPoint on the web, don't support) can't roll back the chart.
    onPhase?.("group");
    await groupAndTagAll(context, [
      { getSlide, created, opts: { ...opts, altText: scene.desc, altTitle: scene.title } },
    ]);
    onPhase?.("done");
  });
}

/** Replace an existing PowerChart group with a re-rendered scene, in place. */
export async function updateChartInSlide(
  scene: Scene,
  target: EditTarget,
  opts: InsertOptions = {},
): Promise<EditTarget | null> {
  const [next] = await updateChartsInSlides([{ scene, target, opts }]);
  return next ?? null;
}

/**
 * Replace any number of existing PowerCharts in place, in ONE request context.
 *
 * Every Office.js sync is a round-trip to PowerPoint, so the thing that must not
 * scale with the chart count is the number of syncs — not the number of shapes,
 * which ride along in a batch for free. Re-rendering charts one at a time (a
 * loop around the single-chart update) cost 4 syncs and a whole PowerPoint.run
 * context EACH: Same Scale across a 20-chart deck was 80 round-trips. This is
 * four, whatever N is.
 *
 * The four phases are ordered, and that order is load-bearing: every old shape
 * resolves before any is deleted, and every new shape COMMITS before anything is
 * grouped — so a host without grouping cannot roll back the charts themselves.
 * Batching happens across charts WITHIN a phase, never across phases.
 */
export async function updateChartsInSlides(
  items: { scene: Scene; target: EditTarget; opts?: InsertOptions }[],
): Promise<EditTarget[]> {
  if (!items.length) return [];
  return PowerPoint.run(async (context) => {
    // 1. Resolve every old shape — one sync for all of them.
    //
    // getItemOrNullObject, never getItem: a target names a slide that the user
    // may have deleted, undone, or closed since we read it, and getItem THROWS
    // on a stale id — "InvalidParam passed to GetItem(id)", code 5010, which is
    // a normal condition wearing a crash's clothes. A chart whose slide is gone
    // is not an error, it is nothing to do.
    const found = items.map((it) => {
      const slide = context.presentation.slides.getItemOrNullObject(it.target.slideId);
      slide.load("isNullObject");
      return { it, slide };
    });
    await context.sync();

    const live = found.filter(({ slide }) => !slide.isNullObject);
    if (!live.length) return [];
    const withOld = live.map(({ it, slide }) => {
      const old = slide.shapes.getItemOrNullObject(it.target.shapeId);
      // The shape's LIVE position, queued in the sync this resolution already
      // costs. The caller's EditTarget is a snapshot — the task pane holds one
      // from whenever the chart was loaded — so measuring the user's drag against
      // it reports no movement for a chart that has since been dragged, and the
      // update puts it back where it was. Only the host knows where the shape is
      // now.
      old.load("left,top");
      return {
        it,
        slide,
        old,
        // An ungrouped chart is more than its tagged shape (see CHART_PARTS_TAG).
        // Resolved in this same sync, so the delete below already knows which of
        // them the user has since removed by hand.
        parts: (it.target.partIds ?? []).map((id) => slide.shapes.getItemOrNullObject(id)),
      };
    });
    await context.sync();

    // A target whose SHAPE is gone gets the same treatment as one whose slide
    // is gone: nothing to do. Re-rendering it would resurrect a chart the user
    // deleted — an in-place update that inserts is not an update.
    // Read the live positions off the proxies BEFORE the delete below detaches
    // them; from here on `at` is where each chart actually sits on the slide.
    const alive = withOld
      .filter(({ old }) => !old.isNullObject)
      .map((e) => ({ ...e, at: { left: e.old.left, top: e.old.top } }));
    if (!alive.length) return [];

    // 2. Drop the old shapes — one sync for all of them, siblings included:
    //    deleting only the tagged shape of an ungrouped chart leaves the rest
    //    of it on the slide, under the redraw.
    for (const { old, parts } of alive) {
      old.delete();
      for (const p of parts) if (!p.isNullObject) p.delete();
    }
    await context.sync();

    // 3. Redraw each chart in batches. One of these charts is on the slide the
    //    user is looking at, and a live canvas will not take a whole chart in
    //    one sync — so the batching is not an optimisation here, it is the only
    //    way the shapes arrive at all. Per chart, because a chart's shapes must
    //    all reach the same slide.
    const rendered: Grouping[] = [];
    for (const { it, slide, at } of alive) {
      const opts: InsertOptions = {
        ...it.opts,
        // The recorded frame origin, shifted by however far the user has dragged
        // the chart since it was tagged (livePos - anchor). Untouched, that delta
        // is zero and the chart re-renders exactly where it is; dragged, it
        // follows the drag. Measured against the LIVE shape (`at`), never the
        // caller's snapshot — the pane holds one EditTarget from when the chart
        // was loaded and reuses it for every update, so a snapshot-based delta
        // reports no movement and puts a dragged chart back. Charts with no
        // usable origin tag fall back to the shape's own position.
        left: it.target.origin ? it.target.origin.left + (at.left - it.target.origin.anchorLeft) : at.left,
        top: it.target.origin ? it.target.origin.top + (at.top - it.target.origin.anchorTop) : at.top,
        altText: it.scene.desc,
        altTitle: it.scene.title,
      };
      // An existing slide's proxy is stable across syncs — hold it. Only a
      // freshly-added slide needs a per-batch fresh proxy; see SlideThunk.
      const getSlide: SlideThunk = () => slide;
      rendered.push({ getSlide, created: await renderShapesChunked(context, getSlide, it.scene, opts), opts });
    }

    // 4-5. Group, then tag — one sync each, however many charts.
    const tagged = await groupAndTagAll(context, rendered);

    // 6. Hand back the NEW targets. An update replaces every shape, so the
    //    caller's target is dead as soon as this returns: the pane used to keep
    //    the old one, and its next update resolved a shape id that no longer
    //    existed, was filtered out as "the user deleted this chart", and did
    //    nothing at all — silently. Auto-update died the same way after its
    //    first push. Returning the new target is what lets a caller stay live.
    return alive.map(({ it, at }, i) => {
      const t = tagged[i]?.target;
      if (!t) return it.target;
      // The origin this pass actually rendered at, paired with where the tagged
      // shape landed — the same (origin, anchor) contract groupAndTagAll wrote to
      // the tag, so the caller's in-memory target and the on-slide tag agree.
      const o = it.target.origin;
      return {
        slideId: it.target.slideId,
        shapeId: t.id,
        left: t.left,
        top: t.top,
        partIds: tagged[i]?.partIds,
        origin: {
          left: o ? o.left + (at.left - o.anchorLeft) : at.left,
          top: o ? o.top + (at.top - o.anchorTop) : at.top,
          anchorLeft: t.left,
          anchorTop: t.top,
        },
      };
    });
  });
}

/**
 * Read the PowerChart config back from the current selection (the tag written
 * at insert time). Returns null when the selection is not a PowerChart.
 * Requires PowerPointApi 1.5 (getSelectedShapes).
 */
export async function loadChartFromSelection(): Promise<{ configJson: string; target: EditTarget } | null> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    const slide = slides.getItemAt(0);
    slide.load("id");
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/left,items/top");
    await context.sync();

    const tags = shapes.items.map((s) => chartTagsOf(s));
    await context.sync();

    for (let i = 0; i < shapes.items.length; i++) {
      const { config, parts, origin } = tags[i];
      if (!config.isNullObject && config.value) {
        const s = shapes.items[i];
        return {
          configJson: config.value,
          target: {
            slideId: slide.id,
            shapeId: s.id,
            left: s.left,
            top: s.top,
            partIds: partIdsOf(parts),
            origin: originOf(origin),
          },
        };
      }
    }
    return null;
  });
}

/** Both PowerChart tags of one shape, queued for the next sync. */
interface ChartTags {
  config: PowerPoint.Tag;
  parts: PowerPoint.Tag;
  origin: PowerPoint.Tag;
}

/**
 * Queue a shape's PowerChart tags for reading: the config, and the sibling ids
 * an ungrouped chart carries alongside it. Both in the same sync — an update
 * needs the parts wherever it needs the config, and a second round-trip per
 * scan is exactly what the batching elsewhere in this file exists to avoid.
 */
function chartTagsOf(shape: PowerPoint.Shape): ChartTags {
  const config = shape.tags.getItemOrNullObject(CHART_TAG);
  config.load("value");
  const parts = shape.tags.getItemOrNullObject(CHART_PARTS_TAG);
  parts.load("value");
  const origin = shape.tags.getItemOrNullObject(CHART_ORIGIN_TAG);
  origin.load("value");
  return { config, parts, origin };
}

/**
 * The sibling shape ids a parts tag carries, or undefined for anything else.
 *
 * Parsed defensively although we wrote it: a shape tag is editable in the host,
 * survives a copy into another deck, and reaches this code from a file we did
 * not author — a malformed one must degrade to "no siblings", never throw
 * inside the update's first sync.
 */
function partIdsOf(parts: PowerPoint.Tag): string[] | undefined {
  if (parts.isNullObject || !parts.value) return undefined;
  try {
    const ids: unknown = JSON.parse(parts.value);
    if (!Array.isArray(ids)) return undefined;
    const ok = ids.filter((id): id is string => typeof id === "string" && id.length > 0);
    return ok.length ? ok : undefined;
  } catch {
    return undefined;
  }
}

/**
 * The frame origin a chart recorded, or undefined for anything else. Parsed as
 * defensively as partIdsOf: a shape tag is editable in the host and can arrive
 * from a deck we did not author, so a malformed one must degrade to "no recorded
 * origin" (falling back to the shape position), never throw inside a sync.
 */
function originOf(origin: PowerPoint.Tag): EditTarget["origin"] {
  if (origin.isNullObject || !origin.value) return undefined;
  try {
    const v: unknown = JSON.parse(origin.value);
    // [originLeft, originTop, anchorLeft, anchorTop]. All four or nothing: without
    // the anchor an update cannot tell "untouched" from "the user dragged it",
    // so a partial tag falls back to the shape's own position.
    if (!Array.isArray(v) || v.length < 4) return undefined;
    const [left, top, anchorLeft, anchorTop] = v;
    if (![left, top, anchorLeft, anchorTop].every((n) => typeof n === "number" && Number.isFinite(n))) return undefined;
    return { left, top, anchorLeft, anchorTop };
  } catch {
    return undefined;
  }
}

/**
 * Bounds of the currently selected shape when it is NOT a PowerChart —
 * used to insert a new chart into a selected placeholder/frame.
 */
export async function getSelectionBounds(): Promise<{
  left: number;
  top: number;
  width: number;
  height: number;
} | null> {
  try {
    return await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height");
      await context.sync();
      if (shapes.items.length !== 1) return null;
      const s = shapes.items[0];
      const tag = s.tags.getItemOrNullObject(CHART_TAG);
      tag.load("value");
      await context.sync();
      if (!tag.isNullObject && tag.value) return null; // it's a chart — edit, don't cover
      return { left: s.left, top: s.top, width: s.width, height: s.height };
    });
  } catch {
    return null;
  }
}

/** All PowerCharts in the current selection (for Same Scale on a subset). */
export async function listChartsInSelection(): Promise<{ configJson: string; target: EditTarget }[]> {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("id");
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/left,items/top");
    await context.sync();
    const tags = shapes.items.map((s) => chartTagsOf(s));
    await context.sync();
    const charts = shapes.items
      .map((s, i) => ({ s, ...tags[i] }))
      .filter(({ config }) => !config.isNullObject && config.value)
      .map(({ s, config, parts, origin }) => ({
        configJson: config.value,
        target: {
          slideId: slide.id,
          shapeId: s.id,
          left: s.left,
          top: s.top,
          partIds: partIdsOf(parts),
          origin: originOf(origin),
        },
      }));
    for (const t of tags) {
      untrack(t.config);
      untrack(t.parts);
      untrack(t.origin);
    }
    for (const s of shapes.items) untrack(s);
    return charts;
  });
}

/**
 * Find every PowerChart in the deck (any shape carrying the config tag),
 * across all slides. Used by "Same scale" to re-render charts together.
 */
export async function listChartsInDeck(): Promise<{ configJson: string; target: EditTarget }[]> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const perSlide = slides.items.map((slide) => {
      slide.shapes.load("items/id,items/left,items/top");
      return slide;
    });
    await context.sync();

    const lookups: ({ slideId: string; shape: PowerPoint.Shape } & ChartTags)[] = [];
    for (const slide of perSlide) {
      for (const shape of slide.shapes.items) {
        lookups.push({ slideId: slide.id, shape, ...chartTagsOf(shape) });
      }
    }
    await context.sync();

    const charts = lookups
      .filter((l) => !l.config.isNullObject && l.config.value)
      .map((l) => ({
        configJson: l.config.value,
        target: {
          slideId: l.slideId,
          shapeId: l.shape.id,
          left: l.shape.left,
          top: l.shape.top,
          origin: originOf(l.origin),
          partIds: partIdsOf(l.parts),
        },
      }));
    // Every value we need is now a plain string/number in `charts`; drop the
    // whole proxy sweep (one shape + its tags per shape, deck-wide) from memory.
    for (const l of lookups) {
      untrack(l.config);
      untrack(l.parts);
      untrack(l.shape);
    }
    return charts;
  });
}

/**
 * A slide reference that is RE-ACQUIRED on every call, never held.
 *
 * This is the crux of drawing onto a freshly-added slide. `slides.getItemAt(i)`
 * is a positional proxy, and the moment the host resolves one Office.js rewrites
 * its object path to `getItem(id)` (`createChildItemObjectPathUsingIndexerOr-
 * GetItemAt` / `fixObjectPathIfNecessary` in office-js). For a slide that was
 * just `add()`ed, PowerPoint on the web returns an id that does not round-trip
 * through `getItem`, so every *later* use of that same proxy throws "InvalidParam
 * passed to GetItem(id)" (code 5010). A *brand-new* `getItemAt(i)` proxy has not
 * been resolved, so it is still positional and executes cleanly — which is why
 * the fix is to call this thunk again for each sync-batch instead of holding one
 * proxy across them. Pre-existing slides never hit this (their id round-trips),
 * which is why inserting onto the current slide and editing in place always
 * worked.
 */
type SlideThunk = () => PowerPoint.Slide;

/**
 * How many `slides.add()` calls issued by `addSlides` did NOT survive to a
 * settled, fresh-context readback — even after the one retry round. This is a
 * DIFFERENT count from `DemoReport.addsIssued − slidesAdded`: that pair
 * measures overall deck-growth shortfall across a whole `insertDemoDeck` run
 * (a retry/fail elsewhere can cancel it out — see the comment on
 * `addsIssued`). `lastAddsLost` instead accumulates only the adds `addSlides`
 * itself confirmed missing at commit time, post-retry — the corruption the
 * fresh verify below exists to catch and recover from. Reset by
 * `insertDemoDeck` at the start of every run; surfaced via
 * `DemoReport.addsLostAtCommit` once the run finishes.
 */
let lastAddsLost = 0;

/** At most one retry round after the initial add — see `addSlides`. */
const MAX_ADD_RETRY_ROUNDS = 1;

/**
 * Append `count` blank slides and return a fresh-proxy thunk for each.
 *
 * By index, off a `getCount()` taken in its OWN sync BEFORE the adds —
 * `slides.add()` always appends to the end, so the new slides are
 * `start .. start+count-1`. The adds then get their own commit sync.
 *
 * The count as seen by THIS context right after that commit sync is not
 * trusted: `getCount()` queued in the SAME sync as the adds returns the
 * PRE-add total on PowerPoint web — the adds are queued but not yet reflected
 * in the count — so a delta assertion there throws "added 0 of N" while the
 * slides are in fact appearing. Instead, once the 2 syncs above land, open a
 * FRESH context (the same settled-read trick as `slideCount()`) and compare
 * its count against `start + count`. PowerPoint on the web silently drops some
 * `slides.add()` calls under load — observed in Presentation_3.pptx, where
 * the deck grew by 10 slides for 20 issued adds — and a fresh context is the
 * only reliable way to see that it happened.
 *
 * A deficit gets ONE retry round: issue the missing adds in another fresh
 * context and re-verify from a third. If the deficit persists (the drop was
 * not transient, or the retry itself got dropped), give up — log via
 * `console.warn`, add whatever is still missing to `lastAddsLost`, and return
 * thunks for only the slides that actually landed (fewer than `count`).
 *
 * Also NOT via `slides.items` (a snapshot, stale in the adds' sync — the bug that
 * returned zero new slides) and NOT by loading ids to re-acquire `getItem(id)`
 * (that id is the very thing the web host mis-round-trips).
 */
async function addSlides(
  context: PowerPoint.RequestContext,
  count: number,
  layoutId: string | undefined,
): Promise<SlideThunk[]> {
  if (count <= 0) return [];
  const slides = context.presentation.slides;
  const before = slides.getCount();
  await context.sync();
  const start = before.value;
  for (let i = 0; i < count; i++) slides.add(layoutId ? { layoutId } : undefined);
  await context.sync();

  let landed = await slideCount();
  let have = landed - start;
  let deficit = count - have;
  for (let round = 0; deficit > 0 && round < MAX_ADD_RETRY_ROUNDS; round++) {
    const toAdd = deficit;
    await PowerPoint.run(async (retryContext) => {
      const retrySlides = retryContext.presentation.slides;
      for (let i = 0; i < toAdd; i++) retrySlides.add(layoutId ? { layoutId } : undefined);
      await retryContext.sync();
    });
    landed = await slideCount();
    have = landed - start;
    deficit = count - have;
  }
  if (deficit > 0) {
    lastAddsLost += deficit;
    console.warn(
      `PowerChart: addSlides lost ${deficit} of ${count} requested slide${count === 1 ? "" : "s"} — ` +
        `the host dropped the add() and the retry did not recover it. Returning ${have} thunk${have === 1 ? "" : "s"}.`,
    );
  }
  const actual = Math.min(have, count);
  return Array.from({ length: actual }, (_, i) => () => context.presentation.slides.getItemAt(start + i));
}

/**
 * Append one agenda slide per chapter, each highlighting its own chapter
 * (think-cell's agenda). Slides land at the END of the deck: PowerPointApi's
 * slides.add has no insert-at-position (AddSlideOptions.index is preview-only),
 * and repositioning needs Slide.index (1.8) — so for now they stay appended, the
 * same as the demo deck. Requires PowerPointApi 1.3 (slides.add).
 */
export async function insertAgendaSlides(scenes: Scene[]): Promise<void> {
  await PowerPoint.run(async (context) => {
    const layoutId = await blankLayoutId(context);
    const slideThunks = await addSlides(context, scenes.length, layoutId);
    // Batched like every other render: a slide's worth of shapes in one sync is
    // what the host refuses. Off-screen slides tolerate more than the live
    // canvas does, but "more" is not a number worth betting on twice.
    for (let i = 0; i < scenes.length; i++) {
      await renderShapesChunked(context, slideThunks[i], scenes[i], {
        left: 0,
        top: 0,
        group: false,
        tagData: undefined,
        // Same off-screen slide, same larger batch as the demo deck.
        shapesPerSync: SHAPES_PER_SYNC_OFFSCREEN,
      });
    }
  });
}

/**
 * A chart above this many native shapes is not attempted on the demo deck: on
 * PowerPoint web it will not finish inside the batch timeout, and trying it both
 * wastes ~45s and loads the host toward the "we ran into a problem" crash. The
 * densest charts (a filled area is one line per edge) run 100-200 shapes; the
 * rest are well under. Tunable — the point is to skip the few that can't land,
 * not to trim the deck.
 */
const DEMO_SHAPE_BUDGET = 90;

/**
 * Draw a bold red banner ACROSS THE TOP of a slide so an incomplete one is
 * unmistakable — a half-rendered chart looks almost right, which is the trap.
 *
 * Deliberately a top strip, not a slab over the middle: a stamp that lands on a
 * real chart (a mis-targeted skip once landed on the butterfly) must not destroy
 * it, and a partial chart under a failed render should still be legible beneath
 * the banner. Best-effort styling: a host that lacks a property skips it, the
 * text still lands.
 */
async function stampSlide(
  context: PowerPoint.RequestContext,
  getSlide: SlideThunk,
  title: string,
  detail: string,
): Promise<void> {
  const box = (
    getSlide().shapes as unknown as {
      addTextBox(text: string, box: { left: number; top: number; width: number; height: number }): PowerPoint.Shape;
    }
  ).addTextBox(`${title} — ${detail}`, { left: 24, top: 12, width: 912, height: 46 });
  box.name = "PowerChart:not-complete";
  try {
    box.fill.setSolidColor("#c0392b");
    const font = (box.textFrame.textRange as unknown as { font: Record<string, unknown> }).font;
    font.color = "#ffffff";
    font.bold = true;
    font.size = 18;
    const para = (box.textFrame.textRange as unknown as { paragraphFormat: Record<string, unknown> }).paragraphFormat;
    para.horizontalAlignment = "Center";
  } catch {
    /* a styling property the host lacks — the banner text is what matters */
  }
  await context.sync();
}

/** Stamp the LAST slide from a FRESH context — used after a render poisoned its own. */
async function stampLastSlide(title: string, detail: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    const count = context.presentation.slides.getCount();
    await context.sync();
    if (count.value < 1) return;
    await stampSlide(context, () => context.presentation.slides.getItemAt(count.value - 1), title, detail);
  });
}

/**
 * Testing aid: append one slide per item and render its chart, tagged so each
 * stays re-editable. Returns the indices of any items that were NOT drawn as a
 * real chart — skipped as too dense, or failed mid-render — the caller names
 * them so the user knows what to retry on its own. Every such slide is left with
 * a "NOT COMPLETE" stamp so a placeholder is never mistaken for a real chart.
 *
 * ONE `PowerPoint.run` per slide, for two reasons learned on the real host:
 * - Isolation. A chart the host cannot finish (a dense area chart is ~200 native
 *   shapes, since a filled outline is one line per edge) fails ALONE and is
 *   reported, rather than aborting the whole deck. A timed-out sync leaves its
 *   context unusable, so recovery HAS to be a fresh context — i.e. the next slide.
 * - Weight. A chunk of four dense charts piled 400-500 shapes into one context;
 *   one chart per context keeps each run light, which the host tolerates better.
 *
 * Off-screen slides, so the extra round-trips are cheap next to reliability. The
 * per-slide sync ORDER still holds (shapes commit before grouping), so a host
 * lacking grouping never rolls back the chart. Requires PowerPointApi 1.3.
 */
/** One item's outcome from a demo-deck insert — the raw material for self-check. */
export interface DemoResult {
  /** Chart shapes actually drawn (0 when skipped as too dense, or failed early). */
  created: number;
  status: "rendered" | "skipped" | "failed";
  /** Wall-clock time for this slide, ms. A value nearing BATCH_TIMEOUT_MS (45s) is
   *  a near-miss stall — the host was one hair from losing it. */
  ms: number;
  /** The item stalled on its first attempt but a single retry recovered it. The
   *  first attempt may have left a stray partial slide (see slidesAdded). */
  retried?: boolean;
  /** The item's sync timed out but a slide readback showed every shape had
   *  landed — the host settled after we gave up, so the chart is real. Counted
   *  as "rendered", NOT "failed": stamping a complete chart NOT COMPLETE is
   *  the false-negative this exists to prevent. */
  lateSettled?: boolean;
  /** The item's sync threw but a readback showed most (≥ PARTIAL_RENDER_THRESHOLD)
   *  of the expected shapes on the slide. A retry would only add a duplicate
   *  slide, so the item is counted as "rendered" and a rescue attempts to
   *  group whatever landed. `created` carries the actual on-slide count so a
   *  console.table row shows "created 31 of 36". */
  partialLanded?: boolean;
  /** The `lastLateSync` value observed during this item, if any — an abandoned
   *  sync that later reported success or a real RichApi error. Read the run's
   *  console.table to see which item stalled and how it eventually resolved. */
  lateOutcome?: string;
  /**
   * The shapes made it onto one native PowerPoint group — the state that makes
   * a chart re-editable, drags-as-one, and lands its POWERCHART_CONFIG tag.
   * False on: hosts without PowerPointApi 1.8 (grouping unsupported),
   * single-shape items (nothing to group), addGroup rejected by the host, or a
   * late-settled render where the group sync never got the chance to run.
   */
  grouped?: boolean;
}

/** A demo-deck insert's self-verification report. */
export interface DemoReport {
  results: DemoResult[];
  /** How much the deck ACTUALLY grew (settled getCount, after − before). */
  slidesAdded: number;
  /**
   * How many slide-adds the run ISSUED: one per item, plus one more for each
   * retried or failed item (both make a second attempt). Comparing this to
   * `slidesAdded` — NOT `results.length` — is what un-masks a lost slide: a
   * retry/fail leaves a stray that inflates `slidesAdded`, so measuring loss
   * against `results.length` reads 0 during real corruption when a stray happens
   * to cancel a lost slide. `addsIssued − slidesAdded` counts adds that did not
   * land (a stray that landed cancels; a swallowed/lost add does not).
   */
  addsIssued: number;
  /**
   * How many of THIS run's `slides.add()` calls `addSlides` itself confirmed
   * missing at a settled, fresh-context readback — even after its one retry
   * round (see `addSlides`, module-scope `lastAddsLost`). A DIFFERENT metric
   * from `addsIssued − slidesAdded`: that pair reads overall deck-growth
   * shortfall for the whole run (a stray from an unrelated retry/fail can
   * cancel a lost slide out of it), while this one is a direct, per-add
   * confirmed-vs-verified deficit — 0 whenever every add either landed or was
   * recovered by the retry.
   */
  addsLostAtCommit: number;
  /**
   * 1-based deck positions of ADDED slides that read back with ZERO shapes — the
   * host committed the slide but its content detached (a silent partial a visual
   * scan misses). Each candidate 0 is re-read once after a short backoff (a
   * struggling host reports transient zeros for hundreds of ms after commit),
   * then recorded only if it re-reads as 0. Honest limits: it cannot see a chart
   * grouped onto ONE shape that lost children, a MERGE of two items onto one
   * slide (that slide isn't blank), or a paint-only blank (office-js#2699 —
   * shapes exist, getCount > 0).
   */
  blankSlides: number[];
  /**
   * Named blanks — same slides as `blankSlides`, enriched with the item title
   * carried on the slot tag (see `DEMO_SLOT_TAG`). A slide without a slot tag
   * (host lacks PowerPointApi 1.3, or the tag itself was lost) appears here
   * with `title: null` — position-only, same as before.
   */
  blankItems: { position: number; title: string | null }[];
  /** False if the blank readback faulted mid-pass — so an empty `blankSlides` is
   *  not mistaken for "no blanks" when it really means "not fully measured". */
  blanksRead: boolean;
  /** Wall-clock time for the whole run, ms — the headline regression metric. */
  totalMs: number;
}

/** The current slide count, read in its own settled sync (reliable on web). */
async function slideCount(): Promise<number> {
  return PowerPoint.run(async (context) => {
    const c = context.presentation.slides.getCount();
    await context.sync();
    return c.value;
  });
}

/**
 * Top-level shape count of a slide, in a fresh settled sync. Used after an
 * addAndRenderItem throw to see what actually landed — a sync that timed out
 * (see `withTimeout`) may leave every shape on the slide anyway, and treating
 * that as a failure both wastes a retry and stamps a real chart NOT COMPLETE.
 */
async function slideShapeCount(index: number): Promise<number> {
  return PowerPoint.run(async (context) => {
    const c = context.presentation.slides.getItemAt(index).shapes.getCount();
    await context.sync();
    return c.value;
  });
}

/**
 * Fresh-context group+tag rescue for a slide whose addAndRenderItem context died
 * with shapes on the slide but no group (a lateSettled render, or a retry whose
 * grouping sync was swallowed). The original context is dead by the time we
 * decide to rescue; open a new one, re-load the slide's shape collection,
 * addGroup them, and write the CHART_TAG so the chart is re-editable.
 *
 * Best-effort: returns false when the host lacks grouping (pre-1.8), addGroup
 * throws, the slide has fewer than 2 shapes, or any sync in the rescue path
 * rejects. Called AFTER the per-item status has been decided — never used to
 * turn a "failed" into a "rendered", only to lift a "rendered" that stayed
 * ungrouped into a re-editable one.
 */
async function rescueGroupAndTag(
  slideIndex: number,
  tagData: string | undefined,
  origin: { left: number; top: number },
): Promise<boolean> {
  if (!supports("1.8")) return false;
  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(slideIndex);
      const shapes = slide.shapes as unknown as PowerPoint.ShapeCollection & {
        items: PowerPoint.Shape[];
        addGroup(items: PowerPoint.Shape[]): PowerPoint.Shape;
      };
      shapes.load("items");
      await context.sync();
      const items = shapes.items;
      if (items.length < 2) return false;
      let group: PowerPoint.Shape;
      try {
        group = shapes.addGroup(items);
      } catch {
        return false;
      }
      group.name = "PowerChart";
      const canTag = supports("1.3") && !!tagData;
      if (canTag) {
        group.tags.add(CHART_TAG, tagData);
        // load left/top to write the CHART_ORIGIN_TAG below.
        group.load("id,left,top");
      }
      try {
        await context.sync();
      } catch {
        // Group did not commit (host refused the addGroup) — leave the slide
        // with its loose shapes. Caller's `grouped` stays false.
        return false;
      }
      if (canTag) {
        group.tags.add(CHART_ORIGIN_TAG, JSON.stringify([origin.left, origin.top, group.left, group.top]));
        try {
          await context.sync();
        } catch {
          /* origin tag didn't land — chart is grouped and re-editable, drag
           * behavior degrades gracefully to updating without an anchor. */
        }
      }
      return true;
    });
  } catch {
    return false;
  }
}

/**
 * Fresh-context unstamp-and-rescue for a slide whose item ended up "failed" —
 * both the render and its retry stalled, but the LAST attempt's readback
 * showed real chart shapes on the slide, just under `PARTIAL_RENDER_THRESHOLD`.
 * The slide already carries the NOT-COMPLETE banner `stampLastSlide` left on
 * it; this deletes that banner and hands the slide to `rescueGroupAndTag` to
 * group + tag whatever chart shapes remain, so a genuinely-partial chart is
 * re-editable instead of permanently hidden under a red stripe.
 *
 * Two fresh contexts, not one: the banner has to be gone (and that deletion
 * committed) before `rescueGroupAndTag` re-loads the shape collection, or its
 * addGroup would sweep the banner into the chart's own group.
 *
 * Returns true only when BOTH steps land: the stamp shape is found and its
 * deletion sync succeeds, AND `rescueGroupAndTag` itself returns true. Any
 * other outcome (no stamp shape found, the delete sync rejects, grouping
 * unsupported/refused) leaves the slide exactly as "failed" left it — still
 * stamped, still ungrouped — so the caller's status/grouped fields stay
 * untouched.
 */
async function unstampAndRescue(
  slideIndex: number,
  tagData: string | undefined,
  origin: { left: number; top: number },
): Promise<boolean> {
  let unstamped: boolean;
  try {
    unstamped = await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(slideIndex);
      const shapes = slide.shapes as unknown as PowerPoint.ShapeCollection & {
        items: (PowerPoint.Shape & { name: string; delete(): void })[];
      };
      shapes.load("items/name");
      await context.sync();
      const stamp = shapes.items.find((s) => s.name === "PowerChart:not-complete");
      if (!stamp) return false;
      stamp.delete();
      await context.sync();
      return true;
    });
  } catch {
    unstamped = false;
  }
  if (!unstamped) return false;
  return rescueGroupAndTag(slideIndex, tagData, origin);
}

/** The blank-layout id, resolved lazily on the first slide's context and reused. */
interface LayoutRef {
  id?: string;
  resolved: boolean;
}

/**
 * Add one slide in a FRESH context and either render the item onto it, or — when
 * the scene is too dense for the web host — stamp a placeholder instead of
 * attempting it (which would burn the timeout and push the host toward a crash).
 * Returns the shapes drawn (0 when stamped). Throws on a host stall; the caller
 * decides whether to retry. A fresh context per call is what makes a retry safe —
 * the failed attempt's context is already discarded.
 */
/**
 * Outcome of one `addAndRenderItem` attempt — the shape count actually drawn,
 * and whether the addGroup landed. The demo self-check surfaces "N landed
 * ungrouped" so a chart that isn't re-editable doesn't silently pass for one
 * that is.
 */
interface AddAndRenderOutcome {
  created: number;
  grouped: boolean;
}

async function addAndRenderItem(
  item: { scene: Scene; tagData?: string; slotTag?: string },
  tooDense: boolean,
  shapeCount: number,
  layout: LayoutRef,
): Promise<AddAndRenderOutcome> {
  let created = 0;
  let grouped = false;
  await PowerPoint.run(async (context) => {
    if (!layout.resolved) {
      layout.id = await blankLayoutId(context);
      layout.resolved = true;
    }
    const [getSlide] = await addSlides(context, 1, layout.id);
    // Slot tag first, so identity survives even a render that later stalls —
    // the blank readback can then name the missing chart, not just its position.
    // Best-effort: a host without Slide.tags (pre-1.3) silently skips.
    if (item.slotTag && supports("1.3")) {
      try {
        (getSlide().tags as unknown as { add(k: string, v: string): unknown }).add(DEMO_SLOT_TAG, item.slotTag);
      } catch {
        /* no slide tags on this host — position-only identity, as before */
      }
    }
    if (tooDense) {
      // Leave a stamped placeholder so the slide count and order still line up.
      await stampSlide(
        context,
        getSlide,
        "NOT COMPLETE",
        `Too dense for this host — ${shapeCount} shapes, not rendered`,
      );
      return;
    }
    const opts: InsertOptions = {
      left: 60,
      top: 90,
      group: true,
      tagData: item.tagData,
      altText: item.scene.desc,
      altTitle: item.scene.title,
      // Off-screen — the host tolerates far larger batches than the live canvas.
      shapesPerSync: SHAPES_PER_SYNC_OFFSCREEN,
    };
    const drawn = await renderShapesChunked(context, getSlide, item.scene, opts);
    created = drawn.length;
    // Shapes are committed by now, so grouping cannot roll them back. Ask
    // groupAndTagAll to re-fetch the shape collection before addGroup — the
    // proxies in `drawn` came from earlier syncs and the web host will silently
    // drop addGroup(theseStaleProxies), leaving the chart loose and untagged.
    // A demo slide is blank at start, so the last N shapes ARE the chart's.
    // Refresh only when the render spanned more than one batch: a single-batch
    // chart's proxies are still in their sync's window at group time, so the
    // reload sync adds no safety and just costs a round-trip. Multi-batch
    // charts (>SHAPES_PER_SYNC) are the ones the web host tripped on — every
    // ungrouped chart in the real run has more than 10 native shapes.
    const needsRefresh = drawn.length > SHAPES_PER_SYNC;
    const [result] = await groupAndTagAll(context, [{ getSlide, created: drawn, opts, refreshShapes: needsRefresh }]);
    grouped = !!result?.grouped;
  });
  return { created, grouped };
}

export async function insertDemoDeck(
  items: {
    scene: Scene;
    tagData?: string;
    title?: string;
    /**
     * Skip the DEMO_SHAPE_BUDGET too-dense check for this item. The budget
     * exists to bail on charts whose wedge/polygon flood times out the web
     * host; text-only scenes (title, contents, results) are much cheaper per
     * shape and should always be attempted, even when the row count pushes
     * them past 90. Without this, the run's own results slide is the first
     * casualty of a failure-heavy run — it's over budget precisely BECAUSE
     * there were failures to record.
     */
    bypassBudget?: boolean;
  }[],
  onProgress?: (done: number, total: number) => void,
): Promise<DemoReport> {
  // Reset the module-scope lost-adds counter at the start of every run, so
  // `addsLostAtCommit` below reports only THIS run's confirmed losses, not a
  // stale accumulation from an earlier insertDemoDeck call.
  lastAddsLost = 0;
  const results: DemoResult[] = [];
  let lastError: unknown;
  const layout: LayoutRef = { resolved: false };
  // Bracket the run with a settled slide count, so a regression run can prove the
  // deck grew by exactly one slide per item — the lost-slide check.
  const before = await slideCount();
  const runStart = Date.now();
  // Running slide count, refreshed with a settled getCount only on the failure
  // path. On the happy path (addAndRenderItem returns) we know exactly one slide
  // landed — no extra round-trip per item.
  let runningCount = before;
  // False once a failure-path getCount rejected: `runningCount` is then a guess,
  // and the rescues below index slides BY it. Grouping the wrong slide writes
  // this item's config tag onto a neighbour's chart — silent cross-chart
  // corruption — so an untrusted count disables the rescues for the rest of the
  // run instead of aiming them at a number we know is stale.
  let countTrusted = true;
  for (let i = 0; i < items.length; i++) {
    const shapeCount = estimateOfficeShapes(items[i].scene);
    const tooDense = !items[i].bypassBudget && shapeCount > DEMO_SHAPE_BUDGET;
    let created = 0;
    let grouped = false;
    let status: DemoResult["status"] = tooDense ? "skipped" : "rendered";
    let retried = false;
    let lateSettled = false;
    let partialLanded = false;
    // Shape count read back off the LAST failed attempt's slide (attempt 1, or
    // the retry if it also failed). Declared here, not inside the catch, so a
    // "failed" outcome below still knows whether the stamped slide has real
    // chart shapes worth rescuing — see the unstamp-and-rescue block after the
    // retried/grouped rescue.
    let readback = 0;
    const lateSeqBefore = lastLateSyncSeq;
    const t0 = Date.now();
    // Slot tag: JSON envelope with the deck-position index and the item title.
    // Written on both attempts (a retry lands on a NEW slide, so its tag has to
    // land too or the retry's slide reads as "unknown item" in a blank check).
    const slotTag = JSON.stringify({ i, title: items[i].title ?? null });
    const itemWithTag = { ...items[i], slotTag };
    try {
      ({ created, grouped } = await addAndRenderItem(itemWithTag, tooDense, shapeCount, layout));
      runningCount++;
    } catch (err) {
      // A throw is not always proof the chart is missing. When `withTimeout`
      // races a slow sync and abandons it, the queued shapes STILL commit —
      // the sync just resolves late. Real host: sync at 60s, shapes on the
      // slide, no error to blame. Readback in a fresh context proves it.
      //
      // Gate on lastLateSync changing: a plain RichApi rejection means the
      // queued commands did NOT commit, and skipping the readback keeps sync
      // counts stable for the existing failure-path tests (no extra syncs on
      // the retry's chosen index). Too-dense items never rendered — skip too.
      // Wait for the abandoned sync (if any) to report its outcome so
      // lastLateSync reflects reality — a timeout that resolves late means
      // the host settled; a plain rejection means nothing to wait for.
      if (!tooDense) await waitForLateSync();
      const lateFired = lastLateSyncSeq !== lateSeqBefore;
      // Re-read the settled slide count on EVERY failure, not only when the
      // shape readback below is worth opening. `runningCount` is what both
      // rescues index slides by, and a throw whose slide DID land (a too-dense
      // item whose stamp sync rejected, a zero-shape scene) otherwise leaves the
      // count one behind for the WHOLE REST of the run: the next ungrouped
      // chart's rescue then reaches the previous item's slide and stamps this
      // chart's config tag onto that one's shapes.
      const afterFail = await slideCount().then(
        (n) => n,
        () => {
          // Nothing settled to trust — freeze the rescues rather than aim them.
          countTrusted = false;
          return runningCount;
        },
      );
      // Read back the failed slide's shape count in a fresh context: the same
      // number decides BOTH "late-settled" (100% + lateFired) and
      // "partial-landed" (>= threshold, < 100%, no lateFired needed). Only
      // opens the shape-count sync when a NEW slide actually appeared, so a
      // catch that fired before addSlides (its getCount rejected) still pays
      // only one round-trip.
      if (!tooDense && shapeCount > 0) {
        const failedSlideIndex = afterFail > runningCount ? afterFail - 1 : -1;
        readback = failedSlideIndex >= 0 ? await slideShapeCount(failedSlideIndex).catch(() => 0) : 0;
      }
      runningCount = afterFail;
      if (readback >= shapeCount && shapeCount > 0 && lateFired) {
        // The host settled after the timeout — every shape is on the slide.
        // Don't retry, don't stamp. The grouping/tagging did not run (the
        // failed context is dead), so this chart is real but ungrouped —
        // the fresh-context rescue below promotes it to re-editable.
        created = readback;
        lateSettled = true;
      } else if (
        !lateSettled &&
        readback > 0 &&
        readback < shapeCount &&
        readback >= Math.ceil(shapeCount * PARTIAL_RENDER_THRESHOLD)
      ) {
        // Most of the chart is on the slide. A retry lands a duplicate slide
        // and rarely helps — the host dropped the LAST batch's sync, not the
        // whole render. Presentation_3.pptx: Line 31/36 shapes (86%), Gantt
        // 23/24 (96%); today's strict `>= shapeCount` gate turned both into
        // dup-slide pairs stamped NOT COMPLETE. Now: rendered-partial, no
        // retry, rescue below groups whatever landed. lateFired NOT required
        // — a plain RichApi rejection can still leave most shapes committed
        // because the batches before it already synced.
        created = readback;
        partialLanded = true;
      }
      if (!lateSettled && !partialLanded) {
        // A host stall is often transient, so retry the RENDER once — a fresh run
        // and a NEW slide, because we must NOT delete the first failed slide: a
        // mis-identified last-slide delete could destroy a good one. A recovered
        // item may thus leave a stray partial slide from attempt 1; slidesAdded
        // surfaces that. Too-dense items only stamp a placeholder — nothing to
        // re-render — so they fail straight through.
        let recovered = false;
        if (!tooDense) {
          try {
            ({ created, grouped } = await addAndRenderItem(itemWithTag, false, shapeCount, layout));
            recovered = true;
            runningCount++;
          } catch (err2) {
            // Second attempt may also be a late-success or a partial-landing —
            // recheck before giving up. Same threshold gate as attempt 1.
            await waitForLateSync();
            const lateFired2 = lastLateSyncSeq !== lateSeqBefore;
            const afterFail2 = await slideCount().then(
              (n) => n,
              () => {
                countTrusted = false;
                return runningCount;
              },
            );
            if (shapeCount > 0) {
              const failedSlideIndex2 = afterFail2 > runningCount ? afterFail2 - 1 : -1;
              readback = failedSlideIndex2 >= 0 ? await slideShapeCount(failedSlideIndex2).catch(() => 0) : 0;
            }
            runningCount = afterFail2;
            if (readback >= shapeCount && shapeCount > 0 && lateFired2) {
              created = readback;
              recovered = true;
              lateSettled = true;
            } else if (
              readback > 0 &&
              readback < shapeCount &&
              readback >= Math.ceil(shapeCount * PARTIAL_RENDER_THRESHOLD)
            ) {
              created = readback;
              recovered = true;
              partialLanded = true;
            }
            /* stalled again — fall through to the failed stamp */
            if (!recovered) lastError = err2;
          }
        }
        if (recovered) {
          retried = true; // status stays "rendered"
        } else {
          // One chart the host would not draw does not sink the rest of the deck.
          // Mark the half-rendered slide so a partial chart is not mistaken for a
          // real one. A fresh context, because the failed render poisoned its own.
          if (lastError === undefined) lastError = err;
          status = "failed";
          await stampLastSlide("NOT COMPLETE", "PowerPoint stopped responding while drawing this chart").catch(
            () => {},
          );
        }
      }
    }
    // Rescue: a "rendered" item that landed ungrouped — its addAndRenderItem
    // context died (lateSettled path) or the group sync was swallowed by the
    // host — gets one more attempt in a fresh context to group + tag its
    // shapes. Turns a loose chart into a re-editable one. Skip for "failed"
    // (its slide has a stamp banner we don't want to group in) and "skipped".
    if (status === "rendered" && !grouped && created > 1 && runningCount > before && countTrusted) {
      const rescued = await rescueGroupAndTag(runningCount - 1, items[i].tagData, { left: 60, top: 90 }).catch(
        () => false,
      );
      if (rescued) grouped = true;
    }
    // Unstamp-and-rescue: a "failed" item whose last attempt still left real
    // chart shapes on the slide (readback > 0) — just under
    // PARTIAL_RENDER_THRESHOLD, so it didn't qualify as rendered-partial above.
    // The slide is currently a stamp banner sitting over a genuine, if
    // incomplete, chart. Delete the banner and group whatever landed, in a
    // fresh context per `unstampAndRescue`. Only promotes a "failed" item —
    // never touches "skipped" (nothing rendered) or an already-"rendered" one
    // (handled by the rescue above).
    if (status === "failed" && readback > 0 && runningCount > before && countTrusted) {
      const rescued = await unstampAndRescue(runningCount - 1, items[i].tagData, { left: 60, top: 90 }).catch(
        () => false,
      );
      if (rescued) {
        status = "rendered";
        partialLanded = true;
        grouped = true;
        created = readback;
      }
    }
    const lateOutcome = lastLateSync && lastLateSyncSeq !== lateSeqBefore ? lastLateSync : undefined;
    results.push({
      created,
      status,
      ms: Date.now() - t0,
      retried,
      lateSettled,
      partialLanded,
      lateOutcome,
      grouped,
    });
    onProgress?.(i + 1, items.length);
  }
  const totalMs = Date.now() - runStart;
  const after = await slideCount();
  const slidesAdded = after - before;
  // One add per item, plus a second for every retried/failed item — the count of
  // adds we ISSUED. Loss is `addsIssued − slidesAdded`, never `items.length −
  // slidesAdded`: a stray from a retry/fail cancels a lost slide against the
  // latter, hiding corruption (observed on a real run: 2 slides lost, reported 0).
  const addsIssued = results.length + results.filter((r) => r.retried || r.status === "failed").length;
  // A whole deck lost to HOST errors (not just skipped-as-dense) is a real
  // failure — surface it so the pane says "Failed", not "inserted 0 of N". If
  // everything was merely skipped, there is no error to throw.
  if (items.length && results.every((r) => r.status !== "rendered") && lastError !== undefined) {
    throw lastError instanceof Error ? lastError : new Error(String(lastError));
  }
  // Blank readback: find added slides the host kept but left EMPTY. Counted by
  // position, not mapped to items — a scrambled deck (lost/merged/reordered
  // slides, plus retry strays) breaks any positional item mapping, and a blank
  // slide has no tag to identify it anyway. Best-effort: a fault leaves
  // blanksRead false so an empty list is not read as "no blanks".
  const { positions: blankSlides, items: blankItems, complete: blanksRead } = await findBlankAddedSlides(before, after);
  return {
    results,
    slidesAdded,
    addsIssued,
    addsLostAtCommit: lastAddsLost,
    blankSlides,
    blankItems,
    blanksRead,
    totalMs,
  };
}

/** Slides read per sync in the readback — kept modest to stay clear of the web
 *  >50-item load ceiling (office-js#4272), though getCount is a scalar, not a load. */
export const READBACK_PAGE = 20;

/** Top-level shape counts for slides [start, end), read in one settled sync. */
async function shapeCounts(start: number, end: number): Promise<number[]> {
  return PowerPoint.run(async (context) => {
    const counts: { value: number }[] = [];
    for (let i = start; i < end; i++) counts.push(context.presentation.slides.getItemAt(i).shapes.getCount());
    await context.sync();
    return counts.map((c) => c.value);
  });
}

/**
 * Backoff before the blank re-read. A struggling host reports transient zeros
 * for hundreds of ms after a shape commit — the first pass may see 0, the
 * settled slide has content. The 200ms wall-clock is under the observed 3%
 * false-blank rate on the real deck; tuneable so tests can pass 0 and skip it.
 */
let BLANK_REREAD_DELAY_MS = 200;

/** Test-only: change the backoff between the blank readback and its re-read. */
export function _setBlankReReadDelayForTest(ms: number): void {
  BLANK_REREAD_DELAY_MS = ms;
}

/**
 * Fetch the slot tag value from a slide by position, in a settled sync. Absent
 * when the host lacks Slide.tags (pre-1.3) or the tag write itself was lost.
 */
async function readSlotTag(index: number): Promise<string | null> {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(index) as unknown as {
      tags: { getItemOrNullObject(k: string): { isNullObject: boolean; value: string; load(): void } };
    };
    try {
      const tag = slide.tags.getItemOrNullObject(DEMO_SLOT_TAG);
      tag.load();
      await context.sync();
      if (tag.isNullObject) return null;
      return tag.value;
    } catch {
      return null;
    }
  });
}

/**
 * Find the ADDED slides (indices [before, after)) the host kept but left EMPTY,
 * returned as 1-based deck positions. Paged to stay clear of the web load ceiling.
 * A struggling host can report a transient 0 for a slide whose shapes have not
 * yet reconciled — a settled re-read after `BLANK_REREAD_DELAY_MS` catches this.
 * `complete` is false if any page's read faulted — so callers don't read an
 * empty list as "no blanks" when it really means "not fully measured".
 *
 * Enriched: each confirmed blank's slot tag is read back, so a caller can name
 * the missing chart by title rather than only its deck position.
 */
async function findBlankAddedSlides(
  before: number,
  after: number,
): Promise<{ positions: number[]; items: { position: number; title: string | null }[]; complete: boolean }> {
  const candidates: number[] = [];
  let complete = true;
  for (let start = before; start < after; start += READBACK_PAGE) {
    const end = Math.min(start + READBACK_PAGE, after);
    try {
      const counts = await shapeCounts(start, end);
      counts.forEach((n, k) => {
        if (n === 0) candidates.push(start + k);
      });
    } catch {
      complete = false; // a page we could not read — do not claim it as clean
    }
  }
  if (candidates.length && BLANK_REREAD_DELAY_MS > 0) {
    // Let the host reconcile before re-reading — see BLANK_REREAD_DELAY_MS.
    await new Promise((r) => setTimeout(r, BLANK_REREAD_DELAY_MS));
  }
  const positions: number[] = [];
  for (let start = 0; start < candidates.length; start += READBACK_PAGE) {
    const page = candidates.slice(start, start + READBACK_PAGE);
    try {
      const again = await PowerPoint.run(async (context) => {
        const cs = page.map((i) => context.presentation.slides.getItemAt(i).shapes.getCount());
        await context.sync();
        return cs.map((c) => c.value);
      });
      page.forEach((i, k) => {
        if (again[k] === 0) positions.push(i + 1);
      }); // 1-based deck position
    } catch {
      complete = false;
    }
  }
  positions.sort((a, b) => a - b);
  // Name each confirmed blank via its slot tag (best-effort; null if untagged).
  const items: { position: number; title: string | null }[] = [];
  for (const pos of positions) {
    const raw = await readSlotTag(pos - 1);
    let title: string | null = null;
    if (raw) {
      try {
        const parsed = JSON.parse(raw) as { title?: string | null };
        title = parsed.title ?? null;
      } catch {
        /* not our shape of tag — leave title null */
      }
    }
    items.push({ position: pos, title });
  }
  return { positions, items, complete };
}

/**
 * `Shape.rotation` is PowerPointApi **1.10**, and the manifests admit hosts from
 * 1.4 — so it must be GATED, not merely wrapped in try/catch.
 *
 * Wrapping catches nothing: Office.js proxy setters do not throw synchronously.
 * The assignment is a queued command the host rejects at the NEXT
 * `context.sync()`, and that sync carries the whole batch — so on a 1.4–1.9 host
 * a single rotated shape took down every pie, doughnut, sunburst, gauge,
 * arrowhead and diagonal line with it, instead of degrading. This is exactly the
 * reasoning `wantsAltText` already applies to `altTextDescription` (same 1.10
 * set); rotation simply never got the same treatment.
 */
const canRotate = (): boolean => supports("1.10");

/**
 * True when the host can paint pixels into a shape: `ShapeFill.setImage` is
 * PowerPointApi **1.8** (@types/office-js: "Sets the fill formatting of the
 * shape to an image. This changes the fill type to `PictureAndTexture`") and the
 * manifests admit hosts from 1.4. Same gate-not-wrap reasoning as `canRotate`.
 *
 * Exported so a caller can skip a pointless rasterisation AND say so: an image
 * insert that silently became native shapes is the failure mode that would be
 * undiagnosable from a bug report.
 */
export const canInsertPicture = (): boolean => supports("1.8");

/**
 * Ceiling on the base64 payload handed to one sync — a GUARD, not a limit we
 * measured on the host. The real cap is undocumented (office-js #225 is fixed
 * with no published threshold), and `BATCH_TIMEOUT_MS`'s rationale ("the budget
 * measures a batch we know the host can swallow") stops holding for a
 * multi-megabyte payload.
 *
 * 4 MB is ~30x the worst payload this engine actually produces. Measured across
 * all 25 kinds at the shipping 480x300pt frame, 2x oversample: 20-58 KB base64
 * (median 31.5 KB); at 4x: 48-133 KB; a full-slide 960x540pt violin at 4x —
 * the densest kind at the largest sane frame — is 0.11 MB. So this fires only
 * for a pathological custom width/height, which is exactly why crossing it
 * emits a specific, greppable code rather than degrading quietly.
 */
const MAX_PICTURE_BASE64 = 4_000_000;

/**
 * Normalise any of the three base64 spellings in circulation to the bare
 * payload Office.js wants:
 *   - `data:image/png;base64,AAAA`  (a browser `canvas.toDataURL`)
 *   - `image/png;base64,AAAA`       (what skill/scripts/render-pptx.mjs passes
 *                                    to pptxgen — no `data:` scheme)
 *   - `AAAA`                        (already bare)
 * Splitting on the LAST comma rather than testing for a `data:` prefix is what
 * makes the middle form work; a `startsWith("data:")` guard would hand the host
 * a payload with `image/png;base64,` still glued to the front.
 */
function barePng(png: string): string {
  const comma = png.lastIndexOf(",");
  return comma >= 0 ? png.slice(comma + 1) : png;
}

/**
 * Whether this insert should draw ONE picture instead of the scene's nodes.
 *
 * Re-checks the requirement set even though callers are expected to, so a caller
 * that passes a payload on a 1.4 host still degrades instead of poisoning the
 * sync. Over-budget payloads and the unsupported host both log a specific code
 * before falling through, because the fallback is otherwise invisible: the chart
 * appears, just made of shapes.
 */
function wantsPicture(opts: InsertOptions, scene: Scene): boolean {
  const png = opts.pictureBase64;
  if (!png) return false;
  if (!canInsertPicture()) {
    console.warn(`PC-IMG-NO-1.8 host lacks PowerPointApi 1.8 (ShapeFill.setImage) — inserted native shapes instead.`);
    return false;
  }
  const bytes = barePng(png).length;
  if (bytes > MAX_PICTURE_BASE64) {
    console.warn(
      `PC-IMG-TOOBIG picture payload ${bytes} B over the ${MAX_PICTURE_BASE64} B guard ` +
        `(chart ${scene.width}x${scene.height}pt) — inserted native shapes instead.`,
    );
    return false;
  }
  return true;
}

/** True when the host advertises the given PowerPointApi requirement set. */
function supports(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", version);
  } catch {
    return false;
  }
}

/**
 * Shapes committed per sync when drawing onto the slide the user is LOOKING at.
 *
 * PowerPoint on the web repaints the live canvas as shapes arrive, and past
 * roughly twenty in one batch it stops answering — the sync never settles and
 * nothing lands at all. Measured against the real host: ~10 shapes insert
 * instantly, the 18-shape table element works, a 30-shape butterfly never
 * commits in 90 seconds. The same shapes go onto NEW slides by the hundred
 * without trouble, because an off-screen slide is not painted.
 *
 * Ten is comfortably under the last known-good (18) and still coarse enough
 * that the round-trips (~0.1s each) disappear next to the drawing.
 */
const SHAPES_PER_SYNC = 10;

/**
 * Off-screen slides (appended by the demo deck or the agenda) don't repaint
 * mid-render, so the host swallows batches ~4-5x larger than the live canvas
 * tolerates. Measured against the real host: a stacked chart at 10 shapes
 * takes 4 syncs, at 40 it takes 1; the extra round-trips (~0.1s each) dominate
 * a 60-slide run's wall clock. 40 keeps a comfortable safety margin against
 * the observed ceiling.
 */
const SHAPES_PER_SYNC_OFFSCREEN = 40;

/**
 * Minimum readback-vs-expected ratio for a caught insert to count as
 * "rendered-partial" instead of retried+stamped. Set at 85% so a single
 * dropped batch (SHAPES_PER_SYNC_OFFSCREEN = 40 / typical ~40-shape chart) is
 * still a visible chart, not a NOT COMPLETE banner. Presentation_3.pptx:
 * Line at 31/36 = 86% and Gantt at 23/24 = 96% both qualify.
 */
const PARTIAL_RENDER_THRESHOLD = 0.85;

/**
 * Draw the whole scene as ONE picture: a rectangle the size of the chart frame,
 * its fill set to the caller's PNG. Returns the single shape, or `[]` when the
 * host refused — the caller then falls through to the native-shape path in the
 * SAME request context, which is safe because this sync carries only this
 * shape's own queued commands (tags come later by design, and
 * updateChartsInSlides renders charts one at a time).
 *
 * The `try` encloses the QUEUEING calls, not just the `await`. Office.js proxy
 * setters do not throw synchronously as a rule, but `setImage` is absent
 * outright on a pre-1.8 host, so the property access itself can throw — and a
 * throw outside the try would take the whole insert down instead of degrading.
 */
async function renderPictureShape(
  context: PowerPoint.RequestContext,
  getSlide: SlideThunk,
  scene: Scene,
  opts: InsertOptions,
): Promise<PowerPoint.Shape[]> {
  const left = opts.left ?? 60;
  const top = opts.top ?? 90;
  let shape: PowerPoint.Shape | undefined;
  try {
    const shapes = getSlide().shapes;
    shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
      left,
      top,
      width: scene.width,
      height: scene.height,
    });
    // Named like a chart group, so the Selection Pane reads the same either way.
    shape.name = "PowerChart";
    (shape.fill as unknown as { setImage(b64: string): void }).setImage(barePng(opts.pictureBase64!));
    // A picture-filled shape must not wear PowerPoint's default outline.
    // UNVERIFIED on a real host: every existing precedent for this assignment is
    // a solid-fill shape, and whether a PictureAndTexture fill keeps the outline
    // suppressed is not knowable from the typings. Wrapped with the rest, so if
    // the host rejects it the whole picture degrades to shapes rather than
    // landing with a stray border.
    shape.lineFormat.visible = false;
    await context.sync();
    return [shape];
  } catch (err) {
    console.warn(
      `PC-IMG-REFUSED host rejected the picture insert — inserting native shapes instead. ${errorText(err)}`,
    );
    // Best-effort cleanup. Whether the rectangle survived is genuinely
    // ambiguous: this repo's fake host models a rejected sync as discarding the
    // whole batch (test/office-render.test.ts), while the partial-landed logic
    // in insertDemoDeck exists precisely because EARLIER batches do commit. Both
    // can be true — the open question is only what happens to commands queued
    // before the failing one in the SAME batch. So: try to delete, ignore the
    // outcome, and never let cleanup failure mask the fallback.
    try {
      shape?.delete();
      await context.sync();
    } catch {
      /* nothing landed, or the context is poisoned — either way the caller
       * draws native shapes next and a stray rect would be visible there. */
    }
    return [];
  }
}

/**
 * Render a scene onto a slide, committing in small batches.
 *
 * This forfeits all-or-nothing, which is a real loss: a failure now strands a
 * partial chart instead of leaving the slide clean. It buys the only thing that
 * matters more — the chart arriving at all — and it is why `created` is
 * returned even on failure, so a caller can clean up what landed.
 *
 * The batches also make progress REAL: shapes committed over shapes total is a
 * fact, not the estimate a single opaque sync would force us to invent.
 *
 * `opts.pictureBase64` short-circuits all of that: one picture, one sync, no
 * batching to do — see `wantsPicture` for when that applies and
 * `renderPictureShape` for the fall-through when the host refuses.
 */
async function renderShapesChunked(
  context: PowerPoint.RequestContext,
  getSlide: SlideThunk,
  scene: Scene,
  opts: InsertOptions,
  onBatch?: (sending: number, total: number) => void,
): Promise<PowerPoint.Shape[]> {
  if (wantsPicture(opts, scene)) {
    // Report the picture as a single unit of work so a caller's progress bar
    // doesn't sit at zero and then jump — the picture IS the whole chart.
    onBatch?.(1, 1);
    const picture = await renderPictureShape(context, getSlide, scene, opts);
    if (picture.length) return picture;
    // Refused: fall through and draw the nodes instead, in this same context.
  }
  const left = opts.left ?? 60;
  const top = opts.top ?? 90;
  const batchSize = opts.shapesPerSync ?? SHAPES_PER_SYNC;
  const created: PowerPoint.Shape[] = [];
  // SHAPES, not nodes. One node is not one shape: a wedge fans out into up to 62
  // triangles and a polygon becomes one line per edge, so slicing scene.nodes by
  // batchSize handed the host ~50 shapes for a pie and 253 for a violin —
  // exactly the flood this batching exists to prevent. (Every batching test used
  // the all-rect `stacked` config, where nodes and shapes happen to be 1:1, so
  // none of them saw it.) Big polygons are split across batches too; each edge is
  // an independent shape, so there is nothing to keep together.
  const steps: ((shapes: PowerPoint.ShapeCollection) => PowerPoint.Shape[])[] = [];
  for (const n of scene.nodes) {
    if (n.kind === "polygon" && n.points.length > batchSize) {
      for (let from = 0; from < n.points.length; from += batchSize) {
        steps.push((sh) => addPolygonEdges(sh, n, left, top, from, batchSize));
      }
    } else {
      steps.push((sh) => addNode(sh, n, left, top, opts));
    }
  }

  const total = estimateOfficeShapes(scene);
  let sent = 0;
  let s = 0;
  while (s < steps.length) {
    // Fresh slide proxy per batch: a proxy held across the previous sync may have
    // been rewritten to an unusable getItem(id) — see SlideThunk.
    const shapes = getSlide().shapes;
    const before = created.length;
    // Always draw at least one step, then keep going while the batch stays under
    // budget — a single indivisible node (a 16-triangle wedge fan) is its own
    // floor, and drawing it alone is the smallest batch available.
    do {
      created.push(...steps[s++](shapes));
    } while (s < steps.length && created.length - before < batchSize);
    sent += created.length - before;
    const upTo = Math.min(sent, total);
    // Reported BEFORE the sync, and deliberately: the sync is where a bad host
    // stops answering, so this is the number that has to be on screen WHILE we
    // wait. Reporting after would leave the pane naming the previous phase and
    // blaming the wrong one for the stall.
    onBatch?.(upTo, total);
    // Budget per BATCH, not per chart: a stalled host must still be caught, but
    // the limit now measures a batch we know the host can swallow.
    await withTimeout(
      context.sync(),
      BATCH_TIMEOUT_MS,
      `drawing shapes ${upTo - (created.length - before) + 1}-${upTo} of ${total}`,
    );
  }
  return created;
}

/** One chart's committed shapes, awaiting grouping and tagging. */
interface Grouping {
  getSlide: SlideThunk;
  created: PowerPoint.Shape[];
  opts: InsertOptions;
  /**
   * Re-load the slide's shape collection right before addGroup, taking the LAST
   * `created.length` shapes as the group's members. The Shape proxies returned
   * by earlier syncs' add*() calls have their object paths rewritten to
   * getItem(id) by the time this sync runs, and the web host can't
   * round-trip those ids — addGroup(theseStaleProxies) silently drops the group.
   * Set true whenever the caller can guarantee the target N shapes are the last
   * N on the slide (a fresh demo slide is the canonical case).
   */
  refreshShapes?: boolean;
}

/**
 * Group the inserted shapes and persist the config tag for ANY number of charts
 * — grouping in one sync, tagging in a second, however many charts there are.
 *
 * Two properties, and the ORDER of these syncs is what buys them:
 * - A host that lacks grouping (e.g. PowerPoint on the web) or tags must never
 *   roll back the already-committed shapes. So the shapes must already be
 *   committed (a prior context.sync) before this runs, and group/tag get their
 *   own syncs after.
 * - Round-trips must not scale with the chart count. Each sync is a trip to
 *   PowerPoint; a per-chart sync is what made Same Scale 4N of them.
 *
 * The cost of batching is granularity: one chart's addGroup throwing now leaves
 * every chart in the batch ungrouped rather than just its own. That is the same
 * outcome the per-chart catch already produced, just wider — and in both cases
 * the charts are on the slide, because their shapes committed a phase earlier.
 */
async function groupAndTagAll(
  context: PowerPoint.RequestContext,
  items: Grouping[],
): Promise<{ target?: PowerPoint.Shape; partIds?: string[]; grouped?: boolean }[]> {
  const tagTargets = items.map((it) => it.created[0] as PowerPoint.Shape | undefined);
  // Which charts actually ended up as one shape. The rest hang everything the
  // group would have carried off their first shape instead — see below.
  const grouped = new Set<number>();
  // Grouping is PowerPointApi 1.8+; skip entirely where unsupported.
  const groupable = items
    .map((it, i) => ({ it, i }))
    .filter(({ it }) => it.opts.group !== false && it.created.length > 1);
  // Re-load shape collections for items that asked to refresh — one sync for
  // all of them, then take the last N shapes as the group's members. See
  // Grouping.refreshShapes: bypasses the stale-proxy trap the web host trips on.
  const refresher = groupable.filter(({ it }) => it.refreshShapes);
  const freshMembers = new Map<number, PowerPoint.Shape[]>();
  if (refresher.length) {
    const collections = refresher.map(({ it }) => it.getSlide().shapes);
    for (const c of collections) c.load("items");
    try {
      await context.sync();
      refresher.forEach(({ it, i }, k) => {
        const items = collections[k].items;
        // The chart's shapes are the LAST N on the slide (fresh demo slides
        // start blank). If items are fewer than expected — e.g. a batch was
        // lost — fall back to `it.created` and hope for the best.
        if (items.length >= it.created.length) {
          freshMembers.set(i, items.slice(items.length - it.created.length));
        }
      });
    } catch {
      /* re-load faulted — fall through to the stale proxies */
    }
  }
  if (groupable.length && supports("1.8")) {
    try {
      for (const { it, i } of groupable) {
        // Fresh slide proxy: grouping runs a sync after the render, by which time
        // a held proxy to a new slide could be stale — see SlideThunk.
        const members = freshMembers.get(i) ?? it.created;
        const group = (
          it.getSlide().shapes as unknown as { addGroup(items: PowerPoint.Shape[]): PowerPoint.Shape }
        ).addGroup(members);
        group.name = "PowerChart";
        // Accessible alt text on the group, queued in this same grouping sync so
        // a screen reader announces the chart — the description the engine built.
        applyAltText(group, it.opts);
        tagTargets[i] = group;
        grouped.add(i);
      }
      await context.sync();
    } catch {
      /* grouping failed — shapes stay ungrouped, charts are already on the slide */
      for (const { i } of groupable) tagTargets[i] = items[i].created[0];
      grouped.clear();
    }
  }
  const partsJson = await ungroupedFallback(context, items, tagTargets, grouped);
  // Tags are PowerPointApi 1.3+; keep the chart re-editable where supported.
  const taggable = items
    .map((it, i) => ({ it, i, target: tagTargets[i] }))
    .filter((t) => t.it.opts.tagData && t.target);
  if (taggable.length && supports("1.3")) {
    try {
      for (const { it, i, target } of taggable) {
        target!.tags.add(CHART_TAG, it.opts.tagData!);
        // The rest of an ungrouped chart travels with the tagged shape, so an
        // in-place update can delete all of it — see CHART_PARTS_TAG.
        if (partsJson[i]) target!.tags.add(CHART_PARTS_TAG, partsJson[i]!);
        // Queued in the SAME sync as the tags, so handing the caller a usable
        // new target costs no extra round-trip. An update replaces every shape,
        // so the caller's old target is dead the moment this returns — see
        // updateChartsInSlides.
        target!.load("id,left,top");
      }
      await context.sync();

      // The frame origin the chart was drawn at, AND the position its tagged
      // shape ended up at ("anchor"). Both are needed, and the anchor is only
      // knowable once the sync above has resolved the shape — hence a second
      // round-trip, batched across every chart in the run.
      //
      // Why both: an update must land the chart where the user LAST LEFT IT.
      // Re-rendering at the tagged shape's corner drifts (that corner is a
      // bounding box, not the frame origin), but re-rendering at the recorded
      // origin teleports a chart the user has since dragged back to where it was
      // first inserted. Storing the anchor lets the update shift the origin by
      // exactly how far the shape has moved since — stable when untouched,
      // faithful when moved.
      for (const { it, target } of taggable) {
        target!.tags.add(
          CHART_ORIGIN_TAG,
          JSON.stringify([it.opts.left ?? 60, it.opts.top ?? 90, target!.left, target!.top]),
        );
      }
      await context.sync();
    } catch {
      /* tags unavailable — charts are inserted but not re-editable */
    }
  }
  return items.map((_, i) => ({
    target: tagTargets[i],
    partIds: partsJson[i] ? (JSON.parse(partsJson[i]!) as string[]) : undefined,
    grouped: grouped.has(i),
  }));
}

/**
 * Set the accessible alt text on the shape that stands for the chart.
 *
 * `Shape.altTextDescription` is PowerPointApi **1.10** — assigning it on an
 * older host is a queued command the host rejects at the NEXT sync, and a sync
 * carries more than this: the grouping, or the config tag. So it is gated
 * rather than merely wrapped, because losing a chart's group (or its
 * re-editability) to gain alt text is the wrong trade.
 */
function applyAltText(shape: PowerPoint.Shape, opts: InsertOptions): void {
  if (!wantsAltText(opts)) return;
  try {
    const a = shape as unknown as { altTextDescription?: string; altTextTitle?: string };
    if (opts.altText) a.altTextDescription = opts.altText;
    if (opts.altTitle) a.altTextTitle = opts.altTitle;
  } catch {
    /* alt text unsupported on this host */
  }
}

const wantsAltText = (opts: InsertOptions): boolean => Boolean(opts.altText || opts.altTitle) && supports("1.10");

/**
 * Give an UNGROUPED chart what the group would have carried: the alt text, and
 * the ids of its other shapes (returned as the tag value to write, per item).
 *
 * The ids have to be read back from the host — a shape's id does not exist
 * client-side until it has been committed — so this costs one sync, paid only
 * where grouping was unavailable or refused. Without them an in-place update
 * deletes 1 of the chart's 13 shapes and redraws all 13, so an ungrouped chart
 * grows by a whole chart on every edit; with them the update deletes the set.
 *
 * Best-effort throughout: if this sync fails, the charts are already on the
 * slide and the config tag still lands in the phase after — only the alt text
 * and the leak-free update are lost, which is exactly today's behaviour.
 */
async function ungroupedFallback(
  context: PowerPoint.RequestContext,
  items: Grouping[],
  tagTargets: (PowerPoint.Shape | undefined)[],
  grouped: Set<number>,
): Promise<(string | undefined)[]> {
  const partsJson: (string | undefined)[] = items.map(() => undefined);
  const loose = (i: number) => !grouped.has(i) && tagTargets[i];
  // Only a re-editable chart (one that gets a config tag, so a host with tags)
  // needs its parts written down; the update path is what reads them.
  const hasTags = supports("1.3");
  const siblings = items.map((it, i) => (hasTags && loose(i) && it.opts.tagData ? it.created.slice(1) : []));
  const alt = items.map((it, i) => ({ it, i })).filter(({ it, i }) => loose(i) && wantsAltText(it.opts));
  if (!alt.length && !siblings.some((s) => s.length)) return partsJson;
  for (const s of siblings.flat()) s.load("id");
  for (const { it, i } of alt) applyAltText(tagTargets[i]!, it.opts);
  try {
    await context.sync();
  } catch {
    /* no alt text or id read-back here — the chart is on the slide regardless */
    return partsJson;
  }
  siblings.forEach((shapes, i) => {
    const ids = shapes.map((s) => s.id).filter((id) => typeof id === "string" && id.length > 0);
    if (ids.length) partsJson[i] = JSON.stringify(ids);
  });
  return partsJson;
}

function getTargetSlide(context: PowerPoint.RequestContext): PowerPoint.Slide {
  try {
    return context.presentation.getSelectedSlides().getItemAt(0);
  } catch {
    return context.presentation.slides.getItemAt(0);
  }
}

/** A straight segment's appearance — the subset of a line node the host needs. */
interface SegmentStyle {
  stroke: string;
  strokeWidth?: number;
  dash?: number[];
  name?: string;
}

/**
 * Draw one straight segment as a native shape, in absolute slide coordinates.
 *
 * PowerPoint's addLine takes only a bounding box, so it can't tell an up-right
 * line from a down-right one — and a zero-thickness box makes the web host
 * substitute a default and draw a giant diagonal. Three cases:
 *
 * - Axis-aligned (baselines, gridlines, connectors, value lines): addLine with
 *   the near-zero dimension clamped. The box is unambiguous.
 * - Dashed diagonal (scatter trend lines, forecast segments, pie breakout
 *   connectors): a real line shape, the only kind that can carry a native dash.
 *   addLine draws the box's top-left→bottom-right diagonal and the lineInverse
 *   geometry draws top-right→bottom-left, so between them the direction is
 *   explicit rather than guessed.
 * - Solid diagonal (line-chart series, the common case): a thin rotated
 *   rectangle, which is direction-correct on every host.
 */
/**
 * Draw a RANGE of a polygon's edges. No freeform paths in Office.js, so the
 * outline becomes connected line segments (translucent fills degrade to
 * outline-only in PowerPoint). These go through addSegment like any other line —
 * passing each edge's bounding box straight to addLine mirrored every up-right
 * edge and gave horizontal ones a zero-height box.
 *
 * The range exists because a single polygon can be far bigger than one batch: a
 * violin body is one node but 82 edges, i.e. 82 shapes. Each edge is an
 * independent shape with no ordering constraint, so the renderer splits them
 * across syncs rather than handing the host 82 at once (see renderShapesChunked).
 */
function addPolygonEdges(
  shapes: PowerPoint.ShapeCollection,
  n: PolygonNode,
  dx: number,
  dy: number,
  from = 0,
  count = Number.POSITIVE_INFINITY,
): PowerPoint.Shape[] {
  const created: PowerPoint.Shape[] = [];
  const pts = n.points;
  const end = Math.min(pts.length, from + count);
  for (let i = from; i < end; i++) {
    const a = pts[i];
    const b = pts[(i + 1) % pts.length];
    created.push(
      addSegment(shapes, dx + a.x, dy + a.y, dx + b.x, dy + b.y, {
        stroke: n.stroke ?? n.fill ?? "#000000",
        strokeWidth: n.strokeWidth,
        name: n.name ? `${n.name}-e${i}` : undefined,
      }),
    );
  }
  return created;
}

function addSegment(
  shapes: PowerPoint.ShapeCollection,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  s: SegmentStyle,
): PowerPoint.Shape {
  const w = Math.abs(x2 - x1);
  const h = Math.abs(y2 - y1);
  const box = {
    left: Math.min(x1, x2),
    top: Math.min(y1, y2),
    width: Math.max(w, 0.5),
    height: Math.max(h, 0.5),
  };
  const setDash = (shape: PowerPoint.Shape) => {
    if (!s.dash) return;
    try {
      // Map to the nearest native line style: a dotted array (e.g. waterfall
      // carry connectors) stays dotted instead of flattening to a generic dash.
      shape.lineFormat.dashStyle =
        dashKind(s.dash) === "dot" ? PowerPoint.ShapeLineDashStyle.roundDot : PowerPoint.ShapeLineDashStyle.dash;
    } catch {
      /* dash style unsupported on this host */
    }
  };

  if (w < 0.5 || h < 0.5 || s.dash) {
    const downRight = (x2 - x1) * (y2 - y1) > 0;
    const line =
      w < 0.5 || h < 0.5 || downRight
        ? shapes.addLine(PowerPoint.ConnectorType.straight, box)
        : shapes.addGeometricShape(PowerPoint.GeometricShapeType.lineInverse, box);
    strokeColor(line.lineFormat, s.stroke);
    line.lineFormat.weight = s.strokeWidth ?? 1;
    setDash(line);
    if (s.name) line.name = s.name;
    return line;
  }

  const len = Math.hypot(x2 - x1, y2 - y1);
  const weight = Math.max(0.5, s.strokeWidth ?? 1);
  const rect = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left: (x1 + x2) / 2 - len / 2,
    top: (y1 + y2) / 2 - weight / 2,
    width: len,
    height: weight,
  });
  solidFill(rect.fill, s.stroke);
  rect.lineFormat.visible = false;
  // Gated, not wrapped — see canRotate. Without it the line renders horizontally.
  if (canRotate()) rect.rotation = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
  if (s.name) rect.name = s.name;
  return rect;
}

/**
 * A colour Office.js will accept. `setSolidColor` / `lineFormat.color` validate
 * only a 6-digit `#RRGGBB` or a named HTML colour — but the engine's allow-list
 * also passes rgb()/hsl()/3-digit/8-digit forms, which would mis-render or be
 * dropped. Named colours go through verbatim (Office.js knows them, and toHex6
 * would flatten them to grey); everything else normalises to `#RRGGBB`.
 */
const officeHex = (color: string): string => (/^[a-zA-Z]+$/.test(color.trim()) ? color.trim() : toHex6(color));

/**
 * Set a solid fill, splitting any alpha (8-digit hex, rgba/hsla) off into
 * `fill.transparency` (PowerPointApi 1.4, the pinned set) so a translucent colour
 * authored in the config matches the SVG preview and the skill's pptx instead of
 * being dropped.
 */
function solidFill(fill: PowerPoint.ShapeFill, color: string): void {
  fill.setSolidColor(officeHex(color));
  const t = 1 - alphaOf(color);
  if (t > 0.0001) {
    try {
      fill.transparency = t;
    } catch {
      /* transparency unsupported on this host — solid is the graceful floor */
    }
  }
}

/**
 * Set a line/border colour AND its alpha. addSegment used to set a line's colour
 * with a bare hex (dropping any alpha) while its rotated-rect fallback went
 * through solidFill (which kept the alpha) — so one series colour rendered at two
 * different opacities depending on segment slope. Routing both through this keeps
 * them consistent, and picks up the rgb()/hsl()/3-digit forms too.
 */
function strokeColor(lf: PowerPoint.ShapeLineFormat, color: string): void {
  lf.color = officeHex(color);
  const t = 1 - alphaOf(color);
  if (t > 0.0001) {
    try {
      lf.transparency = t;
    } catch {
      /* line transparency unsupported — opaque is the graceful floor */
    }
  }
}

function addNode(
  shapes: PowerPoint.ShapeCollection,
  n: SceneNode,
  dx: number,
  dy: number,
  opts: InsertOptions,
): PowerPoint.Shape[] {
  switch (n.kind) {
    case "rect": {
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: dx + n.x,
        top: dy + n.y,
        width: Math.max(0.2, n.w),
        height: Math.max(0.2, n.h),
      });
      // A "none" fill is an outlined/hollow rect (IBCS plan/budget columns) — no
      // fill, just the border below.
      if (n.fill === "none") shape.fill.clear();
      else solidFill(shape.fill, n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        strokeColor(shape.lineFormat, n.stroke);
        shape.lineFormat.weight = n.strokeWidth ?? 1;
      } else {
        shape.lineFormat.visible = false;
      }
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "line":
      return [addSegment(shapes, dx + n.x1, dy + n.y1, dx + n.x2, dy + n.y2, n)];
    case "ellipse": {
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
        left: dx + n.cx - n.rx,
        top: dy + n.cy - n.ry,
        width: Math.max(0.2, n.rx * 2),
        height: Math.max(0.2, n.ry * 2),
      });
      // Stroke-only ellipses (radar circle grid) carry fill "none".
      if (n.fill === "none") shape.fill.clear();
      else solidFill(shape.fill, n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        strokeColor(shape.lineFormat, n.stroke);
        shape.lineFormat.weight = n.strokeWidth ?? 1;
      } else {
        shape.lineFormat.visible = false;
      }
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "chevron": {
      const shape = shapes.addGeometricShape(
        n.flatLeft ? PowerPoint.GeometricShapeType.homePlate : PowerPoint.GeometricShapeType.chevron,
        { left: dx + n.x, top: dy + n.y, width: Math.max(0.2, n.w), height: Math.max(0.2, n.h) },
      );
      solidFill(shape.fill, n.fill);
      shape.lineFormat.visible = false;
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "symbol": {
      // Native preset geometry, so the marker stays FILLED here — the reason a
      // symbol is its own kind rather than a polygon, which PowerPoint can only
      // outline. SYMBOL_PRESET names are GeometricShapeType keys.
      const geo = (PowerPoint.GeometricShapeType as unknown as Record<string, PowerPoint.GeometricShapeType>)[
        SYMBOL_PRESET[n.shape]
      ];
      const shape = shapes.addGeometricShape(geo, {
        left: dx + n.cx - n.size,
        top: dy + n.cy - n.size,
        width: Math.max(0.2, n.size * 2),
        height: Math.max(0.2, n.size * 2),
      });
      solidFill(shape.fill, n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        strokeColor(shape.lineFormat, n.stroke);
        shape.lineFormat.weight = n.strokeWidth ?? 1;
      } else {
        shape.lineFormat.visible = false;
      }
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "wedge":
      return addWedgeFan(shapes, n, dx, dy);
    case "polygon":
      return addPolygonEdges(shapes, n, dx, dy);
    case "text":
      return [addText(shapes, n, dx, dy, opts)];
    case "arrowhead": {
      // No freeform API in Office.js: a rotated geometric triangle whose tip is
      // offset onto (n.x, n.y) about the box centre — see arrowheadBox.
      const box = arrowheadBox(n.x, n.y, n.size, n.angle);
      const shape = shapes.addGeometricShape(PowerPoint.GeometricShapeType.triangle, {
        left: dx + box.left,
        top: dy + box.top,
        width: box.size,
        height: box.size,
      });
      solidFill(shape.fill, n.fill);
      shape.lineFormat.visible = false;
      // Geometric 'triangle' points up (= -90° in scene terms). Gated, not
      // wrapped — see canRotate; without it the arrowhead stays axis-aligned.
      if (canRotate()) (shape as unknown as { rotation: number }).rotation = box.rotation;
      if (n.name) shape.name = n.name;
      return [shape];
    }
  }
}

function addText(
  shapes: PowerPoint.ShapeCollection,
  n: TextNode,
  dx: number,
  dy: number,
  opts: InsertOptions,
): PowerPoint.Shape {
  const shape = shapes.addTextBox(n.text, {
    left: dx + n.x,
    top: dy + n.y,
    width: Math.max(4, n.w),
    height: Math.max(4, n.h),
  });
  shape.fill.clear();
  shape.lineFormat.visible = false;
  const tf = shape.textFrame;
  try {
    tf.wordWrap = false;
    tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    tf.leftMargin = 0;
    tf.rightMargin = 0;
    tf.topMargin = 0;
    tf.bottomMargin = 0;
    tf.verticalAlignment =
      n.valign === "top"
        ? PowerPoint.TextVerticalAlignment.top
        : n.valign === "bottom"
          ? PowerPoint.TextVerticalAlignment.bottom
          : PowerPoint.TextVerticalAlignment.middle;
  } catch {
    /* margin/alignment properties unavailable on this host */
  }
  const font = tf.textRange.font;
  font.size = n.fontSize;
  font.color = officeHex(n.color);
  font.bold = !!n.bold;
  font.name = n.fontFamily ?? opts.fontFamily ?? DEFAULT_FONT;
  try {
    tf.textRange.paragraphFormat.horizontalAlignment =
      n.align === "left"
        ? PowerPoint.ParagraphHorizontalAlignment.left
        : n.align === "right"
          ? PowerPoint.ParagraphHorizontalAlignment.right
          : PowerPoint.ParagraphHorizontalAlignment.center;
  } catch {
    /* paragraph alignment unavailable */
  }
  if (n.name) shape.name = n.name;
  return shape;
}

/**
 * Approximate a pie wedge with a fan of rotated triangles — Office.js has no
 * adjustable pie geometry or freeform paths. Needs Shape.rotation (1.10);
 * older hosts get no wedge. Doughnut holes are separate ellipse nodes
 * emitted by the layout, so wedges here are always full slices.
 */
function addWedgeFan(shapes: PowerPoint.ShapeCollection, n: WedgeNode, dx: number, dy: number): PowerPoint.Shape[] {
  const created: PowerPoint.Shape[] = [];
  // Every shape in the fan is rotated, so on a host without 1.10 there is no fan
  // to draw. Checked ONCE, up front: gated rather than wrapped, because the
  // rejection would otherwise land on the shared sync — see canRotate.
  if (!canRotate()) return created;
  const cx = dx + n.cx;
  const cy = dy + n.cy;
  const span = n.endAngle - n.startAngle;
  // Annular wedge (sunburst ring / gauge): a triangle can't leave a hole, so
  // the band from innerR→r is drawn as radial rectangles; solid slices keep the
  // triangle fan (which tapers to the centre).
  const annular = n.innerR > 0;
  const midR = annular ? (n.innerR + n.r) / 2 : n.r / 2;
  const bandH = annular ? n.r - n.innerR : n.r;
  // Adaptive fan density (chord sagitta under ~0.5pt), capped — see wedgeFanSteps.
  const { steps, step } = wedgeFanSteps(n.r, span);
  for (let i = 0; i < steps; i++) {
    const mid = n.startAngle + step * (i + 0.5);
    // Width sized at the OUTER rim so adjacent shapes meet there and tile into a
    // solid arc — at midR they were half-width on a solid slice and rendered as
    // gapped spokes on the web host. See wedgeFanChord.
    const chord = wedgeFanChord(n.r, step);
    const center = polar(cx, cy, midR, mid);
    try {
      const shape = shapes.addGeometricShape(
        annular ? PowerPoint.GeometricShapeType.rectangle : PowerPoint.GeometricShapeType.triangle,
        {
          left: center.x - chord / 2,
          top: center.y - bandH / 2,
          width: chord,
          height: bandH,
        },
      );
      solidFill(shape.fill, n.fill);
      shape.lineFormat.visible = false;
      // Unrotated the rectangle's height / the triangle's base points south
      // (180° in wedge terms); rotate so it runs along `mid`.
      (shape as unknown as { rotation: number }).rotation = annular ? mid : mid - 180;
      if (n.name) shape.name = `${n.name}-f${i}`;
      created.push(shape);
    } catch {
      /* rotation unsupported — skip the fan on this host */
      break;
    }
  }
  // Best-effort slice outline: the two radial boundary edges as thin rectangles
  // in the stroke colour (stroking every fan seam would web the slice). This
  // reproduces think-cell's thin separators between adjacent slices. Drawn as
  // rotated rectangles, not addLine, since a line's bounding box can't encode a
  // diagonal's direction.
  if (n.stroke && span < 359.9) {
    const eInner = annular ? n.innerR : 0;
    const eLen = n.r - eInner;
    const eMidR = (eInner + n.r) / 2;
    const sw = n.strokeWidth ?? 1;
    for (const ang of [n.startAngle, n.endAngle]) {
      const c = polar(cx, cy, eMidR, ang);
      try {
        const edge = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: c.x - sw / 2,
          top: c.y - eLen / 2,
          width: sw,
          height: eLen,
        });
        solidFill(edge.fill, n.stroke);
        edge.lineFormat.visible = false;
        (edge as unknown as { rotation: number }).rotation = ang;
        if (n.name) edge.name = `${n.name}-edge`;
        created.push(edge);
      } catch {
        /* rotation unsupported — skip the separator */
      }
    }
  }
  return created;
}

/**
 * Read the presentation's theme accent colors (Accent1-6) from the current
 * slide's color scheme — the deck's actual corporate palette. Requires
 * PowerPointApi 1.10 (ThemeColorScheme); returns null on older hosts.
 */
export async function loadThemePalette(): Promise<string[] | null> {
  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const scheme = (slide as unknown as { themeColorScheme: { getThemeColor(c: string): { value: string } } })
        .themeColorScheme;
      const accents = ["Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6"].map((a) =>
        scheme.getThemeColor(a),
      );
      await context.sync();
      const palette = accents
        .map((r) => r.value)
        .filter(Boolean)
        .map((c) => (c.startsWith("#") ? c : `#${c}`).toLowerCase());
      return palette.length >= 3 ? palette : null;
    });
  } catch {
    return null; // no selection, or host below PowerPointApi 1.10
  }
}

/** True when running inside an Office host with the PowerPoint JS API. */
export function isPowerPointHost(): boolean {
  return typeof Office !== "undefined" && typeof PowerPoint !== "undefined" && !!Office.context?.host;
}
