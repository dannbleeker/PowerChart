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
import { polar, arrowheadBox, wedgeFanSteps, wedgeFanChord, symbolPreset, dashKind } from "../core/geometry";
import { estimateOfficeShapes } from "../core/scene";
import { buildChart } from "../core/chart";
import type { ChartConfig } from "../core/types";
import { planSceneUpdate, sceneFingerprint, worthUpdating } from "../core/scene-diff";
import { toHex6, alphaOf, isNamedColor } from "../core/color";
import type { PolygonNode, Scene, SceneNode, TextNode, WedgeNode } from "../core/scene";
import { NOT_COMPLETE_NAME, planReconcile } from "../core/reconcile";
import { trace, traceAbout } from "../core/trace";
import type { Rect } from "../core/placement";
import type { ExpectedItem, ReconcileOptions, ReconcilePlan, SlideSnapshot } from "../core/reconcile";
import { parseSlideSizeEmu, EMU_PER_POINT } from "./ooxml";

/* global PowerPoint, Office */

export interface InsertOptions {
  /** Top-left of the chart frame on the slide, in points. */
  left?: number;
  top?: number;
  /**
   * Draw onto THIS slide instead of the one the user is looking at.
   *
   * The insert path targets the selection, which is right for a user pressing
   * Insert and wrong for a recovery: when an in-place update has already
   * deleted a chart and then failed to redraw it, the replacement has to go
   * back to the slide the chart came from, whatever is on screen by then.
   * Falls back to the selected slide when the id no longer resolves.
   */
  slideId?: string;
  /**
   * Fingerprint of the scene being drawn, written alongside the config tag.
   *
   * Set by whoever HAS the scene — the tagging pass only sees `opts`. Absent is
   * fine and means the chart's next update takes the redraw, which is what
   * every chart drawn before this existed will do once, on its first edit.
   */
  sceneTag?: string;
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
 * THE TAG ANCHOR IS `created[0]`, AND MOVING IT TO THE LAST SHAPE WAS TRIED AND
 * MEASURED AT NO EFFECT. Recorded here because the reasoning is good enough that
 * somebody will have it again — it took four rounds and a host probe to build.
 *
 * The theory: the tag write goes through the handle that DREW the shape, and
 * this host refuses such a handle once a `load()` has resolved it into
 * `shapes.getItem(id)`. The draw loop loaded every shape's id as it went, the
 * anchor included, several batches before the write. So `tagAnchorIndex` moved
 * the anchor to the LAST shape drawn and ran the draw loop's `load("id")` one
 * shape behind, leaving exactly that shape unresolved. Four measured facts
 * supported it and none of them were guesses:
 *
 *   survives-8                                    an unresolved creation handle keeps taking
 *                                                 writes for at least eight syncs
 *   tag-through-refetched-shape: no-id            there is no id to re-fetch a fresh shape by
 *   collection-read-poisons-the-creation-handle:  the pre-grouping re-read does NOT resolve it,
 *     yes                                         so holding one load back is enough
 *   from: created×1, four rounds running          production's failures all went through it
 *
 * **It changed nothing on the host it was built for** — five rounds across four
 * builds, with zero within-build spread, run as same-build pairs so the noise
 * floor was known. It shipped a deferred id load in the draw loop and a
 * contiguity deduction in the pre-grouping matcher for a difference this host
 * does not pay.
 *
 * What the archive says instead, and why this was aimed a level too low:
 * **grouping is what saves a config, not the tag handle** — grouped charts lose
 * it 2% of the time, ungrouped ones 66%. The anchor only ever affected the
 * ungrouped fallback, i.e. the path a chart takes once it has already lost. See
 * `docs/BACKLOG.md`; the fake still models the refusal
 * (`refuseTagWritesOnResolvedProxy`), so a future attempt has its reproduction
 * waiting.
 */

/**
 * A fingerprint of the scene this chart was DRAWN from.
 *
 * Read by the in-place update to decide whether the scene it rebuilds from the
 * stored config is the one actually on the slide. See `sceneFingerprint` — an
 * update that skipped unchanged nodes without this check would leave a chart
 * drawn by an older engine stale forever, where redraw-everything repairs it.
 *
 * Its own tag rather than a field inside the config, because the config is a
 * document the user edits, exports and pastes between decks; a rendering detail
 * has no business in it. A chart without this tag simply takes the redraw.
 */
export const CHART_SCENE_TAG = "POWERCHART_SCENE";

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

/**
 * Put a `getItemOrNullObject` proxy into the next sync, so that `isNullObject`
 * answers once that sync lands.
 *
 * `load("isNullObject")` does NOT do this, and that is the whole reason this
 * function exists. It reads like it should — the property is right there on the
 * proxy and `load` takes any name — but `isNullObject` is not a property the
 * host holds. It is a flag Office.js sets from the RESPONSE to a load of real
 * properties. Ask for it by name and the load selects nothing the host knows,
 * the proxy takes no part in the sync, and reading the flag afterwards throws
 * `PropertyNotLoaded` — with no `errorLocation`, because the getter lives on
 * Office.js's base class and does not know which type it is standing on.
 *
 * That is what the self-test's "edit a chart on the visible slide" died on in
 * PowerPoint on the web: an in-place update resolved the chart's slide, asked
 * for `isNullObject`, and threw reading it back — before it had deleted
 * anything, so the update simply refused. Every other resolve in this file
 * happens to load a REAL property (`value` on a tag, `left,top` on a shape) and
 * has always worked, which is why the failure was confined to the paths that
 * had nothing else to ask for.
 *
 * `id` because every proxy resolved this way is a Slide or a Shape, and both
 * carry one. The value is not used — being in the sync is the point.
 *
 * One correction from a real host, since the paragraphs above are a claim about
 * Office.js and it turns out not to hold everywhere. The host probe's first
 * question asks exactly this, and PowerPoint on the web answered **yes**: a
 * `load("isNullObject")` there populated the flag and read back `false`. So the
 * negative is host-specific rather than universal, and on that host this
 * function is a harmless no-op instead of a necessary workaround. Keep it: the
 * host where it was necessary is also a real one, loading a real property is
 * correct on both, and the cost is a property nobody reads.
 */
function queueNullCheck(proxy: { load(propertyNames: string): void }): void {
  proxy.load("id");
}

/**
 * A collection's `items`, or `undefined` when the host never answered the load
 * that was queued for it.
 *
 * PowerPoint on the web answers a shape-collection read short (see the repair
 * pass's `shapesSeen` counters, and `faults.hollowReads`) — and at the limit it
 * answers with nothing at all. Office.js then leaves the collection unpopulated
 * and the plain `.items` read throws `PropertyNotLoaded` at
 * `ShapeCollection.items`, which took the whole deck-wide chart scan down with
 * it: one silent slide, and Same Scale reported a thrown error instead of
 * rescaling the charts it could see.
 *
 * A collection that did not answer is a collection we know nothing about, so
 * every caller has to decide what to do with the gap rather than inherit a
 * throw. None of them may report it as "no charts here" without saying so.
 */
function loadedItems<T>(collection: { items: T[] }): T[] | undefined {
  try {
    const items = collection.items;
    return Array.isArray(items) ? items : undefined;
  } catch {
    return undefined;
  }
}

/**
 * A loaded property, or `undefined` when the host did not answer for it.
 *
 * Same hazard as `loadedItems`, one level down: a proxy whose load was queued
 * in a sync that failed — or that the host simply did not answer — throws on
 * every property read. Reading through this turns that into a missing value,
 * which callers can fall back from.
 */
function loadedValue<T>(read: () => T): T | undefined {
  try {
    return read();
  } catch {
    return undefined;
  }
}

/**
 * Whether the host CONFIRMED this object exists.
 *
 * Three states collapse to two here, and the direction matters: "the host said
 * it is there" is true, while both "the host said it is gone" and "the host
 * never answered" are false. A caller that cannot tell must behave as if the
 * object is absent — every use of this guards a delete or a redraw, and doing
 * either to a shape we cannot see is how an edit destroys something.
 */
function isLive(proxy: { isNullObject: boolean }): boolean {
  return loadedValue(() => proxy.isNullObject) === false;
}

/**
 * A tag's value, or undefined for "absent, or the host did not say".
 *
 * The `!tag.isNullObject && tag.value` pair was written out at eight call
 * sites, and every one of them read BOTH properties raw. Either read throws if
 * the load did not land, which is a `PropertyNotLoaded` in the middle of a scan
 * — the same shape of failure as the shape-collection one, one level down. A
 * tag nobody can read is a chart that is not re-editable; that is a fact to
 * degrade to, never to throw over.
 */
function tagValue(tag: { isNullObject: boolean; value: string }): string | undefined {
  if (!isLive(tag)) return undefined;
  const v = loadedValue(() => tag.value);
  return typeof v === "string" && v ? v : undefined;
}

/**
 * Where an error was when it escaped — attached to the error itself.
 *
 * Read off by `errorText`, so every place that reports an error says WHICH
 * phase it came out of, not only what the host called it.
 */
const STEP_KEY = "__powerchartStep";

/**
 * Run a host operation under a label, so an error escaping it carries WHERE.
 *
 * The run log used to record what the host said and nothing about what the
 * add-in was doing when it said it. Three self-test scenarios failed on a real
 * host with `PropertyNotLoaded` on three different proxies, and placing each
 * one took reasoning from timestamps and call order back to a line — the log
 * could not say. `errorLocation` names the Office.js type; this names the
 * phase, and between them a failure is located rather than reconstructed.
 *
 * The INNERMOST label wins: an error is annotated once, by the first `step`
 * it unwinds through, so nesting narrows the answer instead of widening it.
 * The same moment is traced, so the log carries the phase in timeline order
 * even when the error is later swallowed by a best-effort catch — which is
 * exactly the case that used to leave nothing behind at all.
 */
async function step<T>(label: string, fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (err) {
    if (err && typeof err === "object" && !(STEP_KEY in err)) {
      try {
        Object.defineProperty(err, STEP_KEY, { value: label, enumerable: false, configurable: true });
      } catch {
        /* a frozen error still carries its own message */
      }
      const text = errorText(err);
      hostFriction.errors += 1;
      // The two refusals this host is made of, counted apart from everything
      // else. Both are named in a dozen places in this file as the cause of a
      // lost chart; neither has ever been a NUMBER a reader could compare
      // between one scenario and the next.
      if (/InvalidParam|GetItem/i.test(text)) hostFriction.idRefusals += 1;
      if (/GeneralException/i.test(text)) hostFriction.generalExceptions += 1;
      trace("error", label, { error: text });
    }
    throw err;
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
  /**
   * Why this target cannot be used to edit the chart again — absent when it can.
   *
   * The chart is ON THE SLIDE in both cases. What is missing is the ability to
   * come back to it, and it used to go missing silently, in green:
   *
   * - `"no-config"` — the shapes landed but `POWERCHART_CONFIG` did not.
   *   `groupAndTagAll` has always answered this honestly (`tagged`), and the
   *   answer had nowhere to go: `EditTarget` carried no such field, so the demo
   *   path (which has a repair pass) consumed it and the everyday insert and
   *   in-place update (which have none) did not. The user gets "Done." in
   *   green, clicking the chart says "the selection is not a PowerChart", and
   *   reopening the deck loses the config permanently. A real host produced
   *   this four times in one run.
   * - `"unknown-shape"` — the chart was redrawn and the host would not say
   *   where. **`shapeId` names the shape this same call already deleted**, so
   *   it must not be kept: the next update would resolve a dead id, or worse,
   *   resolve a group member and draw a second chart over the rest. There is no
   *   correct id to return here — the new one was never read back — so the
   *   honest move is to say the target is spent.
   *
   * Kept as a reason rather than a boolean because the two need different
   * sentences: one says "your chart is fine but not editable", the other says
   * "the pane has lost track of it".
   */
  lost?: "no-config" | "unknown-shape";
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
/**
 * How many bounded waits have hit their deadline since the pane loaded.
 *
 * A COUNT, deliberately, and read by diffing it around a piece of work: "did a
 * deadline fire while that ran" is a fact, and every previous way of asking it
 * was an inference from prose.
 *
 * The self-test's host-sickness breaker is why this exists. It decided whether
 * the host was in trouble by matching the words a scenario happened to choose
 * for its verdict — `/did not respond|gave up/` — and a real run walked
 * straight through it: two scenarios timed out and said so, a third timed out
 * after 49.8 seconds and reported *"the host stopped answering selection
 * calls"*, which matches neither phrase. The counter reset one short of the
 * limit, the battery ran two more scenarios on a host that had been dead for
 * three, and PowerPoint killed the tab and took the remaining verdicts with it.
 *
 * Adding that third phrase to the pattern would have been the same bug with a
 * longer list. There is exactly one place a deadline fires; count it there.
 */
export let deadlinesFired = 0;

/** Test-only: forget the deadlines counted so far. */
export function _resetDeadlinesFiredForTest(): void {
  deadlinesFired = 0;
}

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

/**
 * Which run ISSUED the call that reported the most recent late sync.
 *
 * A demo run degrades to pictures when a call it gave up on answers late — a
 * good signal that the host is drowning. But the counter above is global and
 * carries no owner, so ANY abandoned promise settling during a run bumped it:
 * a chart the user edited before the run started, whose stalled sync happened
 * to answer during item 0, degraded a perfectly healthy 37-item deck to
 * rasters from item 1 and reported "the host answered after we gave up
 * waiting" about an operation the run never made.
 *
 * Captured at ISSUE time, not at settle time — that is what distinguishes a
 * call this run made from one that merely finished during it.
 */
export let lastLateSyncOwner: string | null = null;

/** The run currently issuing host calls, or null outside a demo run. */
let activeRun: string | null = null;

/**
 * The last bounded call the host actually ANSWERED, and when it finished.
 *
 * Every draw stall this project has recorded — thirteen of them, across seven
 * real-host rounds — is the FIRST batch of a scenario's draw, never a later
 * one. Which makes the interesting question "what did the host do immediately
 * before the sync it then refused to answer", and no round file can say: the
 * log records nothing at all between a scenario announcing itself and its first
 * `batch issued`, so three to five seconds of probe reads, deck inventories
 * and selection calls happen invisibly.
 *
 * That absence is why the only account of the stalls anyone could give was at
 * SCENARIO level — "it follows the selection ladder" — and scenario order is
 * fixed, so predecessor and position in the battery are the same variable and
 * nothing in seven rounds separates them. A call-level record does separate
 * them, in every round, with no experiment to design: the stall names the call
 * before it.
 *
 * Read at TIMEOUT time, when it still holds the previous COMPLETED call —
 * a call that never answered never writes here.
 */
let lastAnsweredCall: string | null = null;
let lastAnsweredAt = 0;

/**
 * How long that last answered call TOOK.
 *
 * `afterAnswering` names the call a draw followed and nothing else about it,
 * which left the round of 2026-08-11 unable to split its own control apart: the
 * four arms of `does a rasterise poison the next draw` took 22.7s, 25.6s and
 * 28.9s each, and a rasterise plus a seven-shape draw is a single lump in that
 * number. Which half is growing is the whole question the arm exists to answer.
 *
 * It costs nothing — `withTimeout` already stamps the start of every named call
 * and the moment it answers, so the duration is a subtraction at a seam that is
 * already there. Recorded on the same line as the name, on every first batch and
 * on every stall, so it has the baseline `idleMs` spent two rounds acquiring.
 */
let lastAnsweredMs = 0;

/**
 * How many shapes THIS RUN has already put on each slide.
 *
 * The input to this project's main performance claim and the one number it has
 * never recorded. `what makes a long run slow down` measured that drawing cost
 * grows with the shapes already on the slide — a round costs about
 * `2339 + 630(n-1)` ms for the nth twelve — and every reading of a stalled draw
 * since has had to guess at n. It is free: the draw loop already knows how many
 * shapes it queued and onto which slide.
 *
 * Counts this run's NET contribution to a slide, which is the honest limit and
 * is stated on the trace line: a slide that arrived with fifty shapes reads as
 * zero here.
 *
 * Net, not cumulative, and that took a round to notice. The line used to say
 * "a chart redrawn in place accumulates exactly as the host sees it", which is
 * the opposite of true — a redraw deletes twenty-four shapes and adds
 * twenty-four, the slide does not change size, and the counter went up by
 * twenty-four anyway. Round `957aca0` drove slide 257's counter to **92** while
 * the deck inventory taken at the end of the same run shows that slide holding
 * **3** shapes, so every reading of "shapes already on the slide" taken from an
 * update path was inflated — including the flip in `same scale across the
 * deck`, which is the one measurement that turns on it.
 *
 * `forgetShapesDrawnOn` is the other half. Flooring at zero is what makes the
 * arithmetic right in both directions: a chart this run drew nets out to what
 * it drew, and a chart that predates the run nets out to zero rather than
 * going negative, which is the correct answer for a slide whose size the run
 * has not changed.
 */
const shapesDrawnOnSlide = new Map<string, number>();

/**
 * Take shapes back off a slide's count when the run deletes them.
 *
 * Floored at zero: the deleted shapes are not necessarily ones this run drew,
 * and a count that can go negative is worse than one that under-reports.
 */
function forgetShapesDrawnOn(slideId: string | undefined, n: number): void {
  if (n <= 0) return;
  const key = slideId ?? "(visible)";
  shapesDrawnOnSlide.set(key, Math.max(0, (shapesDrawnOnSlide.get(key) ?? 0) - n));
}

/**
 * How many shapes a redraw's delete took off the slide — which is NOT the
 * number of `delete()` calls it made.
 *
 * Deleting a GROUP takes its children with it in a single call, and a grouped
 * chart's parts tag does not list those children (that is the same absence
 * `tryInPlaceUpdate` refuses on), so the call count says 1 for a chart that
 * occupied twenty-four shapes. Left at 1 the counter keeps 23 shapes that are
 * no longer there, which is the inflation round `957aca0` recorded — slide
 * 257's counter at 92 against 3 shapes in the deck.
 *
 * Two ways of knowing, and the caller cannot pick between them:
 *
 * - **More than one delete went out.** The parts tag enumerated the chart, so
 *   the calls ARE the shapes and the count is exact.
 * - **Exactly one did.** Either a group went (children unknown) or the chart
 *   really was one shape. Both are a chart being replaced by a chart, so the
 *   least-wrong answer is the size of the one going back — a same-size redraw
 *   then nets to zero, which is the truth about the slide, and a chart that
 *   grew nets the growth.
 *
 * `forgetShapesDrawnOn` floors at zero, so the second case cannot drive the
 * count negative when it overshoots (a one-shape picture exploded into
 * twenty-four is the case that does).
 */
export function replacedShapeCount(deleteCalls: number, redrawnShapes: number): number {
  return deleteCalls > 1 ? deleteCalls : redrawnShapes;
}

/**
 * Which slide a draw's shapes should be counted against.
 *
 * `opts.slideId` when the caller named one. Otherwise the slide's OWN id, if
 * the host has answered it — `insertSceneIntoSlide` queues `slide.load("id")`
 * before the first batch, so every batch after the first can have it for free.
 * `(visible)` only when neither is available, which is the first batch of an
 * unnamed draw and any caller that never queued the load.
 *
 * Read through `loadedValue` because an unloaded property THROWS
 * `PropertyNotLoaded` on access rather than answering undefined, and a counter
 * must never be able to take a draw down with it. An empty-string id is
 * rejected for the same reason a null one is: it is not a slide.
 */
export function slideKeyFor(opts: InsertOptions, getSlide: SlideThunk): string {
  if (opts.slideId) return opts.slideId;
  const id = loadedValue(() => (getSlide() as unknown as { id?: string }).id);
  return typeof id === "string" && id ? id : "(visible)";
}

/**
 * How many shapes THIS RUN has already drawn on a slide.
 *
 * Exported so a caller can avoid the slide it has been filling. Drawing cost on
 * this host grows with what is already on the target — about +0.44s per shape
 * present, measured — so a long run that keeps picking the same slide makes
 * each of its own later draws slower than the last, quadratically.
 *
 * The self-test was doing exactly that: eight of its ten scenarios take the
 * FIRST chart the deck scan returns, which is the same chart on the same slide
 * every time. Round `275a76a` recorded that slide going 20 → 68 → 92 → 116 →
 * 140 → 144 → 165 shapes while nothing else in the deck passed 34, and the
 * draw at 144 stalled — a nine-shape chart, timed out at 45 seconds, which
 * ~63 seconds of per-slide overhead entirely accounts for.
 *
 * A count this run kept itself, with no host call to make. It cannot see what
 * the user's deck already held, which is the honest limit: it steers away from
 * slides this run has loaded, not from slides that were busy to begin with.
 */
export function shapesDrawnOn(slideId: string | undefined): number {
  return shapesDrawnOnSlide.get(slideId ?? "(visible)") ?? 0;
}

/**
 * Running counts of the ways this host refuses, since the run began.
 *
 * The self-test's scenarios are the things that actually fail, and they carry
 * no host-state context at all — the probe's `regime` is stamped on probe
 * answers, and no probe answer happens while the battery runs. So a verdict
 * says "the host stopped answering" with nothing beside it saying how the host
 * had been behaving in the seconds before.
 *
 * These are cumulative on purpose. A scenario records the delta across itself,
 * which makes the number comparable between a scenario that passed and one that
 * did not — the property four fields in this project have been built without.
 */
const hostFriction = {
  errors: 0,
  idRefusals: 0,
  generalExceptions: 0,
  emptyReReads: 0,
  /**
   * Re-reads that named SOME of a chart's shapes and not all, after the settled
   * retry had its go. The sibling of `emptyReReads`: the two are the only ways
   * the pre-grouping re-read fails, they drive different branches, and a round
   * that reports one total cannot say which it hit.
   */
  shortReReads: 0,
  /**
   * Charts whose re-read was short or empty FIRST and complete after a pause.
   *
   * The number that says whether `REREAD_RETRY_MS` is buying anything on the
   * real host — measured in production, on every round, rather than argued from
   * the fake. A round where this stays 0 while the two above climb is the round
   * that says the settling-slide theory is wrong.
   */
  reReadsRepaired: 0,
};

/** A snapshot a caller can difference against a later one. */
export function hostFrictionCounts(): Readonly<typeof hostFriction> {
  return { ...hostFriction };
}

/** Reset alongside the stall context — same lifetime, same reason. */
function resetHostFriction(): void {
  hostFriction.errors = 0;
  hostFriction.idRefusals = 0;
  hostFriction.generalExceptions = 0;
  hostFriction.emptyReReads = 0;
}

/**
 * What the host was doing either side of the most recent stall.
 *
 * `idleMs` is measured from the last answer to the moment the stalled call was
 * ISSUED, not to the moment we gave up on it — otherwise every stall would
 * report the 45-second budget back to us and say nothing.
 */
export interface StallContext {
  /** The call that stalled. */
  what: string;
  /** The last call the host answered before it, or null if it never has. */
  afterAnswering: string | null;
  /** Gap between that answer and this call being issued. */
  idleMs: number;
  /** How long that preceding call itself took. */
  afterAnsweringMs: number;
}

export let lastStall: StallContext | null = null;

/**
 * How long the host has been idle — from its last answer to now.
 *
 * Exported because the stall record alone is not evidence. The first round to
 * carry it reported `afterAnswering: "rasterising a slide", idleMs: 1`, which
 * looks damning and may be meaningless: sequential code issues its next call
 * the instant the previous one answers, so a 1ms gap could be true of every
 * draw in the round. A number with no baseline is not a measurement, and this
 * project has been wrong exactly that way before.
 *
 * So the first batch of every draw records the same gap, stalled or not, and
 * one round says whether it discriminates.
 */
export function idleSinceLastAnswer(now = Date.now()): number {
  return lastAnsweredAt ? now - lastAnsweredAt : Infinity;
}

/** Forget the stall record. Test seam, and reset alongside the deadline count. */
export function _resetStallContextForTest(): void {
  lastAnsweredCall = null;
  lastAnsweredAt = 0;
  lastAnsweredMs = 0;
  lastStall = null;
  shapesDrawnOnSlide.clear();
  resetHostFriction();
}

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
  // Whose call this is — read now, while it is being issued. Reading it in
  // `describe` instead would credit the run that happened to be running when
  // the answer arrived, which is exactly the misattribution to avoid.
  const owner = activeRun;
  const describe = (outcome: string): void => {
    lastLateSync = `${what}: ${outcome} after ${Math.round((Date.now() - started) / 1000)}s`;
    trace("host", "a call we gave up on finally answered", { what, outcome, afterMs: Date.now() - started });
    lastLateSyncOwner = owner;
    lastLateSyncSeq += 1;
    lateSubscriber?.(lastLateSync);
  };
  // An ANSWER, not a success: a call the host refused still tells us it was
  // listening, which is the distinction a stall is measured against.
  const answered = (): void => {
    lastAnsweredCall = what;
    lastAnsweredAt = Date.now();
    lastAnsweredMs = lastAnsweredAt - started;
  };
  return new Promise<T>((resolve, reject) => {
    let done = false;
    p.then(
      (v) => {
        if (!done) {
          done = true;
          clearTimeout(timer);
          answered();
          resolve(v);
        } else {
          describe("the host eventually SUCCEEDED");
        }
      },
      (err: unknown) => {
        if (!done) {
          done = true;
          clearTimeout(timer);
          answered();
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
      deadlinesFired += 1;
      // What the host last answered, and how long ago. `lastAnsweredCall` still
      // holds the PREVIOUS completed call here, because this one never wrote to
      // it — which is exactly the fact no round file has ever carried.
      lastStall = {
        what,
        afterAnswering: lastAnsweredCall,
        idleMs: lastAnsweredAt ? started - lastAnsweredAt : Infinity,
        afterAnsweringMs: lastAnsweredMs,
      };
      trace("host", "gave up waiting", {
        what,
        afterMs: ms,
        afterAnswering: lastAnsweredCall ?? "nothing yet this run",
        afterAnsweringMs: lastAnsweredMs,
        idleMs: Math.round(lastStall.idleMs),
      });
      reject(new Error(`PowerPoint did not respond while ${what} (${ms / 1000}s)`));
    }, ms);
  });
}

/**
 * How many of `fullStatements` a trace keeps at each end. See `trimDebugInfo`.
 *
 * Ten and ten, against the batch this project actually issues: a chart is drawn
 * `SHAPES_PER_SYNC` at a time and each shape costs several statements, so ten at
 * the head carries the handles the batch opened with and ten at the tail carries
 * what it was doing when it stopped, without the other eight shapes' noise.
 */
const FULL_STATEMENT_ENDS = 10;

/**
 * Cut `debugInfo.fullStatements` down to something a run log can carry.
 *
 * `extendedErrorLogging` (see `traceEnvironment`) is what makes Office.js fill
 * that array in at all, and it fills it in with the WHOLE batch. One round held
 * 66 of these errors; at a full batch each that is a file nobody can send.
 *
 * BOTH ends, and the first version's tail-only rule is the reason. It assumed
 * the failing statement is last. It is not: Office.js reports it separately in
 * `statement`, and `surroundingStatements` centres its `>>>>>` marker on it —
 * and in the 2026-08-07 round that marker sat on the FIRST statement of the
 * batch while `fullStatements` came back "… 37 earlier statement(s) dropped".
 * The one line worth reading was the one line thrown away.
 *
 * The head is also where the batch's opening handles are, and those are the
 * whole question. A slide printed as `slides.getItem("282#…")` carrying the
 * annotation "originally getItemOrNullObject" is what finally settled whether a
 * printed `getItem` means a held handle. It does not — Office.js annotates the
 * call each path was created by, and the rewrite is just how a resolved proxy
 * prints.
 *
 * Pure, and separate from `errorText`, so the trimming can be tested without a
 * PowerPoint anywhere near it.
 */
export function trimDebugInfo(info: unknown): unknown {
  if (!info || typeof info !== "object") return info;
  const full = (info as { fullStatements?: unknown }).fullStatements;
  if (!Array.isArray(full) || full.length <= FULL_STATEMENT_ENDS * 2 + 1) return info;
  const dropped = full.length - FULL_STATEMENT_ENDS * 2;
  return {
    ...(info as Record<string, unknown>),
    fullStatements: [
      ...full.slice(0, FULL_STATEMENT_ENDS),
      `… ${dropped} statement(s) dropped from the middle`,
      ...full.slice(full.length - FULL_STATEMENT_ENDS),
    ],
  };
}

/**
 * The phase label `step` attached to an error, if it carries one.
 *
 * Exported because a caller sometimes has to tell WHICH operation timed out,
 * not merely that one did. `isTimeout` cannot: every bounded wait produces the
 * same kind of error, so a scenario that catches one and explains it is
 * explaining a guess. The self-test's `edit the chart the user selected` did
 * exactly that — it reported "the host stopped answering selection calls" for a
 * run whose trace says `gave up waiting what=drawing shapes 1-10 of 24`, and
 * sent the reader to two selection bugs that had nothing to do with it.
 */
export function stepOf(err: unknown): string | undefined {
  if (!err || typeof err !== "object") return undefined;
  const at = (err as Record<string, unknown>)[STEP_KEY];
  return typeof at === "string" ? at : undefined;
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
  // The phase the add-in was in, when a `step` recorded one. Office.js says
  // what it refused; this says what we were doing — and a report with only the
  // first half is what made three real-host failures take a session to place.
  const at = (err as Record<string, unknown>)[STEP_KEY];
  if (typeof at === "string") bits.push(`at=${at}`);
  if (e.code) bits.push(`code=${e.code}`);
  if (e.debugInfo) {
    try {
      bits.push(`debugInfo=${JSON.stringify(trimDebugInfo(e.debugInfo))}`);
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
 * How long a SELECTION round trip may take.
 *
 * Much shorter than everything else here, and deliberately so. These calls sit
 * on a click path, they each have a cheap "I do not know" answer already
 * (`null` bounds, the live canvas, a cascade), and a real host stopped
 * answering them twice at ninety seconds and once at ten. Ninety is
 * `READBACK_TIMEOUT_MS`, the budget for a twenty-slide repair page; a question
 * about one shape borrowing it means a user watches "Working…" for a minute and
 * a half over something the pane can shrug off in four seconds.
 *
 * The failure this bounds is not an error. On the web these syncs neither
 * resolve nor reject — office-js#3698 and the wedge measured on this project's
 * own build — so no `catch` can see them, `finally` never runs, and Stop cannot
 * break in because it is checked at batch boundaries this never reaches.
 */
let SELECTION_TIMEOUT_MS = 4_000;

/** Test-only: shorten the selection budget so a wedged host is testable. */
export function _setSelectionTimeoutForTest(ms: number): void {
  SELECTION_TIMEOUT_MS = ms;
}

/**
 * A labelled sync WITH a deadline.
 *
 * `step` says where an error came from; this also says that an answer has to
 * come at all. The two were separate for most of this file's life and the gap
 * between them was the whole post-draw phase — grouping, tagging, the readback
 * — which had a label on every sync and a deadline on none.
 */
/**
 * How many syncs each context has been through — the x-axis the decay curve
 * needs, and the one thing the archive cannot supply.
 *
 * `same scale across the deck` degrades chart by chart inside a single context:
 * charts 1-3 re-read all 24 of their shapes and group, chart 4 matches 20 of 24,
 * chart 5 gets nothing back, and 6-8 are never attempted — identical in six
 * rounds. That is the shape of a context wearing out, and `updateChartsInSlides`
 * was deliberately made ONE context flat in N, so this project's own perf work
 * is the suspect.
 *
 * IT CANNOT BE ASKED FROM A PROBE. Three questions have died on one harness
 * limit — `shapes-items-count-honest`, `which-end-a-short-read-drops` and
 * `how-many-collection-reads-a-context-survives` all answer `unreadable`, the
 * last at its FIRST read — because the scratch slide will not enumerate a
 * collection at all on this host, while a real slide's enumerates fine for the
 * first three charts. The harness is strictly worse than production here, so the
 * measurement has to be taken in production.
 *
 * A WeakMap keyed by the context, so the count dies with it and never leaks
 * across calls. Every sync in this file goes through `boundedSync`, which is
 * what makes one counter enough.
 */
const syncsPerContext = new WeakMap<PowerPoint.RequestContext, number>();

/** How many syncs this context has completed. See `syncsPerContext`. */
export function syncsOf(context: PowerPoint.RequestContext): number {
  return syncsPerContext.get(context) ?? 0;
}

function boundedSync(context: PowerPoint.RequestContext, what: string, budgetMs = BATCH_TIMEOUT_MS): Promise<void> {
  syncsPerContext.set(context, syncsOf(context) + 1);
  return step(what, () => withTimeout(context.sync(), budgetMs, what));
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
): Promise<EditTarget | null> {
  // Stamp what this draw produces, so the chart's next update can diff against
  // the scene actually on the slide rather than against whatever the stored
  // config renders to by then. A caller that set its own wins.
  opts = { sceneTag: sceneFingerprint(scene), ...opts };
  onPhase?.("context");
  // Filled inside the run, acted on after it — a chart whose config tag the
  // drawing context could not write is settled and tagged from a fresh one.
  // See `settleAndTagChart`: on PowerPoint web this is the difference between
  // an inserted chart and a re-editable one.
  const untagged: { key: string; slideId: string; tagData: string; shapeId?: string }[] = [];
  const inserted = await PowerPoint.run(async (context) => {
    // Held for the whole draw, and that is a KNOWN, UNFIXED risk when the
    // caller names a freshly-added slide.
    //
    // A per-batch thunk was tried here and reverted, because on this path it is
    // worse than what it fixes. `groupAndTagAll` re-reads the slide's shape
    // collection through this same thunk one sync before it groups, and a shape
    // proxy carries its parent's object path — so handing out a fresh handle
    // per batch makes every re-read member an orphan by grouping time. Measured
    // on a 40-shape chart inserted onto a freshly-added slide: the chart landed
    // ungrouped with no CHART_PARTS_TAG, and the next edit left the old chart's
    // 39 shapes orphaned under the new one — 41 shapes became 79. Nothing in
    // the suite asserts grouping or parts on this path, which is how 2016 tests
    // passed over it.
    //
    // The question that decided this has been ASKED AND ANSWERED, and the risk
    // it named is now empty:
    //
    // - `getTargetSlide` resolves by `slides.getItem(id)`, and the 2026-08-08
    //   sheet says that call answers `threw` (GeneralException) for a slide
    //   added moments earlier and `yes` for a pre-existing one
    //   (`shape-add-fresh-getitem-slide` and its follow-up partner
    //   `getitem-durable-slide`). So a fresh slide is not merely a risk here —
    //   it does not work, by any route.
    // - NO caller passes a freshly-added slide's id any more. `chartIsVisible`
    //   was the only one, and it stopped: it does its before-and-after on a
    //   slide the run added earlier, because rasterising a fresh slide killed
    //   PowerPoint five rounds running. The demo path does not come through
    //   here at all — it draws via `drawDemoItem` with its own positional
    //   thunk — and every other caller passes no id, or a slide the user was
    //   already editing.
    //
    // So the hold stays, and it is no longer a bet: the case it was unsafe for
    // has no caller. If one is ever added, it must resolve its slide some other
    // way — `getItem(id)` is not available for a new slide on this host.
    const slide = getTargetSlide(context, opts.slideId);
    slide.load("id");
    const getSlide: SlideThunk = () => slide;
    onPhase?.("queue", `${scene.nodes.length} nodes`);
    // Committed in batches: the whole scene in one sync is what a live canvas
    // will not take. Each batch reports, so progress here is measured, not
    // guessed — see renderShapesChunked.
    const created = await step("drawing the chart's shapes", () =>
      renderShapesChunked(context, getSlide, scene, opts, (done, total) =>
        onPhase?.("commit", `${done} of ${total} shapes`),
      ),
    );
    // Shapes are committed by now, so grouping/tagging (which some hosts,
    // notably PowerPoint on the web, don't support) can't roll back the chart.
    onPhase?.("group");
    const [tagged] = await step("grouping and tagging the new chart", () =>
      groupAndTagAll(context, [
        {
          getSlide,
          created,
          opts: { ...opts, altText: scene.desc, altTitle: scene.title },
          // Any chart that will be GROUPED, not just one that crossed a sync
          // boundary — this host refuses a creation proxy for addGroup whatever
          // its age. See `needsPreGroupRefresh`.
          refreshShapes: needsPreGroupRefresh(created, opts),
        },
      ]),
    );
    onPhase?.("done");
    // Hand back an edit target, so a caller that inserted as a RECOVERY (see
    // updateChartResilient) can keep the chart editable instead of telling the
    // user to go and find it.
    //
    // Null only when there is no id at all. It used to be null whenever the
    // host would not TAG the chart too — the comment here said so — and that
    // stopped being true when the `unresolved` re-read was added to
    // `groupAndTagAll`: an untagged chart now comes back with a perfectly good
    // shape id. Which left the caller no way to tell a re-editable chart from
    // one carrying no config, so it printed "Done." over both. `lost` is that
    // difference, carried rather than discarded.
    const t = tagged?.target;
    // The slide's own id, which the load at the top asked for. A host that did
    // not answer leaves nothing to re-open — the same answer as an untagged
    // chart, and better than an EditTarget naming a slide we cannot name.
    const slideId = loadedValue(() => slide.id);
    if (!t || typeof slideId !== "string") return null;
    if (!tagged?.tagged && opts.tagData) untagged.push({ key: t.id, slideId, tagData: opts.tagData, shapeId: t.id });
    return {
      slideId,
      shapeId: t.id,
      left: t.left,
      top: t.top,
      partIds: tagged?.partIds,
      origin: { left: opts.left ?? 60, top: opts.top ?? 90, anchorLeft: t.left, anchorTop: t.top },
      ...(tagged?.tagged ? {} : { lost: "no-config" as const }),
    };
  });
  // After the run, never inside it: the context that could not write the tag is
  // the one thing that certainly cannot write it now.
  const settled = (await settleUntaggedCharts(untagged)).size;
  // A chart that is re-editable again must not still be reported as lost —
  // `lost: "no-config"` is what makes the pane tell the user their chart cannot
  // be re-opened. Decided from the settle's own committed sync, not from a
  // second readback: `settleAndTagChart` returns true only after its write
  // landed, which is the same bar `taggedOk` uses inside the drawing context.
  if (inserted?.lost === "no-config" && settled > 0) {
    const { lost: _lost, ...reEditable } = inserted;
    return reEditable;
  }
  return inserted;
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
 * What an update hands back for a chart whose group-and-tag step produced no
 * target at all — three different situations that look identical from here.
 *
 * Extracted because getting one of the three wrong is invisible in a green
 * suite and expensive on a real host, and because the state that produces it
 * needs four simultaneous host failures to reproduce end-to-end. The rule is
 * pure; it should be checkable without a PowerPoint.
 *
 * - **Drew, then lost its identity.** `groupAndTagAll` could not group it,
 *   could not tag it, and could not even read an id back — a real PowerPoint
 *   did all three in one scenario on 2026-08-06, ending `reading back where the
 *   charts landed | InvalidParam passed to GetItem(id)`. The chart is on the
 *   slide and carries no config, so it is `no-config`: THERE, but not
 *   re-openable. This case used to return the target bare, and that hole is why
 *   the settle pass never ran on that host — `lost` stayed undefined, the
 *   `no-config` filter matched nothing, and `settleUntaggedCharts` returned
 *   before it could even trace. Five "tagging failed" events in one run and not
 *   one settle attempt, which from outside looked exactly like a settle that
 *   had tried and failed. Two entirely different bugs; only one was real.
 * - **Wrecked before it drew.** A stop broke in after the delete committed, so
 *   the shapes are gone and the old `shapeId` names something deleted:
 *   `unknown-shape`. Returning it bare was the original trap — the pane kept it
 *   as the live edit target, printed "Done." in green, and the next push
 *   resolved a dead id and told the user their chart was gone.
 * - **Neither.** A stop broke in BEFORE the delete committed. The chart still
 *   has its shapes and its old target is still true, so it comes back untouched.
 */
export function targetWithNoTagResult(old: EditTarget, o: { drew: boolean; wrecked: boolean }): EditTarget {
  if (o.wrecked && !o.drew) return { ...old, lost: "unknown-shape" };
  if (o.drew) return { ...old, lost: "no-config" };
  return old;
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
 * The phases are ordered, and that order is load-bearing: every old shape
 * resolves before any is deleted, and every new shape COMMITS before anything is
 * grouped — so a host without grouping cannot roll back the charts themselves.
 * Batching happens across charts WITHIN a phase, never across phases.
 *
 * The one thing NOT batched across charts is the delete, and that is deliberate
 * — see the loop below.
 *
 * `onFailed` is called for each chart whose redraw did not finish. Charts are
 * independent here: one that stalls no longer takes the rest of the batch with
 * it. The call throws only when EVERY chart failed, which is what keeps the
 * single-chart wrapper's contract — `updateChartResilient` needs that throw to
 * reach its slide-swap and picture fallbacks.
 */
/**
 * What a failed in-place update left on the slide.
 *
 * An update deletes before it draws, and the delete COMMITS. So a redraw that
 * dies half way is not a no-op that can be retried — it is a hole in the deck
 * with some of a chart in it. Every recovery above this layer needs to know
 * that, and none of it could: the call reported the same empty result for "the
 * user deleted this chart, nothing to do" as for "I deleted it and could not
 * put it back", so the pane told the user their chart was gone while half of it
 * sat on the slide, and the picture fallback re-resolved a shape id that this
 * very call had destroyed.
 */
export interface UpdateWreckage {
  /** The slide the chart was on — where a replacement has to go back. */
  slideId: string;
  /** Where it sat, so the replacement lands in the same place. */
  at: { left: number; top: number };
  /** The partial chart the failed redraw committed, by shape id. */
  strayIds: string[];
}

const WRECKAGE_KEY = "__powerchartWreckage";

const STOPPED_KEY = "__powerchartStopped";

/**
 * Whether the user has asked the work in flight to stop.
 *
 * Cooperative, and it has to be: Office.js has no abort. A `context.sync()`
 * already handed to PowerPoint runs to completion (or to `BATCH_TIMEOUT_MS`)
 * whatever we want, so the only honest place to stop is BETWEEN batches — and
 * the only honest promise to make the user is "no more after this one".
 *
 * Module state rather than a token threaded through every signature, for the
 * same reason `lastAddsLost` is: one pane, one host, one operation at a time.
 * `guard()` clears it before every action, so a stop can never leak into the
 * next one.
 */
let stopRequested = false;

/** Ask the render in flight to stop at its next batch boundary. */
export function requestStop(): void {
  stopRequested = true;
  trace("pane", "stop requested", {});
}

/** Clear any pending stop. Called at the start of every guarded pane action. */
export function resetStop(): void {
  stopRequested = false;
}

export function isStopRequested(): boolean {
  return stopRequested;
}

/**
 * The error a stopped render throws.
 *
 * Marked with an own property rather than an `instanceof` check: the same
 * reason `wreckageOf` uses one. `wreck()` may re-carry a stall on a fresh Error
 * of its own, and a marker survives that where a subclass would not.
 */
function stopped(): Error {
  const e = new Error("Stopped.") as Error & Record<string, unknown>;
  e[STOPPED_KEY] = true;
  return e;
}

/** True when this error is the user's own stop, not a host failure. */
export function isStopped(err: unknown): boolean {
  if (!err || typeof err !== "object") return false;
  return Object.prototype.hasOwnProperty.call(err, STOPPED_KEY);
}

/** Throw if a stop is pending. Call at a point where stopping is SAFE. */
function throwIfStopped(): void {
  if (stopRequested) throw stopped();
}

/**
 * Tag the host's own error with what the failure destroyed, rather than
 * replacing it. The message, `code` and `debugInfo` of a RichApi.Error are the
 * only evidence anyone gets about a host that stopped answering — wrapping it
 * in an error of our own would throw exactly that away.
 */
function wreck(
  err: unknown,
  destroyed: boolean,
  target: EditTarget,
  at: { left: number; top: number },
  drawn: PowerPoint.Shape[],
): unknown {
  if (!destroyed) return err;
  const strayIds: string[] = [];
  for (const s of drawn) {
    // Only the batches that COMMITTED have ids: `load("id")` resolves on the
    // batch's own sync, and the batch that stalled never got one. Those shapes
    // are unaddressable, so they are not sweepable either — read past them
    // rather than letting a PropertyNotLoaded lose the ids we do have.
    try {
      if (typeof s.id === "string" && s.id) strayIds.push(s.id);
    } catch {
      /* never came back from the host — nothing to sweep by */
    }
  }
  const wreckage: UpdateWreckage = { slideId: target.slideId, at, strayIds };
  if (err && typeof err === "object") {
    try {
      (err as Record<string, unknown>)[WRECKAGE_KEY] = wreckage;
      return err;
    } catch {
      /* frozen error — fall through and carry the wreckage on a fresh one */
    }
  }
  const carrier = new Error(errorText(err)) as Error & Record<string, unknown>;
  carrier[WRECKAGE_KEY] = wreckage;
  return carrier;
}

/** What an update destroyed before it failed, if it destroyed anything. */
export function wreckageOf(err: unknown): UpdateWreckage | undefined {
  if (!err || typeof err !== "object") return undefined;
  // hasOwnProperty, not `in` or a bare read: an error whose prototype chain
  // happens to carry this name would otherwise hand back something that is not
  // a wreckage at all, and the recovery path would act on it.
  if (!Object.prototype.hasOwnProperty.call(err, WRECKAGE_KEY)) return undefined;
  const w = (err as Record<string, unknown>)[WRECKAGE_KEY];
  return w && typeof w === "object" ? (w as UpdateWreckage) : undefined;
}

/**
 * Delete named shapes from one slide, in a context of its own.
 *
 * A fresh context on purpose: this runs after a redraw that stalled, and the
 * context it stalled in is exactly the one that cannot be trusted to carry a
 * repair. Best-effort — a stray that will not go is worth less than the
 * recovery that follows it.
 */
export async function deleteShapesById(slideId: string, ids: string[]): Promise<number> {
  if (!ids.length) return 0;
  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemOrNullObject(slideId);
      queueNullCheck(slide);
      await context.sync();
      if (!isLive(slide)) {
        trace("insert", "wreckage sweep could not resolve the slide", { slideId, strays: ids.length });
        return 0;
      }
      const shapes = ids.map((id) => slide.shapes.getItemOrNullObject(id));
      for (const s of shapes) queueNullCheck(s);
      await context.sync();
      let gone = 0;
      let unresolved = 0;
      for (const s of shapes) {
        // Through `isLive`, not a raw `s.isNullObject`. Every id here names a
        // shape that COMMITTED, so it is on the slide by construction — a host
        // that will not resolve it has not told us the stray is gone. The raw
        // read was worse than wrong: on an unanswered proxy it throws
        // `PropertyNotLoaded`, which took the whole sweep down and left the
        // catch below reporting 0 for the addressable strays too.
        if (!isLive(s)) {
          unresolved++;
          continue;
        }
        s.delete();
        gone++;
      }
      await context.sync();
      // The strays were drawn by this run by construction, so the slide's count
      // owes back every one the sweep committed — and this is the sweep that
      // runs before `slideHoldsOnlyChart`, which now reads that count.
      forgetShapesDrawnOn(slideId, gone);
      // Say what was left. Both callers use this sweep as the precondition for
      // what they do next — the picture fallback draws over whatever is still
      // there, `slideHoldsOnlyChart` gates the slide swap on the slide being
      // clean, and Stop's promise that nothing is left behind rests on it —
      // and all three used to be handed a bare number that could not say
      // whether it meant "nothing to do" or "could not touch any of it".
      if (unresolved) trace("insert", "wreckage the host would not resolve", { slideId, unresolved, swept: gone });
      return gone;
    });
  } catch {
    return 0;
  }
}

export async function updateChartsInSlides(
  items: { scene: Scene; target: EditTarget; opts?: InsertOptions }[],
  onFailed?: (item: { scene: Scene; target: EditTarget; opts?: InsertOptions }, err: unknown) => void,
): Promise<EditTarget[]> {
  if (!items.length) return [];
  // Same contract as the insert path: charts whose config tag the drawing
  // context could not write, settled and tagged from a fresh one afterwards.
  // This is the path `same scale across the deck` drives, and the one that
  // reported "3 of 8 charts carry the shared scale" on a real host.
  const untagged: { key: string; slideId: string; tagData: string; shapeId?: string }[] = [];
  /**
   * Charts that reached the slide and came back with NO tagged target.
   *
   * Their `shapeId` is the caller's old one, naming a shape this call deleted,
   * so the settle must not be handed it — `settleAndTagChart` would resolve a
   * dead id and give up without ever trying the collection it could still have
   * searched. Tracked by index because that is the only thing left that
   * identifies them.
   */
  const untargeted = new Set<number>();
  /**
   * Each returned target's own `tagData`, in the SAME index space as `updated`.
   *
   * `items` is not that space. See where this is filled, inside the map below.
   */
  const planned: (string | undefined)[] = [];
  const updated = await PowerPoint.run(async (context) => {
    // 1. Resolve every old shape — one sync for all of them.
    //
    // getItemOrNullObject, never getItem: a target names a slide that the user
    // may have deleted, undone, or closed since we read it, and getItem THROWS
    // on a stale id — "InvalidParam passed to GetItem(id)", code 5010, which is
    // a normal condition wearing a crash's clothes. A chart whose slide is gone
    // is not an error, it is nothing to do.
    const found = items.map((it) => {
      const slide = context.presentation.slides.getItemOrNullObject(it.target.slideId);
      queueNullCheck(slide);
      return { it, slide };
    });
    await step("resolving the charts' slides", () => context.sync());

    // Confirmed present, never merely "not confirmed gone". A slide the host
    // would not answer for is one we cannot safely delete a chart off.
    const live = found.filter(({ slide }) => isLive(slide));
    if (!live.length) return [];
    const withOld = live.map(({ it }) => {
      // A FRESH handle for the lookups, not the one the liveness check above
      // just resolved — that one is a sync old by now, and a sync-old by-id
      // handle on a slide added this session is what the web host refuses. The
      // check keeps using its own proxy, because deciding whether the slide is
      // there is exactly what it was resolved for.
      const shapes = context.presentation.slides.getItemOrNullObject(it.target.slideId).shapes;
      const old = shapes.getItemOrNullObject(it.target.shapeId);
      // The shape's LIVE position, queued in the sync this resolution already
      // costs. The caller's EditTarget is a snapshot — the task pane holds one
      // from whenever the chart was loaded — so measuring the user's drag against
      // it reports no movement for a chart that has since been dragged, and the
      // update puts it back where it was. Only the host knows where the shape is
      // now.
      old.load("left,top");
      // The config this chart was drawn from, and the fingerprint of the scene
      // it produced — both read in the sync that was already resolving the
      // shape, so the fast path below costs no extra round trip when it does
      // not apply. See `tryInPlaceUpdate`.
      //
      // Wrapped, because `old` may name a shape the user has since deleted and
      // reading `.tags` off a resolved-to-nothing proxy THROWS at queue time —
      // synchronously, out of the whole PowerPoint.run. That is not a theory:
      // the same access is what produced "Cannot read properties of undefined
      // (reading 'add')" on a real host, four times in one run. The liveness
      // check that would have caught it is a sync away, and this has to be
      // queued before it. No tags read is simply no fast path.
      let wasConfig: PowerPoint.Tag | undefined;
      let wasScene: PowerPoint.Tag | undefined;
      try {
        wasConfig = old.tags.getItemOrNullObject(CHART_TAG);
        wasConfig.load("value");
        wasScene = old.tags.getItemOrNullObject(CHART_SCENE_TAG);
        wasScene.load("value");
      } catch {
        wasConfig = undefined;
        wasScene = undefined;
      }
      // An ungrouped chart is more than its tagged shape (see CHART_PARTS_TAG).
      // Resolved in this same sync, so the delete below already knows which of
      // them the user has since removed by hand.
      const parts = (it.target.partIds ?? []).map((id) => shapes.getItemOrNullObject(id));
      // Each part has to be LOADED, not merely resolved. A getItemOrNullObject
      // proxy nobody loads takes no part in the sync, so `isNullObject` is never
      // populated and reading it throws PropertyNotLoaded — which is what
      // editing any UNGROUPED chart did, and PowerPoint on the web ungroups
      // every chart it cannot group. `old` above only escaped it by accident:
      // its load("left,top") is what put it in the sync — a REAL property,
      // which is the only kind that counts. See `queueNullCheck`.
      for (const p of parts) queueNullCheck(p);
      return { it, old, parts, wasConfig, wasScene };
    });
    await step("resolving the charts' shapes", () => context.sync());

    // A target whose SHAPE is gone gets the same treatment as one whose slide
    // is gone: nothing to do. Re-rendering it would resurrect a chart the user
    // deleted — an in-place update that inserts is not an update.
    // Read the live positions off the proxies BEFORE the delete below detaches
    // them; from here on `at` is where each chart actually sits on the slide.
    const alive = withOld
      .filter(({ old }) => isLive(old))
      .map((e) => {
        // The live position, falling back to the caller's snapshot when the
        // host answered for the shape but not for where it is. The snapshot is
        // stale for a chart the user has since dragged, which costs that chart
        // its drag delta — a chart drawn back at its recorded origin, not a
        // failed update. Losing the whole redraw over an unread number is the
        // worse trade.
        const left = loadedValue(() => e.old.left) ?? e.it.target.left;
        const top = loadedValue(() => e.old.top) ?? e.it.target.top;
        return { ...e, at: { left, top } };
      });
    if (!alive.length) return [];

    // 2-3. Per chart: drop its old shapes, then redraw it. Both, before moving
    //    to the next chart.
    //
    //    Deleting every chart's old shapes up front was one sync cheaper and a
    //    great deal worse. Those deletes COMMIT; the redraws then run one at a
    //    time, and a single stalled redraw rejected the whole PowerPoint.run —
    //    leaving every chart after it in the batch blank, its old shapes gone
    //    and its new ones never queued. Same Scale runs across the whole deck
    //    and necessarily includes the chart on the visible slide, which is the
    //    one documented condition that reliably stalls a redraw. So the cheap
    //    ordering turned one slow chart into a deck-wide wipe.
    //
    //    Deleting per chart costs one extra sync each (against the many
    //    `renderShapesChunked` already spends) and bounds the damage to the
    //    chart that actually failed: the ones after it are still untouched.
    //
    //    Siblings go with it: deleting only the tagged shape of an ungrouped
    //    chart leaves the rest of it on the slide, under the redraw.
    //
    //    The redraw itself is batched. One of these charts is on the slide the
    //    user is looking at, and a live canvas will not take a whole chart in
    //    one sync — so the batching is not an optimisation here, it is the only
    //    way the shapes arrive at all. Per chart, because a chart's shapes must
    //    all reach the same slide.
    const rendered: Grouping[] = [];
    /** Index in `alive` → index in `rendered`. Absent means that chart failed. */
    const placed = new Map<number, number>();
    /**
     * Charts whose OLD shapes are committed gone.
     *
     * The difference between "this target is stale" and "this target is
     * untouched", and the two are not interchangeable. A stop taken before a
     * chart's delete costs it nothing — its shapes are still on the slide and
     * its target still names them — while a chart whose delete committed and
     * whose redraw then failed has a target pointing at nothing. Marking both
     * as lost would tell the user their chart had gone when the stop had in
     * fact protected it, which is the opposite of the truth.
     */
    const wrecked = new Set<number>();
    let firstFailure: unknown;
    for (const [i, entry] of alive.entries()) {
      // Stop BEFORE this chart's delete, and break rather than throw. Every
      // chart past here is still whole — untouched old shapes, nothing queued —
      // so the cheapest, safest stop in the whole add-in is the one taken at
      // this line. The charts already redrawn keep their new targets; the rest
      // keep their old ones, which is exactly what the return below does for a
      // chart that was never reached.
      if (isStopRequested()) break;
      const { it, old, parts, at } = entry;
      const opts: InsertOptions = {
        ...it.opts,
        // Which slide this redraw is aimed at. Read ONLY by the per-slide shape
        // counter in the trace — the slide itself comes from `getSlide` below,
        // and must go on doing so: `getTargetSlide`, the other reader of this
        // field, resolves by `slides.getItem(id)`, which this host refuses for
        // any slide a run has just added.
        //
        // Absent until now, so the counter pooled every chart in a deck-wide
        // rescale under one `(visible)` key and reported 260 shapes on a slide
        // that held 24.
        slideId: it.target.slideId,
        // What this redraw will have drawn, for the NEXT update to diff against.
        sceneTag: sceneFingerprint(it.scene),
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
      // Re-acquired per batch, always. The comment here used to say "an
      // existing slide's proxy is stable across syncs — hold it. Only a
      // freshly-added slide needs a per-batch fresh proxy", and every word of
      // that is true except the assumption underneath it: that this path only
      // ever edits slides the user has had for a while.
      //
      // It does not. The demo deck's slides are added minutes before someone
      // edits a chart on one, and the self-test's `same scale` scenario updates
      // charts on slides the battery itself has just inserted. On PowerPoint
      // web that is the one case a held by-id handle refuses — `GeneralException`
      // at `SlideCollection.getItem`, thrown out of the whole update rather than
      // degraded — and it is why editing a chart worked in every test and every
      // hand-check: those all used a slide that was already in the deck.
      //
      // A fresh by-id resolve costs nothing and works on both kinds of slide;
      // the host probe's `shape-add-fresh-slide-proxy` says yes and
      // `shape-add-held-slide-proxy` threw.
      // Write only what changed, when that is provably the same thing as
      // redrawing. Before the delete, so a refusal costs nothing: nothing has
      // been removed yet and the loop below does the whole job as it always
      // has. See `tryInPlaceUpdate` for every reason it says no.
      if (
        await tryInPlaceUpdate(
          context,
          entry,
          opts,
          { config: entry.wasConfig && tagValue(entry.wasConfig), scene: entry.wasScene && tagValue(entry.wasScene) },
          step,
        )
      )
        continue;
      const getSlide: SlideThunk = () => context.presentation.slides.getItemOrNullObject(it.target.slideId);
      // Everything the redraw manages to commit, whether or not it finishes.
      // On the failure path this is the litter to clear; see the catch.
      const drawn: PowerPoint.Shape[] = [];
      let deleted = false;
      try {
        old.delete();
        // Only the siblings the host CONFIRMED are there. One it would not
        // answer for is left alone: deleting a shape we cannot see risks
        // taking something that is not ours, and the redraw covering a stray
        // is a visible, fixable outcome where a wrong delete is neither.
        let removed = 1;
        for (const p of parts)
          if (isLive(p)) {
            p.delete();
            removed++;
          }
        await step("deleting the chart being replaced", () => context.sync());
        // Committed gone, so the slide's count owes them back.
        forgetShapesDrawnOn(it.target.slideId, replacedShapeCount(removed, estimateOfficeShapes(it.scene)));
        // From here the old chart is committed GONE. Anything that throws below
        // leaves a hole, and the caller has to be told which chart it is.
        deleted = true;
        wrecked.add(i);
        const created = await step("redrawing the chart's shapes", () =>
          renderShapesChunked(context, getSlide, it.scene, opts, undefined, drawn),
        );
        // Only once it landed, or `placed` would point at a chart that isn't
        // in `rendered` and every target after it would come back mismatched.
        placed.set(i, rendered.length);
        rendered.push({ getSlide, created, opts });
      } catch (err) {
        // This chart's old shapes are committed gone and its redraw did not
        // finish, so it is blank and nothing here can undo that. The charts
        // after it can still be saved, which is the whole point of the loop.
        //
        // Wreck EVERY failure, and hand the wrecked error to `onFailed` — not
        // the raw one. Both used to be true only of the first: `wreck` ran
        // inside the `firstFailure === undefined` test, so failures 2..n were
        // never annotated at all, and `onFailed` was passed `err` regardless.
        // A caller doing a deck-wide update is exactly the caller that gets
        // more than one failure, and `wreckageOf` on what it received returned
        // undefined every time — so it could not sweep, and left a half-drawn
        // chart on each stalled slide while reporting them empty. The
        // single-chart path never noticed because it only ever has failure #1.
        const wrecked = wreck(err, deleted, it.target, at, drawn);
        if (firstFailure === undefined) firstFailure = wrecked;
        trace("draw", "chart update failed mid-batch", { index: i, error: errorText(err) });
        onFailed?.(it, wrecked);
      }
    }
    // Nothing landed at all: that is the single-chart case, and its caller
    // recovers by catching. Swallowing it would strand `updateChartResilient`
    // on layer 1 with a chart it just deleted.
    if (!rendered.length && firstFailure !== undefined) throw firstFailure;

    // 4-5. Group, then tag — one sync each, however many charts.
    const tagged = await step("grouping and tagging the redrawn charts", () =>
      groupAndTagAll(
        context,
        // An update redraws every shape, so the same multi-batch staleness that
        // costs a fresh insert its group costs an edit its group too.
        rendered.map((r) => ({ ...r, refreshShapes: r.refreshShapes || needsPreGroupRefresh(r.created, r.opts) })),
      ),
    );

    // 6. Hand back the NEW targets. An update replaces every shape, so the
    //    caller's target is dead as soon as this returns: the pane used to keep
    //    the old one, and its next update resolved a shape id that no longer
    //    existed, was filtered out as "the user deleted this chart", and did
    //    nothing at all — silently. Auto-update died the same way after its
    //    first push. Returning the new target is what lets a caller stay live.
    return alive.map(({ it, at }, i) => {
      // The config THIS entry's chart carries, captured here and nowhere else.
      //
      // The settle plan below used to read `items[i]`, and `i` counts `alive`,
      // not `items` — `alive` is what survives the live-slide and live-shape
      // filters. Drop one item and every later pair is off by one, so the
      // settle wrote chart N's POWERCHART_CONFIG onto chart N+1's shape, then
      // returned true, which cleared `lost` and reported the chart fully
      // re-editable. The user opens a good chart, gets someone else's data, and
      // the next Update writes it in. That is precisely what
      // `settleByCollectionRead`'s own header forbids, reached from the far side.
      //
      // Both triggers are ordinary on the web host: the user deletes one chart
      // and runs Same Scale (the `alive` filter drops it), or the host declines
      // to resolve one shape with the 5010 this repo has logged 66 times in a
      // single run.
      planned[i] = it.opts?.tagData;
      // Through `placed`, never by position: a chart that failed above is
      // absent from `rendered`, so `tagged` is shorter than `alive` and
      // indexing it directly would hand every chart after the failure the
      // NEXT chart's shape id — an edit target pointing at somebody else's
      // chart, which the next update would then overwrite.
      const ri = placed.get(i);
      const t = ri === undefined ? undefined : tagged[ri]?.target;
      // No id for the new shape. The caller's old target is the only thing left
      // to hand back, and its `shapeId` names the shape THIS CALL DELETED — so
      // it comes back marked spent. Returning it bare was a trap: the pane kept
      // it as the live edit target and printed "Done." in green, and the next
      // push resolved a dead id, was filtered out as "the user deleted this
      // chart", and told them their chart was gone. There is no correct id here
      // — the new one was never read back — so the signal is the fix, not the
      // value.
      //
      // `wrecked` is what keeps this honest. A chart the stop broke out BEFORE
      // still has its shapes and its target; only one whose delete committed has
      // lost them. An existing test caught a version of this that marked both.
      if (!t) {
        const back = targetWithNoTagResult(it.target, { drew: placed.has(i), wrecked: wrecked.has(i) });
        if (back.lost === "no-config") untargeted.add(i);
        return back;
      }
      // The origin this pass actually rendered at, paired with where the tagged
      // shape landed — the same (origin, anchor) contract groupAndTagAll wrote to
      // the tag, so the caller's in-memory target and the on-slide tag agree.
      const o = it.target.origin;
      return {
        slideId: it.target.slideId,
        shapeId: t.id,
        left: t.left,
        top: t.top,
        partIds: ri === undefined ? undefined : tagged[ri]?.partIds,
        origin: {
          left: o ? o.left + (at.left - o.anchorLeft) : at.left,
          top: o ? o.top + (at.top - o.anchorTop) : at.top,
          anchorLeft: t.left,
          anchorTop: t.top,
        },
        // Redrawn and located, but the config tag did not land — so the chart is
        // on the slide and cannot be re-opened. Same signal the insert path
        // carries, for the same reason.
        ...(ri !== undefined && tagged[ri]?.tagged ? {} : { lost: "no-config" as const }),
      };
    });
  });
  // Outside the run, for the reason the insert path is: the context that could
  // not write the tag cannot write it now either.
  const wanted = updated
    // `planned`, not `items` — the two are different index spaces the moment a
    // single item is filtered out, and mixing them tagged the wrong chart.
    .map((t, i) => ({ t, i, tagData: planned[i] }))
    .filter(({ t, tagData }) => t.lost === "no-config" && tagData);
  for (const { t, i, tagData } of wanted)
    untagged.push({
      key: String(i),
      slideId: t.slideId,
      tagData: tagData!,
      // Omitted for a chart whose new shapes were never read back: its
      // `shapeId` names the shape THIS call deleted. Handing that to the settle
      // buys a guaranteed failure on a dead id in place of the collection
      // search that is the only thing left with a chance.
      shapeId: untargeted.has(i) ? undefined : t.shapeId,
    });
  if (!untagged.length) return updated;
  const settledIds = await settleUntaggedCharts(untagged);
  // Per chart, by INDEX: a run where one settle landed and another did not must
  // not report both as re-editable, and the pane reads `lost` per target. Shape
  // id cannot do this job — the charts that most need the settle are exactly
  // the ones with no id to be keyed by.
  return updated.map((t, i) => afterSettle(t, { settled: settledIds.has(String(i)), untargeted: untargeted.has(i) }));
}

/**
 * What a target is worth once the settle has had its turn.
 *
 * A settle that lands puts the config back on the slide, so the chart IS
 * re-editable again and `no-config` should go. That is what this did — and it
 * stripped the marker without asking whether the target it was clearing still
 * NAMED anything.
 *
 * For a chart whose new shapes were never read back, `shapeId` names the shape
 * THIS call deleted; `untargeted` is the set of exactly those. Handing one back
 * bare is the trap the update path already documents twenty lines up: the pane
 * keeps it as the live edit target and prints "Done." in green, and the next
 * push resolves a dead id, is filtered out as "the user deleted this chart",
 * and tells them their chart is gone. The settle succeeding made that WORSE,
 * because it is the one path that cleared the marker protecting against it.
 *
 * `unknown-shape` is the honest answer and the pane already says the right
 * thing for it — "PowerPoint would not say where the new chart landed … click
 * the chart and press Edit it to carry on" — which is now true rather than
 * pessimistic: the config did land, so clicking the chart works.
 */
export function afterSettle(t: EditTarget, o: { settled: boolean; untargeted: boolean }): EditTarget {
  if (t.lost !== "no-config" || !o.settled) return t;
  if (o.untargeted) return { ...t, lost: "unknown-shape" };
  const { lost: _lost, ...ok } = t;
  return ok;
}

/**
 * The deck's own id for a slide named by a SELECTION.
 *
 * office-js#2474: a `SlideRange`'s `id` is not roundtrippable — it lacks the
 * `#XYZ` suffix the same slide carries when read from `presentation.slides`, so
 * `slides.getItem(rangeId)` answers InvalidArgument where `getItemAt(index)`
 * works. Reported on Windows desktop, closed `not planned`, and this repo hands
 * exactly such an id back as an EditTarget's `slideId` from
 * `loadChartFromSelection` — the pane's most-used read.
 *
 * The failure it produces is the silent kind. An update resolves that id with
 * `getItemOrNullObject`, gets a null object, and the chart is filtered out as
 * "the slide is gone, nothing to do": the user clicks their chart, edits it,
 * and nothing happens, with no error anywhere. On this web host the ids happen
 * to round-trip today — `edit the chart the user selected` has passed every
 * round — which is precisely why this must not be left to luck on a host
 * nobody here has run.
 *
 * The repair is the issue's own observation turned into a rule: the deck's id
 * is the range's id plus a `#suffix`, so an id that is not in the deck's list
 * is matched by that prefix. Pure, because the alternative is discovering on
 * someone's desktop that Edit does nothing.
 */
export function deckIdForSelectedSlide(rangeId: string | undefined, deckIds: string[]): string | undefined {
  if (!rangeId) return undefined;
  if (deckIds.includes(rangeId)) return rangeId;
  const withSuffix = deckIds.filter((id) => id.startsWith(`${rangeId}#`));
  // Exactly one, or nothing. Two slides answering to one prefix means the
  // assumption behind this whole repair is wrong on that host, and guessing
  // between them would put the user's edit on the wrong slide — which is worse
  // than the silent no-op this exists to prevent.
  return withSuffix.length === 1 ? withSuffix[0] : undefined;
}

/**
 * Read the PowerChart config back from the current selection (the tag written
 * at insert time). Returns null when the selection is not a PowerChart.
 * Requires PowerPointApi 1.5 (getSelectedShapes).
 */
export async function loadChartFromSelection(
  budgetMs?: number,
): Promise<{ configJson: string; target: EditTarget } | null> {
  return boundedRun(
    "reading the selected chart",
    async (context) => {
      const slides = context.presentation.getSelectedSlides();
      const slide = slides.getItemAt(0);
      slide.load("id");
      // The DECK's ids too, in the sync this read already costs. A SlideRange's
      // id is not roundtrippable (office-js#2474) and the target built below is
      // resolved by id on every later edit — see `deckIdForSelectedSlide`.
      const deck = context.presentation.slides;
      deck.load("items/id");
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/id,items/left,items/top");
      await context.sync();

      // The host may answer a shape-collection read with nothing at all; then
      // there is no selection to speak of, which is the same answer as "the
      // selection is not a PowerChart". See `loadedItems`.
      const selected = loadedItems(shapes) ?? [];
      const tags = selected.map((s) => chartTagsOf(s));
      await context.sync();

      // The selected slide's id, once. A host that will not name the slide leaves
      // no target to hand back — the pane's "not a PowerChart" answer, which is
      // survivable, rather than a throw out of the pane's most-used read.
      const rangeId = loadedValue(() => slide.id);
      // Matched against the deck's own list rather than trusted. A host that
      // will not answer for the deck leaves the range's id as the only thing
      // there is, which is the behaviour every round on this host has had.
      const deckIds = (loadedItems(deck) ?? [])
        .map((sl) => loadedValue(() => sl.id))
        .filter((id): id is string => !!id);
      const slideId = deckIds.length ? deckIdForSelectedSlide(rangeId, deckIds) : rangeId;
      if (rangeId && slideId !== rangeId)
        trace("pane", "the selected slide's id is not the deck's id for it", { rangeId, slideId: slideId ?? null });
      for (let i = 0; i < selected.length; i++) {
        const { config, parts, origin } = tags[i];
        const configJson = tagValue(config);
        const at = targetRef(selected[i]);
        if (!configJson || !at || typeof slideId !== "string") continue;
        return {
          configJson,
          target: {
            slideId,
            shapeId: at.id,
            left: at.left,
            top: at.top,
            partIds: partIdsOf(parts),
            origin: originOf(origin),
          },
        };
      }
      return null;
    },
    budgetMs,
  );
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
  const raw = tagValue(parts);
  if (!raw) return undefined;
  try {
    const ids: unknown = JSON.parse(raw);
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
  const raw = tagValue(origin);
  if (!raw) return undefined;
  try {
    const v: unknown = JSON.parse(raw);
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
    // Bounded, and on the short selection budget. This was the only selection
    // read in the file not on `boundedRun`, and it is the FIRST host call the
    // Insert button makes — so a host that went quiet on it took the whole
    // insert with it: buttons disabled, "Working…" counting up forever, nothing
    // drawn and nothing said, and `guard()` has no deadline on `fn()` either.
    // Its answer is already optional (null means "no placeholder to fit"), so
    // giving up costs nothing but the placeholder convenience.
    return await boundedRun(
      "reading the selection's bounds",
      async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/left,items/top,items/width,items/height");
        await context.sync();
        const items = loadedItems(shapes);
        if (!items || items.length !== 1) return null;
        const s = items[0];
        const tag = s.tags.getItemOrNullObject(CHART_TAG);
        tag.load("value");
        await context.sync();
        if (tagValue(tag)) return null; // it's a chart — edit, don't cover
        return { left: s.left, top: s.top, width: s.width, height: s.height };
      },
      SELECTION_TIMEOUT_MS,
    );
  } catch {
    return null;
  }
}

/** All PowerCharts in the current selection (for Same Scale on a subset). */
export async function listChartsInSelection(): Promise<{ configJson: string; target: EditTarget }[]> {
  return boundedRun("reading the charts in the selection", async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("id");
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/left,items/top");
    await context.sync();
    const selected = loadedItems(shapes) ?? [];
    const tags = selected.map((s) => chartTagsOf(s));
    await context.sync();
    const slideId = loadedValue(() => slide.id);
    const charts = selected
      .map((s, i) => ({ at: targetRef(s), configJson: tagValue(tags[i].config), ...tags[i] }))
      // A shape whose position the host would not answer for cannot become an
      // edit target: every later update resolves it by id and re-renders at
      // that corner. Dropping it reports one chart fewer; keeping it would
      // report a chart that cannot be edited.
      .filter((c) => !!c.configJson && !!c.at && typeof slideId === "string")
      .map((c) => ({
        configJson: c.configJson!,
        target: {
          slideId: slideId as string,
          shapeId: c.at!.id,
          left: c.at!.left,
          top: c.at!.top,
          partIds: partIdsOf(c.parts),
          origin: originOf(c.origin),
        },
      }));
    for (const t of tags) {
      untrack(t.config);
      untrack(t.parts);
      untrack(t.origin);
    }
    for (const s of selected) untrack(s);
    return charts;
  });
}

/**
 * Every shape's rectangle on the slide the next insert would draw onto.
 *
 * What the placement rule needs, and nothing else: the insert path had no idea
 * what was already on the slide, so its only answer to "where does this chart
 * go" was a fixed 14pt cascade that stacked charts on top of each other.
 *
 * **`null` when the host would not answer, and never `[]`.** This used to
 * return an empty array on any refusal, with a comment claiming placement then
 * "falls back to the cascade, which is what it always did". That sentence was
 * false, and the failure it hid is the exact pile the placement rule exists to
 * prevent: with nothing occupied, `placeBeside` succeeds on its first pass and
 * hands back the origin UNMOVED, so the caller's cascade is unreachable. Two
 * inserts onto a slide the host will not describe therefore land on precisely
 * the same point, one on top of the other — worse than the fixed cascade it
 * replaced. A real host refused every shape read on a whole deck
 * (`unread=8 slides=8`), so this is its ordinary behaviour, not a corner.
 *
 * "No shapes here" and "I could not look" have to reach the caller as different
 * answers, because only the second one means "do not trust the geometry".
 */
export async function getSlideShapeBounds(slideId?: string): Promise<Rect[] | null> {
  try {
    // On the selection budget, not the readback one: this runs on the Insert
    // click, its answer is optional, and a host that will not give it should
    // cost four seconds and a cascade rather than ninety and a dead button.
    return await boundedRun(
      "reading what is already on the slide",
      async (context) => {
        const slide = getTargetSlide(context, slideId);
        slide.shapes.load("items/left,items/top,items/width,items/height");
        await context.sync();
        const items = loadedItems(slide.shapes);
        if (!items) return null;
        return items
          .map((s) => ({ left: s.left, top: s.top, width: s.width, height: s.height }))
          .filter((r) => [r.left, r.top, r.width, r.height].every((n) => typeof n === "number" && Number.isFinite(n)));
      },
      SELECTION_TIMEOUT_MS,
    );
  } catch {
    return null;
  }
}

/**
 * What a deck-wide chart scan found, AND what it could not see.
 *
 * The second half is the point. `readChartsPage` has always been careful to
 * separate "no charts on this slide" from "this slide told us nothing" — and
 * then the count went to a trace and the function handed back a bare array. The
 * trace is off by default, so in ordinary use that distinction reached nobody:
 * Same Scale took its min and max from whatever answered and reported success
 * over a deck that provably did not share one scale, and the self-test's
 * "nothing was left behind" assertion could not fail once the scan went blind.
 * A real host produced both — `unread=8 slides=8`, then `unread=7 slides=8`.
 *
 * Three ways a scan can be short, kept apart because they are found by
 * different reads and a caller may care about only one:
 *
 * - `unread`  — the slide would not describe itself at all.
 * - `short`   — the slide's shape collection came back with fewer shapes than
 *               the slide's own count. Observed: `shapesExpected=19 shapesSeen=15`.
 * - `tagsUnread` — a shape's config tag would not answer either way, so
 *               "not a chart" and "could not tell" were the same answer.
 */
export interface DeckScan {
  charts: { configJson: string; target: EditTarget }[];
  /** Slides the host would not describe at all. */
  unread: number;
  /** Slides whose shape collection answered SHORT of the slide's own count. */
  short: number;
  /** Shapes whose config tag the host would not answer either way. */
  tagsUnread: number;
  /** Slides in the deck, so a caller can say "7 of 8" rather than "7". */
  slides: number;
  /**
   * Every shape the scan saw, slide by slide — only when it was asked for.
   *
   * The scan already has this in its hands and used to throw it away, keeping
   * the shapes that turned out to be charts and dropping the rest. That is the
   * right trade for the callers that rescale or repair charts, and the wrong one
   * for a diagnostic: "41 shapes became 79", "the chart landed ungrouped", "the
   * slide still holds what was there before" are all questions about the shapes
   * that are NOT charts, and answering them has meant asking the owner to save
   * the deck and upload it.
   *
   * Off by default. `items/name` is a per-shape string deck-wide, and the paths
   * that scan every slide on a live web host are exactly the ones that must not
   * pay for a diagnostic.
   */
  inventory?: SlideInventory[];
}

/** What one slide holds, as the deck scan saw it. */
export interface SlideInventory {
  slideId: string;
  /** Position in the deck at scan time — how a reader finds it by clicking. */
  index: number;
  shapes: { id: string; name?: string; left?: number; top?: number }[];
  /** The slide's own count, when the host gave one — `shapes.length` short of it means a partial answer. */
  count?: number;
}

/**
 * Whether a scan saw the whole deck.
 *
 * Every caller that draws a conclusion from what a scan did NOT contain has to
 * ask this first. "No charts here" and "I could not look" are the same array.
 */
export function scanIsComplete(scan: DeckScan): boolean {
  return scan.unread === 0 && scan.short === 0 && scan.tagsUnread === 0;
}

/** What a scan missed, in words, or "" when it missed nothing. */
export function scanGap(scan: DeckScan): string {
  const parts: string[] = [];
  if (scan.unread) parts.push(`${scan.unread} of ${scan.slides} slide(s) would not answer`);
  if (scan.short) parts.push(`${scan.short} slide(s) answered with fewer shapes than they hold`);
  if (scan.tagsUnread) parts.push(`${scan.tagsUnread} shape(s) would not say whether they are charts`);
  return parts.join("; ");
}

/**
 * Which of the three answers a tag read actually gave.
 *
 * `tagValue` collapses "the host says there is no such tag" and "the host would
 * not say" into `undefined`, which is right for its callers — either way there
 * is no config to read. It is wrong for COUNTING: a shape that is genuinely not
 * a chart and a shape nobody could ask about are the same answer there, and a
 * scan that cannot tell them apart cannot report its own completeness.
 */
function tagAnswered(tag: { isNullObject: boolean; value: string }): boolean {
  return loadedValue(() => tag.isNullObject) !== undefined;
}

/** One page of a deck-wide chart scan. See `listChartsInDeck`. */
async function readChartsPage(
  from: number,
  to: number,
  withInventory: boolean,
): Promise<{
  charts: { configJson: string; target: EditTarget }[];
  unread: number;
  short: number;
  tagsUnread: number;
  inventory: SlideInventory[];
}> {
  return PowerPoint.run(async (context) => {
    const perSlide = [];
    for (let i = from; i < to; i++) {
      const slide = context.presentation.slides.getItemAt(i);
      slide.load("id");
      slide.shapes.load(withInventory ? "items/id,items/left,items/top,items/name" : "items/id,items/left,items/top");
      // The slide's own count, queued in the same sync, purely to catch a
      // collection that answers SHORT without throwing. A scalar rather than a
      // load, so it does not count against the >50-item load ceiling
      // (office-js#4272) that `READBACK_PAGE` exists for. `getCount` can throw
      // on the spot on a host that does not offer it, which costs the
      // corroboration and nothing else.
      let count: { value: number } | undefined;
      try {
        count = slide.shapes.getCount();
      } catch {
        count = undefined;
      }
      perSlide.push({ slide, count, index: i });
    }
    await step(`reading slides ${from}-${to - 1} for charts`, () =>
      withTimeout(context.sync(), READBACK_TIMEOUT_MS, `reading slides ${from}-${to - 1} for charts`),
    );

    let unread = 0;
    let short = 0;
    const lookups: ({ slideId: string; shape: PowerPoint.Shape } & ChartTags)[] = [];
    const inventory: SlideInventory[] = [];
    for (const { slide, count, index } of perSlide) {
      // A slide whose shape collection the host did not answer tells us
      // nothing, and must not be read as "no charts on this one": it is a
      // slide whose charts a deck-wide rescale would silently skip.
      const shapes = loadedItems(slide.shapes);
      const slideId = loadedValue(() => slide.id);
      if (!shapes || typeof slideId !== "string") {
        unread++;
        continue;
      }
      // Corroborate only when there is something to corroborate with. A count
      // the host would not give leaves the read exactly as trusted as it was —
      // never LESS, or every slide would look short on a host without getCount.
      const n = count ? loadedValue(() => count.value) : undefined;
      if (typeof n === "number" && shapes.length < n) short++;
      if (withInventory) {
        inventory.push({
          slideId,
          index,
          // Every property read through `loadedValue`, because a host that
          // answered the collection has not necessarily answered every property
          // on it — and a diagnostic that throws while describing the deck is
          // worse than one that reports a shape with no name.
          shapes: shapes.map((shape) => {
            const at = { id: loadedValue(() => shape.id) };
            return {
              id: typeof at.id === "string" ? at.id : "",
              name: loadedValue(() => (shape as unknown as { name: string }).name),
              left: loadedValue(() => shape.left),
              top: loadedValue(() => shape.top),
            };
          }),
          ...(typeof n === "number" ? { count: n } : {}),
        });
      }
      for (const shape of shapes) lookups.push({ slideId, shape, ...chartTagsOf(shape) });
    }
    await step(`reading chart tags on slides ${from}-${to - 1}`, () =>
      withTimeout(context.sync(), READBACK_TIMEOUT_MS, `reading chart tags on slides ${from}-${to - 1}`),
    );

    let tagsUnread = 0;
    const charts = lookups
      .map((l): { configJson: string; target: EditTarget } | undefined => {
        // Counted before the value is read, and for every shape rather than
        // only the ones that turned out to be charts — the question is how much
        // of the slide this scan could see, not how much of it was a chart.
        if (!tagAnswered(l.config)) tagsUnread++;
        const configJson = tagValue(l.config);
        const at = targetRef(l.shape);
        if (!configJson || !at) return undefined;
        return {
          configJson,
          target: {
            slideId: l.slideId,
            shapeId: at.id,
            left: at.left,
            top: at.top,
            origin: originOf(l.origin),
            partIds: partIdsOf(l.parts),
          },
        };
      })
      .filter((c) => c !== undefined);
    // Every value we need is now a plain string/number in `charts`; drop the
    // whole proxy sweep (one shape + its tags per shape, deck-wide) from memory.
    for (const l of lookups) {
      untrack(l.config);
      untrack(l.parts);
      untrack(l.origin);
      untrack(l.shape);
    }
    return { charts, unread, short, tagsUnread, inventory };
  });
}

/**
 * Find every PowerChart in the deck (any shape carrying the config tag),
 * across all slides. Used by "Same scale" to re-render charts together.
 *
 * Paged, and a page that will not read is skipped rather than fatal — the same
 * rule the repair pass already lives by (`readAddedSlides`), for the same
 * reason. This used to queue one shape-collection load per slide into a SINGLE
 * sync and then read every `.items` straight: on a 38-slide deck in PowerPoint
 * on the web that is exactly the request the host answers incompletely, and one
 * unanswered collection threw `PropertyNotLoaded` at `ShapeCollection.items`
 * out of the whole call. Same Scale then reported a crash rather than rescaling
 * the 37 charts it could see, and so did every scenario in the self-test that
 * starts by asking what is in the deck.
 *
 * What is skipped is RETURNED, not merely traced. It used to be traced only,
 * and tracing is off by default — so the one fact that makes this scan safe to
 * act on reached nobody in ordinary use, and every caller drew conclusions from
 * an array that could not tell "no charts here" from "I could not look". See
 * `DeckScan`, and `scanIsComplete` for the question every such caller owes.
 */
export async function listChartsInDeck(opts: { withInventory?: boolean } = {}): Promise<DeckScan> {
  const t0 = Date.now();
  const total = await slideCount();
  const charts: { configJson: string; target: EditTarget }[] = [];
  const inventory: SlideInventory[] = [];
  let unread = 0;
  let short = 0;
  let tagsUnread = 0;
  for (let start = 0; start < total; start += READBACK_PAGE) {
    const end = Math.min(start + READBACK_PAGE, total);
    try {
      const page = await readChartsPage(start, end, !!opts.withInventory);
      charts.push(...page.charts);
      inventory.push(...page.inventory);
      unread += page.unread;
      short += page.short;
      tagsUnread += page.tagsUnread;
    } catch {
      // A page whose sync rejected told us nothing about any slide on it.
      unread += end - start;
    }
  }
  const scan = { charts, unread, short, tagsUnread, slides: total, ...(opts.withInventory ? { inventory } : {}) };
  // ALWAYS, not only when the scan came back short.
  //
  // This used to trace on failure alone, which made every deck scan in every
  // round invisible — and the battery scans the deck a dozen times a round,
  // once per scenario through `probeCharts`. Round 10 is what that costs: `stop
  // a run part-way` reported 39.4 seconds against 2.6-3.2s in the eight rounds
  // before it, and the log had a 39-second HOLE where the scan was. The stop
  // itself was instant; the verification after it was not, and nothing said so.
  //
  // It is also the one operation the quadratic-cost finding predicts should
  // grow worst — it reads every slide's shapes, on a deck the battery is
  // steadily adding to — and it was the one operation never measured.
  //
  // A value recorded only on failures cannot be compared against anything.
  // That is the third time this session; `idleMs`, `afterAnswering` and the
  // settle's labels were the others.
  trace("pane", "scanned the deck for charts", {
    ms: Date.now() - t0,
    slides: total,
    charts: charts.length,
    unread,
    short,
    tagsUnread,
    complete: scanIsComplete(scan),
  });
  return scan;
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
 * settled, fresh-context readback — even after every retry round. This is a
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

/**
 * How many retry rounds `addSlides` gets after the initial add.
 *
 * Was 1, which assumed a dropped `slides.add()` is a one-off. The host this
 * project actually runs on does not behave that way: Presentation_3.pptx lost
 * 10 of 20 adds in a single burst, and a drop rate that high is not something
 * one retry reliably clears — the retry is issued under the same load that
 * caused the drop, so it can be dropped too, and then the run gives up with
 * slides missing and charts drawn onto the wrong ones.
 *
 * Three rounds is cheap in the only case that matters. A round costs one fresh
 * context plus one `slideCount()`, and rounds 2 and 3 only ever run on a deck
 * that is ALREADY short — a healthy insert exits the loop on the first
 * verification and pays for none of them. The bound stays small because the
 * failure it recovers from is transient by hypothesis: if three rounds under
 * three separate contexts all get dropped, the host is not dropping adds under
 * load, it is refusing them, and retrying harder will not change that.
 */
export const MAX_ADD_RETRY_ROUNDS = 3;

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
 * A deficit gets up to `MAX_ADD_RETRY_ROUNDS` retry rounds: issue the missing
 * adds in another fresh context and re-verify from a third, repeating while the
 * deck is still short. Each round re-reads the deficit, so a round that lands
 * half the missing slides only re-issues the rest. If the deficit outlives
 * every round (the drop was not transient, or the retries got dropped too),
 * give up — log via `console.warn`, add whatever is still missing to
 * `lastAddsLost`, and return thunks for only the slides that actually landed
 * (fewer than `count`).
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
  // Every add this add-in makes, traced — not only the ones that go wrong.
  //
  // The 2026-08-09 evening round left 43 slides in the owner's deck and its own
  // log accounted for four of them: the two `handed the host a generated deck`
  // pairs. Every other add came through here and said nothing unless it FAILED,
  // so "36 of these came back empty" could be measured from the deck evidence
  // and attributed to nothing. A record that only speaks up on failure cannot
  // answer "which part of the run made this slide".
  trace("host", "slides added", { requested: count, landed: have, from: start });
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
    trace("host", "slide add(s) never landed", { requested: count, lost: deficit });
    console.warn(
      `PowerChart: addSlides lost ${deficit} of ${count} requested slide${count === 1 ? "" : "s"} — ` +
        `the host dropped the add() and the retry did not recover it. Returning ${have} thunk${have === 1 ? "" : "s"}.`,
    );
  }
  // Confirm the deck really is that long FROM THIS CONTEXT before handing out
  // indices into it. `landed` comes from slideCount(), which opens a context of
  // its own, and the thunks below index the collection in THIS one — so a deck
  // the two contexts disagree about produced `getItemAt(i)` for an `i` the host
  // would not address, and answered with a bare GeneralException from inside the
  // draw. Reading the count here costs one sync against the many the render is
  // about to spend, and it is the only number the thunks are actually indexing.
  const seen = context.presentation.slides.getCount();
  await context.sync();
  const addressable = Math.max(0, seen.value - start);
  const actual = Math.min(have, count, addressable);
  if (actual < Math.min(have, count)) {
    trace("host", "slide count disagreed between contexts", {
      start,
      requested: count,
      countedGlobally: landed,
      countedHere: seen.value,
      handingOut: actual,
    });
  }
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
      // `addSlides` hands back a thunk only for an add that LANDED — its own
      // `Array.from({ length: actual })` is deliberately short of `count`,
      // because PowerPoint on the web drops `slides.add()` under load (this
      // file records a deck growing by 10 slides for 20 issued adds). Unchecked,
      // the first missing thunk dies inside the batch loop as "getSlide is not a
      // function": a TypeError from renderer internals, for a condition
      // `addSlides` diagnosed precisely one frame earlier and already logged.
      //
      // `addAndRenderItem` carries this guard with ten lines explaining it. The
      // fix was never swept to this caller, so an agenda insert of N chapters
      // landed some slides and reported "Failed: getSlide is not a function"
      // instead of the host reason.
      if (!slideThunks[i]) {
        throw new Error(
          `PowerPoint added ${slideThunks.length} of ${scenes.length} agenda slides — the host dropped the rest and every retry.`,
        );
      }
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
export const DEMO_SHAPE_BUDGET = 90;

/**
 * How long a demo item that gave up on a call waits to hear how that call
 * ended, before moving on.
 *
 * Deliberately short next to the 45s deadline that preceded it. This is not an
 * attempt to outlast the host — observed late answers arrive minutes after the
 * timeout, and no per-item wait can catch those without stalling the whole
 * deck. It catches the ones already in flight; the rest are reported by the
 * `a call we gave up on finally answered` trace whenever they land.
 */
const LATE_ANSWER_WAIT_MS = 3_000;

/**
 * Should this chart go on as a picture even though nobody asked?
 *
 * Yes on exactly one host, for exactly one reason. PowerPoint on the web has
 * no resource limits — Microsoft's CPU/memory/crash-tolerance ceilings are
 * scoped to Windows and Mac and explicitly not to a browser — so nothing
 * throttles an add-in that asks too much: the tab dies. The densest chart
 * kinds are far past what it will take as shapes (violin 253, area 176, tile
 * map 122, waffle 103), and `DEMO_SHAPE_BUDGET` is the number this renderer
 * has always used to call a chart too dense for this host. It used to skip
 * those and stamp a placeholder. A picture is strictly better than nothing —
 * it carries the config tag, so the chart stays re-editable, and Explode
 * turns it back on a host that can take it.
 *
 * Pure and exported so the rule can be tested without a DOM, a canvas or a
 * host — the three things that make the pane's own paths awkward to pin down.
 */
export function wantsAutoPicture(
  shapes: number,
  opts: { web: boolean; canPicture: boolean; alreadyPicture: boolean; budget?: number },
): boolean {
  if (opts.alreadyPicture || !opts.web || !opts.canPicture) return false;
  return shapes > (opts.budget ?? DEMO_SHAPE_BUDGET);
}

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
  box.name = NOT_COMPLETE_NAME;
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
async function stampLastSlide(title: string, detail: string, ownedFrom = 0): Promise<void> {
  await PowerPoint.run(async (context) => {
    const count = context.presentation.slides.getCount();
    await context.sync();
    // "The last slide" is only OUR slide when this item's add actually landed.
    // When the host swallowed it, the last slide belongs to whoever went
    // before — and branding it NOT COMPLETE defaces a chart that is fine. A
    // real run did exactly that: a results page whose add vanished stamped the
    // KPI tile slide, which had rendered perfectly.
    if (count.value < 1 || count.value <= ownedFrom) return;
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
  /**
   * This item gave up on a call — a deadline fired inside it. Knowable the
   * instant the item ends, and the reason this field exists separately from
   * `lateOutcome`: which item stalled is worth reading even when the host has
   * not yet said how the stall ended.
   */
  abandoned?: boolean;
  /**
   * How the abandoned sync eventually resolved, when it resolved soon enough
   * to be paired with the item that abandoned it — success, or a real RichApi
   * error. Only ever set on an `abandoned` item.
   *
   * Empty on an `abandoned` item means "the host has not answered yet", not
   * "nothing was outstanding": observed answers arrive minutes after the
   * timeout, and a run cannot wait that long between items. Those land in the
   * `a call we gave up on finally answered` trace and in the pane's late-sync
   * note instead.
   */
  lateOutcome?: string;
  /**
   * The shapes made it onto one native PowerPoint group — the state that makes
   * a chart re-editable, drags-as-one, and lands its POWERCHART_CONFIG tag.
   * False on: hosts without PowerPointApi 1.8 (grouping unsupported),
   * single-shape items (nothing to group), addGroup rejected by the host, or a
   * late-settled render where the group sync never got the chance to run.
   */
  grouped?: boolean;
  /**
   * The `POWERCHART_CONFIG` tag was written AND its sync committed — i.e. what
   * this run believes it left on the slide.
   *
   * Recorded because the settled readback's answer to the same question has
   * been observed to be WRONG. A real 38-slide run reported 20 tagged charts
   * where the produced .pptx provably carried 31, and establishing that took
   * unzipping the file. Intent and observation in the same log make that a
   * diff instead of an investigation: `tagged` true here with `tagged` false
   * in the snapshot is a readback fault, the other way round is a lost write.
   */
  tagged?: boolean;
  /**
   * Slide-adds this item actually issued: 1 normally, 2 when it made a second
   * attempt. Recorded rather than inferred from `retried`/`status`, because the
   * two do not imply each other: a too-dense item whose stamp sync is refused
   * ends "failed" having issued only ONE add (there is nothing to re-render, so
   * it never retries), and inferring a second one from the status reported a
   * slide as lost that the host had in fact kept. See `DemoReport.addsIssued`.
   */
  attempts: number;
}

/** A demo-deck insert's self-verification report. */
export interface DemoReport {
  /**
   * This run's identity — the same token written into every slot tag.
   *
   * Without it a run's report and the .pptx it produced can only be joined by
   * guessing, and a deck routinely holds slides from more than one run: the
   * file that settled the last diagnosis carried 30 slides from the run being
   * read and one from a different one. The token turns that join into a
   * lookup, and the stray into a labelled stray.
   */
  run: string;
  results: DemoResult[];
  /** How much the deck ACTUALLY grew (settled getCount, after − before). */
  slidesAdded: number;
  /**
   * How many slide-adds the run ISSUED: the sum of every item's `attempts` (one
   * per item, plus one more for each that made a second attempt — see
   * `DemoResult.attempts` for why this is counted, not inferred). Comparing this to
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
  /**
   * The end-of-run repair: what the deck actually held once the host settled,
   * and what was fixed. Present only when the caller asked for it
   * (`runOpts.reconcile`). Its verdicts, not `results`, are the honest answer
   * to "what did this run produce" — `results` is written while the host is
   * still catching up and has been observed calling a chart failed that in
   * fact landed twice.
   */
  reconcile?: ReconcileOutcome;
  /**
   * Item index at which the run stopped drawing shapes and started inserting
   * pictures, with the observation that triggered it. Undefined when the host
   * kept up all the way — which is the normal case everywhere but the web.
   */
  degradedAt?: number;
  degradeReason?: string;
  /** Wall-clock time for the whole run, ms — the headline regression metric. */
  totalMs: number;
}

/** The current slide count, read in its own settled sync (reliable on web). */
export async function slideCount(): Promise<number> {
  return PowerPoint.run(async (context) => {
    const c = context.presentation.slides.getCount();
    await boundedSync(context, "counting the deck's slides", READBACK_TIMEOUT_MS);
    // No safe default here, deliberately. Every caller uses this to measure a
    // delta — slides added, slides lost, pages to scan — and a fabricated 0
    // would read as "the deck is empty", which is an answer that gets acted on.
    // Not knowing has to stay distinguishable from knowing.
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
/**
 * Write the config tag onto a slide's single PowerChart object.
 *
 * The other half of `rescueGroupAndTag`, for the case where there is nothing
 * to group: a degraded picture, or a group whose tagging sync was dropped
 * while the shapes themselves committed. Both are ONE shape named
 * `PowerChart` carrying no `POWERCHART_CONFIG` — visibly a chart, and not
 * re-editable, with no repair that could reach them until this existed.
 *
 * Needs only PowerPointApi 1.3 (shape tags), not 1.8 — there is no grouping
 * here, so a host that cannot group can still have its charts made editable.
 */
async function retagSlideChart(slideIndex: number, tagData: string | undefined): Promise<boolean> {
  if (!supports("1.3") || !tagData) return false;
  try {
    return await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(slideIndex).shapes;
      shapes.load("items/name");
      await context.sync();
      // The chart object, never the banner: tagging the NOT COMPLETE stripe
      // would make the stripe the editable "chart".
      const target = shapes.items.find((sh) => (sh as unknown as { name?: string }).name === GROUP_NAME) as
        PowerPoint.Shape | undefined;
      if (!target) return false;
      // ONLY the config tag. The origin is deliberately not written here, and
      // takes no parameter, so it cannot be.
      //
      // It used to write `[caller.left, caller.top, shape.left, shape.top]`
      // using the repair's DEFAULT origin of (60, 90). That is not where the
      // chart was drawn — a generated deck centres its charts — so the pair
      // stopped agreeing and `origin + (live - anchor)` no longer resolved to
      // where the chart sits. On a real 38-slide run it rewrote 14 charts to
      // `[60, 90, 239.988, 120]` against a correct `[239.988, 120, 239.988,
      // 120]`, which would have teleported every one of them ~180pt left on
      // its first edit.
      //
      // A repair pass reads a settled deck; it cannot know where a chart was
      // originally drawn, so it must not claim to. Left absent or untouched,
      // an update falls back to the shape's own position — a slightly
      // different corner, not a different place on the slide.
      target.tags.add(CHART_TAG, tagData);
      try {
        await context.sync();
      } catch {
        return false;
      }
      return true;
    });
  } catch {
    return false;
  }
}

/**
 * Settle, then tag — the ordinary path's version of the repair pass's `retag`.
 *
 * `groupAndTagAll` writes the config tag inside the context that drew the
 * chart, through a shape proxy that is by then several syncs old. PowerPoint on
 * the web refuses those: `InvalidParam passed to GetItem(id)`, code 5010, at
 * `ShapeCollection.getItem` — 46 of them in one 38-item run, and the same
 * failure is what the self-test's `same scale across the deck` scenario reports
 * as "3 of 8 charts carry the shared scale; 3 still re-editable".
 *
 * The demo path already survives it, because `insertDemoDeck` re-reads the
 * settled deck afterwards and plans a `retag`. The ordinary insert and update
 * paths had no such pass, so on that host they simply shipped charts with no
 * config — visibly charts, and not re-editable, with nothing to recover them.
 * That is the whole of this function: give those two paths the recovery the
 * demo path has had all along.
 *
 * Not a retry of the write that just failed. The context that failed is gone;
 * this opens a new one and re-reads the slide, which is the same distinction
 * `groupAndTagAll`'s "no retry here on purpose" comment draws — a retry against
 * a host that just dropped a sync is what gave this project duplicate slides,
 * and a settled re-read is not one.
 *
 * The read-then-write shape is copied from `retagSlideChart` deliberately: a
 * shape handle from an `items` read, used in the batch after it, is the ONE
 * proxy pattern this host has been observed to honour — the repair pass landed
 * 23 retags that way in the same run that lost 46 tag writes.
 *
 * Best-effort, and silent about a chart it cannot identify: tagging the wrong
 * shape would make some other chart answer to this one's config.
 */
async function settleAndTagChart(slideId: string, tagData: string, shapeId?: string): Promise<boolean> {
  if (!supports("1.3") || !tagData) return false;
  // With an id, try the cheap thing first — one batch, resolve and write, no
  // collection read at all.
  //
  // Not merely fewer round trips. A shape-collection read is the one thing on
  // this host that comes back SHORT without saying so (`hollowReads` models a
  // readback that asked about 19 shapes and was told 3), and a chart missing
  // from a short answer looks exactly like a chart that is gone. So the id path
  // stays first, and the collection read is what happens when it is refused.
  //
  // The whole chain is resolved inside the batch that uses it — slide, then
  // shape, then the write — because a handle resolved by an earlier sync is
  // what this host refuses (`shape-add-held-slide-proxy`: GeneralException).
  if (shapeId) {
    let refusal: unknown;
    try {
      const wrote = await PowerPoint.run(async (context) => {
        context.presentation.slides
          .getItemOrNullObject(slideId)
          .shapes.getItemOrNullObject(shapeId)
          .tags.add(CHART_TAG, tagData);
        // Named apart from the collection-read write below, and that is the
        // whole point of the wording. Both used to say "settling the chart's
        // config tag", so a refusal in the round log could be either — and
        // round 8 turned on exactly that: one chart's settle failed with the
        // shared label and nothing said whether the by-id write had been
        // refused, or whether it had been skipped for want of an id and the
        // collection read's write refused instead. Those are different bugs
        // wanting different fixes, and the two readings cost a forensic pass
        // over Office.js statement annotations that could not settle it either.
        await boundedSync(context, "settling the config tag by shape id");
        return true;
      });
      if (wrote) return true;
    } catch (err) {
      refusal = err;
    }
    // Refused — and on the web that is the ordinary case, not the rare one.
    // Run 9 refused all five of them: `InvalidParam passed to GetItem(id)`,
    // code 5010, at `ShapeCollection.getItem`, on a slide and a shape both
    // resolved fresh inside a first sync of their own. That run's verdict was
    // "3 of 8 charts carry the shared scale ... the update reported
    // 5×no-config", which is those five refusals and nothing else.
    //
    // This used to give up here, on the grounds that "a collection search would
    // only find a DIFFERENT shape to put this chart's config on". That is true
    // with no id and false with one: the read below loads `items/id`, so the
    // caller's id picks its own shape out of the answer, and a chart that is
    // not in the answer is simply not tagged. The objection does not apply to
    // this branch, and falling through is the difference between a chart that
    // is re-editable and one that is not.
    //
    // Not a retry of the refused write, either — the same distinction the
    // header draws, and the reason the fallback runs in a context of its own.
    // The write above resolved the shape by id; the one below re-READS the
    // slide and writes through a member of that read, which is the one proxy
    // pattern this host has been observed to honour.
    trace("group", "the host refused a settle by id — re-reading the slide", {
      slideId,
      error: errorText(refusal),
    });
  }
  return settleByCollectionRead(slideId, tagData, shapeId);
}

/**
 * Find the chart by re-reading the slide's shapes, and tag the one that matches.
 *
 * With a `shapeId` this is exact: the read loads `items/id` and the caller's id
 * picks its own shape out. Without one it falls back to the reach the repair
 * pass's `retag` has — the slide must hold exactly ONE thing that calls itself
 * a chart, because with several nothing here can say which of them lost its
 * config and a guess overwrites a bystander's.
 */
/**
 * May this shape take the config the settle is carrying?
 *
 * The settle finds its shape two ways, and only one of them is proof. By ID it
 * is exactly ours — the read loads `items/id` and the caller's id picks its own
 * shape out. By NAME it is a guess, and the guard for that guess used to be
 * "the slide holds exactly one thing that calls itself a chart".
 *
 * That stops the guess when there are several. It does not stop it when the one
 * on the slide is somebody ELSE'S: our chart is ungrouped — which is why it has
 * no shape id, and why the name search cannot see it — while the group that is
 * there belongs to a chart whose own config landed perfectly well. Writing over
 * it reports a repair and hands the user a chart that opens as a different one.
 *
 * A shape that already carries a config is by definition not the shape that
 * lost one. `undefined` means the host would not say, and that is refused too:
 * an ungrouped chart carrying no config can be inserted again, while a
 * bystander carrying the WRONG config means editing one chart silently
 * rewrites another.
 *
 * Extracted because the end-to-end path resolves by id and never reaches the
 * name branch — the same reason `chooseGroupMembers` and `targetWithNoTagResult`
 * are their own functions. The rule is what is worth checking; driving four
 * simultaneous host failures through the fake to reach it tells you less.
 */
export function mayTakeConfig(o: { foundById: boolean; hasConfig?: boolean }): boolean {
  return o.foundById || o.hasConfig === false;
}

async function settleByCollectionRead(slideId: string, tagData: string, shapeId?: string): Promise<boolean> {
  try {
    return await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItemOrNullObject(slideId).shapes;
      shapes.load("items/id,name");
      await boundedSync(context, "re-reading a slide to tag the chart it would not tag");
      let items: PowerPoint.Shape[];
      try {
        items = shapes.items;
      } catch (err) {
        trace("group", "the settle's re-read would not answer", { slideId, error: errorText(err) });
        return false;
      }
      // Say it. This was a bare `return false`, and reconstructing what the
      // settle had actually done on a real host then took a forensic pass over
      // the statement shapes in a run log — because a settle that never ran, one
      // that ran and found nothing, and one that ran and was refused all end at
      // `settled=0 lost=1` with nothing between them. Four of five settles in
      // that round ended HERE, on the same empty collection read that defeats
      // the grouping, and the log did not say so once.
      if (!items?.length) {
        trace("group", "the settle's re-read came back empty", { slideId, askedById: Boolean(shapeId) });
        return false;
      }
      // By id where there is one. Note this is the ONLY way an ungrouped chart
      // can be settled: its anchor is a plain shape with an ordinary name, so
      // the name search below cannot see it at all.
      const byId = shapeId
        ? items.find((sh) => loadedValue(() => (sh as unknown as { id?: string }).id) === shapeId)
        : undefined;
      // Only ever a GROUP or a degraded picture — see above for why the name
      // search cannot reach anything else.
      const named = items.filter((sh) => loadedValue(() => (sh as unknown as { name?: string }).name) === GROUP_NAME);
      const target = byId ?? (named.length === 1 ? named[0] : undefined);
      if (!target) {
        trace("group", "no single untagged chart to settle a tag onto", {
          slideId,
          named: named.length,
          askedById: Boolean(shapeId),
        });
        return false;
      }
      // A chart found BY NAME is not provably ours, and the header above only
      // guarded half of that.
      //
      // "Exactly one thing that calls itself a chart" stops the guess when there
      // are several. It does not stop the guess when the one on the slide is
      // somebody else's: our chart is ungrouped — which is WHY it has no
      // shapeId, and why the name search cannot see it — while the group that
      // is there belongs to a chart whose own config landed perfectly well. The
      // settle would then write our config over theirs and report a repair, so
      // the user opens that chart and gets a different one's data.
      //
      // A shape that already carries a config is by definition not the shape
      // that lost one. Asked only on the name branch, and only after a
      // candidate exists, so the by-id path — the one that carries real charts
      // — pays nothing.
      if (!byId) {
        const existing = target.tags.getItemOrNullObject(CHART_TAG);
        existing.load("value");
        let hasConfig: boolean | undefined;
        try {
          await boundedSync(context, "checking the chart found by name is not already someone else's");
          hasConfig = !loadedValue(() => existing.isNullObject);
        } catch (err) {
          trace("group", "could not read the name-matched chart's tags", { slideId, error: errorText(err) });
          hasConfig = undefined;
        }
        if (!mayTakeConfig({ foundById: false, hasConfig })) {
          trace("group", "the chart found by name is not ours to tag", { slideId, hasConfig });
          return false;
        }
      }
      // Only the config tag. The origin pair is deliberately not rewritten
      // here, for the reason `retagSlideChart` spells out: this context cannot
      // know where the chart was originally drawn, and a wrong origin
      // teleports the chart on its first edit.
      target.tags.add(CHART_TAG, tagData);
      // See the by-id write's label: these two must not share a name.
      await boundedSync(context, `settling the config tag on a shape found by ${byId ? "id" : "name"}`);
      return true;
    });
  } catch {
    return false;
  }
}

/**
 * Tag every chart whose config did not land, once the drawing context is gone.
 *
 * Traced as a group rather than per chart: on a host that refuses the whole
 * batch this runs once per chart in the run, and one line each would bury the
 * run log it shares with the drawing.
 */
async function settleUntaggedCharts(
  pending: { key: string; slideId: string; tagData: string; shapeId?: string }[],
): Promise<Set<string>> {
  const settled = new Set<string>();
  if (!pending.length) return settled;
  // Keyed by the CALLER's own key, not by `shapeId`. A chart whose new shapes
  // could not be read back has no shape id to be keyed by — `shapeId` is
  // legitimately undefined there — and keying on it collapsed every such chart
  // into one bucket, so one settle landing would have cleared the `lost` marker
  // off all of them.
  for (const p of pending) if (await settleAndTagChart(p.slideId, p.tagData, p.shapeId)) settled.add(p.key);
  // Say which of the three happened, in the MESSAGE. It used to be one
  // sentence — "settled the config tag the drawing context could not write" —
  // with the numbers underneath, and the 2026-08-08 round printed it five times
  // carrying `settled: 0, lost: 1` every time. Anyone scanning messages, which
  // is how a 190-entry run log is read, saw five repairs that never happened.
  // A message that names an outcome has to be the outcome.
  //
  // All three share the `settle pass:` prefix, and that is load-bearing rather
  // than tidy. The absence of this trace is itself a diagnosis — it is
  // unconditional whenever `settleUntaggedCharts` is handed anything, so a run
  // with `tagging failed` events and no line here proves the pass was never
  // INVOKED, which is a different bug from a pass that ran and failed. That
  // reading only survives if one search still finds every form.
  const lost = pending.length - settled.size;
  trace(
    "group",
    settled.size === 0
      ? "settle pass: could not repair any config tag the drawing context lost"
      : lost === 0
        ? "settle pass: repaired every config tag the drawing context lost"
        : "settle pass: repaired SOME of the config tags the drawing context lost",
    // `withId` is what separates the settle's two routes, and the round of
    // 2026-08-11 shows why the summary needs it. Four charts in that scenario
    // hit `the settle's re-read came back empty` and were lost; a fifth was
    // repaired — and the call that repaired it was `settling the config tag on
    // a shape found by NAME`, not by id. So the discriminator there was not
    // whether the chart had been grouped (it had): it was whether the settle's
    // own fresh-context collection read answered at all. Reading that out took
    // a hand pass over `afterAnswering` strings; the count says it directly.
    { charts: pending.length, settled: settled.size, lost, withId: pending.filter((p) => p.shapeId).length },
  );
  return settled;
}

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
      shapes.load("items/name");
      await context.sync();
      // Never group the banner in with the chart. Once inside, it is invisible
      // to every later repair (a snapshot reads top-level names) and it rides
      // along with the chart forever. A real run shipped a Line chart whose
      // group held 37 shapes — 36 of them the chart, one a NOT COMPLETE stripe.
      const items = shapes.items.filter((s) => (s as unknown as { name?: string }).name !== NOT_COMPLETE_NAME);
      if (items.length < 2) return false;
      let group: PowerPoint.Shape;
      try {
        group = shapes.addGroup(items);
      } catch {
        trace("group", "the host refused addGroup", { slideIndex, shapes: items.length });
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
  /** The config tag was written and its sync committed — see DemoResult.tagged. */
  tagged: boolean;
}

async function addAndRenderItem(
  item: { scene: Scene; tagData?: string; slotTag?: string; pictureBase64?: string },
  tooDense: boolean,
  shapeCount: number,
  layout: LayoutRef,
): Promise<AddAndRenderOutcome> {
  let created = 0;
  let grouped = false;
  let tagged = false;
  await PowerPoint.run(async (context) => {
    if (!layout.resolved) {
      layout.id = await blankLayoutId(context);
      layout.resolved = true;
    }
    const [getSlide] = await addSlides(context, 1, layout.id);
    // No slide came back. `addSlides` is documented to hand out thunks only for
    // the adds that actually LANDED, so this is its contract working: every
    // add and every retry round was dropped by the host.
    //
    // Unchecked, the destructure above binds `undefined` and the next line to
    // touch it dies with "getSlide is not a function" — a TypeError from
    // renderer internals, thrown for a condition the run diagnoses precisely
    // one frame earlier and already knows how to report. That message reached
    // the user (and the run's own failure record) in place of the host's, on
    // the one path where knowing the host lost a slide is the whole story.
    //
    // Fail with the real reason instead. Everything above this call is
    // per-item recoverable: `runDemoDeck` records the item as failed and moves
    // on, and a run that loses EVERY item still rethrows as before.
    if (!getSlide) {
      throw new Error("PowerPoint did not add a slide for this chart — the host dropped the add and every retry.");
    }
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
      // Set once the run has decided to stop drawing shapes — one picture per
      // slide instead of forty shapes. See `pictureFor` on insertDemoDeck.
      pictureBase64: item.pictureBase64,
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
    // charts are the ones the web host tripped on — every ungrouped chart in
    // the real run had more shapes than its own batch could carry. "Its own",
    // because this sentence used to say `>SHAPES_PER_SYNC` and mean the live
    // canvas's ten while the draw below it used forty; see the note on
    // `needsRefresh`.
    // …or when the item is a single degraded picture. Its one proxy is created
    // in one sync and tagged in another, and the web host rejected exactly that
    // — `InvalidParam passed to GetItem(id)`, code 5010, 28 times in one run,
    // which is where that run's missing config tags went. The original
    // condition only considered multi-batch charts because grouping was the
    // only consumer at the time; tagging crosses the same boundary.
    //
    // Asked with the batch size this draw ACTUALLY used, which is what
    // `spansBatches` is for and what the other two callers already do. This one
    // hand-rolled the comparison against `SHAPES_PER_SYNC` — the LIVE canvas's
    // 10 — while drawing at `SHAPES_PER_SYNC_OFFSCREEN`'s 40. So every chart of
    // 11 to 40 shapes drew in ONE batch, with every proxy still inside its own
    // sync's window, and was then told to go and re-read the collection anyway.
    //
    // That is not a harmless extra round trip on this host. The web host does
    // not list the shapes a run has just added, so the re-read comes back
    // empty — and `chooseGroupMembers` reads an EMPTY answer to a refresh it
    // asked for as "group nothing", where not asking would have grouped the
    // perfectly good created proxies. The chart loses its group, and with it
    // the shape id the settle needs to write the config tag through: round 8's
    // finding is that the settle rescues the charts grouping survived for.
    //
    // NOT YET EXERCISED ON A REAL HOST, and nobody may credit it until it is.
    // This branch only differs from the old hand-rolled comparison for a chart
    // of 11 to 40 shapes drawn at `SHAPES_PER_SYNC_OFFSCREEN`, i.e. the deck
    // path. Round 12 (`3223293`) was the first round on a build carrying the
    // fix and every draw in it went down the LIVE path at `total:24` in batches
    // of 10 — so `spansBatches` answered true throughout, exactly as the old
    // code did, and the round says nothing about the change either way. The
    // round did come back without the empty-re-read chain, and that is NOT
    // evidence: this host produces that chain intermittently and the fix was
    // never reached. What would settle it is a round that inserts a demo deck
    // the shape-by-shape way, where charts of 11-40 shapes exist.
    //
    // SETTLED ON A REAL HOST NOW, and the answer was that the condition here was
    // still too narrow. Archive-wide, a chart that spanned batches grouped 353
    // times out of 452 and one that fitted in a single batch 49 out of 214; on
    // builds carrying the settled retry the multi arm is 10 of 10. The gate, not the
    // host, was the difference. `needsPreGroupRefresh` widens it to any chart
    // that will be grouped, and keeps `spansBatches` in the OR for the
    // ungroupable-but-stale case this branch was originally written for.
    const needsRefresh = needsPreGroupRefresh(drawn, opts, !!item.pictureBase64);
    const [result] = await groupAndTagAll(context, [{ getSlide, created: drawn, opts, refreshShapes: needsRefresh }]);
    grouped = !!result?.grouped;
    tagged = !!result?.tagged;
  });
  return { created, grouped, tagged };
}

/** One slide a demo run should produce. */
export interface DemoItem {
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
}

/** How a demo run should behave when the host starts struggling. */
export interface DemoRunOptions {
  /**
   * Read the deck back when the run finishes and repair what the host got
   * wrong — see `reconcileDeck`. Opt-in, because the results-slide inserts
   * at the end of a run are themselves `insertDemoDeck` calls and must not
   * each trigger a deck-wide sweep.
   */
  reconcile?: boolean;
  /**
   * A picture of item `i`, for when the run gives up on drawing shapes.
   *
   * Rasterizing needs a canvas, which lives in the pane, not here — so the
   * caller supplies it and this module decides WHEN to ask. Omit it and the
   * run never degrades; it just draws shapes until it cannot.
   */
  pictureFor?: (index: number) => Promise<string | undefined>;
  /**
   * Shapes this run may draw before it stops drawing them.
   *
   * The number exists because PowerPoint on the web has no resource limits
   * at all: Microsoft's documented ceilings (CPU, memory, four-crashes,
   * five-seconds-unresponsive) are scoped to Windows and Mac and explicitly
   * NOT to a browser, so nothing throttles a runaway add-in there — the tab
   * simply dies, taking the user's session with it. Twice now.
   *
   * Every 12-item run in this project's history (~400 shapes) survived;
   * every 37-item one (~1850) crashed the client. There is no published
   * number to look up, so this is our own, measured with margin — the same
   * basis as SHAPES_PER_SYNC.
   */
  shapeBudget?: number;
  /**
   * This run's identity, written into every slot tag. Defaults to a fresh
   * one; supplied only by tests that need a predictable tag.
   */
  runId?: string;
}

export async function insertDemoDeck(
  items: DemoItem[],
  onProgress?: (done: number, total: number) => void,
  runOpts: DemoRunOptions = {},
): Promise<DemoReport> {
  // Reset the module-scope lost-adds counter at the start of every run, so
  // `addsLostAtCommit` below reports only THIS run's confirmed losses, not a
  // stale accumulation from an earlier insertDemoDeck call.
  lastAddsLost = 0;
  // Stamped into every slot tag this run writes, so the settled repair pass at
  // the end can tell this run's slides from an earlier run's sitting in the
  // same deck — see `SlideSnapshot.run`.
  const runId = runOpts.runId ?? newRunId();
  // Own the host calls this run makes, so a late answer from somebody else's
  // stalled call cannot be read as this host drowning — see lastLateSyncOwner.
  // Restored rather than nulled: nothing else sets it today, but a run that
  // stomped an outer owner would recreate the very bug this fixes.
  const outerRun = activeRun;
  activeRun = runId;
  try {
    return await runDemoDeck(items, onProgress, runOpts, runId);
  } finally {
    activeRun = outerRun;
  }
}

/** The run itself. Split out only so `activeRun` has a scope to be restored in. */
async function runDemoDeck(
  items: DemoItem[],
  onProgress: ((done: number, total: number) => void) | undefined,
  runOpts: DemoRunOptions,
  runId: string,
): Promise<DemoReport> {
  const results: DemoResult[] = [];
  let lastError: unknown;
  const layout: LayoutRef = { resolved: false };
  // Bracket the run with a settled slide count, so a regression run can prove the
  // deck grew by exactly one slide per item — the lost-slide check.
  const before = await slideCount();
  const runStart = Date.now();
  // Index of the first item drawn as a picture instead of shapes, and why.
  // Undefined while the run is still healthy.
  let degradedAt: number | undefined;
  let degradeReason: string | undefined;
  let shapesDrawn = 0;
  let lostAdsSeen = 0;
  for (let i = 0; i < items.length; i++) {
    // Stop between items. A deck run is the longest thing this add-in does —
    // the one that made a cancel worth having — and an item boundary is where
    // stopping costs nothing: every slide so far is complete, grouped and
    // tagged, and the next one has not been added. The run then reports what it
    // landed, exactly as it does for any other early end, so the reconcile pass
    // still sees a truthful deck.
    if (isStopRequested()) {
      trace("draw", "deck run stopped by the user", { done: i, of: items.length });
      break;
    }
    const shapeCount = estimateOfficeShapes(items[i].scene);
    // Once the run has degraded, every remaining item goes on as a picture:
    // ONE shape, whatever the chart's native count. Asked for first, because
    // whether we have a picture is what decides the density question below.
    const degradedPicture = degradedAt !== undefined ? await runOpts.pictureFor?.(i).catch(() => undefined) : undefined;
    // The budget exists to stop a wedge/polygon flood timing the host out. A
    // picture is not a flood — it is one shape — so a run that has already
    // degraded must not still be skipping its densest charts.
    //
    // It did. A real 38-item run degraded at item 2 and then skipped and
    // stamped Area (176), Tile map (122), Waffle (103), Sunburst (101),
    // Violin (253) and Smoothed line (101): six charts drawn as nothing,
    // while the other thirty went on as one-shape pictures in ~1s each.
    // Those six are precisely the charts picture mode exists for, and they
    // were the six it refused.
    const tooDense = !items[i].bypassBudget && !degradedPicture && shapeCount > DEMO_SHAPE_BUDGET;
    let created = 0;
    let grouped = false;
    let tagged = false;
    let status: DemoResult["status"] = tooDense ? "skipped" : "rendered";
    // One add per item. There used to be a second — a retry, issued when the
    // first attempt's readback came back short — and it is where every
    // duplicate slide in this project came from: the readback ran while the
    // host was still committing, so "short" was routinely wrong and the retry
    // drew the same chart again on a new slide. The settled reconcile pass at
    // the end sees what actually landed and repairs it, which is the same job
    // done with evidence instead of a guess.
    const attempts = 1;
    const lateSeqBefore = lastLateSyncSeq;
    // Did THIS item give up on a call? The only thing about a late sync that is
    // knowable the moment an item ends — see the `abandoned` field.
    const deadlinesBefore = deadlinesFired;
    const t0 = Date.now();
    // Slot tag: JSON envelope with the deck-position index and the item title,
    // written at creation so the reconcile pass can pair a slide back to the
    // item it was drawn for however the deck ends up ordered.
    const slotTag = JSON.stringify({ i, title: items[i].title ?? null, run: runId });
    // The chart stays re-editable — the config tag rides on the picture
    // exactly as it does on a group — and "Explode to native shapes" turns it
    // back when the host is willing.
    const itemWithTag = { ...items[i], slotTag, pictureBase64: degradedPicture };
    try {
      // Which item of how many, on every line this draw writes. The grouping
      // pass already reports a meaningful `index` because it sees the whole
      // batch at once; the draw batches never did, so a deck run's log had to
      // be read by adjacency.
      ({ created, grouped, tagged } = await traceAbout({ item: `${i + 1}/${items.length}` }, () =>
        addAndRenderItem(itemWithTag, tooDense, shapeCount, layout),
      ));
    } catch (err) {
      // Do NOT infer from here what landed. A throw means the sync we were
      // waiting on gave up, not that the host discarded the work: the commits
      // observed on PowerPoint web arrive minutes after the timeout that
      // abandoned them. Record the failure, leave the slide alone, and let the
      // settled pass decide — it is the only reader that sees the end state.
      lastError = err;
      status = "failed";
      // `before` is the deck size when the run started: any slide past it was
      // added by THIS run, so stamping one cannot deface a slide the user
      // already had. An item whose add was swallowed stamps nothing.
      await stampLastSlide("NOT COMPLETE", "PowerPoint stopped responding while drawing this chart", before).catch(
        () => {},
      );
    }
    // Read before anything is awaited below: a wait for a late answer must not
    // be charged to the item as drawing time.
    const ms = Date.now() - t0;
    const abandoned = deadlinesFired !== deadlinesBefore;
    // An abandoned call reports its outcome when the host finally answers,
    // which is by definition AFTER the deadline that gave up on it. Reading
    // `lastLateSync` straight through, as this line did, therefore answered
    // whatever the clock happened to allow: the outcome of THIS item's stall if
    // the catch path ran long enough for it to land, and otherwise an EARLIER
    // item's outcome settling during this one — or a call from outside the run
    // entirely, which is the same bug `lastLateSyncOwner` fixed one level up,
    // at run granularity, and which is what the foreign-stall test caught here.
    //
    // So: give the settle a bounded chance to land, and gate on whether this
    // item abandoned anything at all. Only a stalled item waits, and it has
    // already spent BATCH_TIMEOUT_MS, next to which this is noise.
    //
    // The host often answers minutes late — far past any wait a run can afford
    // — so an empty `lateOutcome` on an `abandoned` item means "not yet", not
    // "never". The `a call we gave up on finally answered` trace and the pane's
    // late-sync note carry those, whenever they arrive.
    if (abandoned) await waitForLateSync(LATE_ANSWER_WAIT_MS);
    const lateOutcome = abandoned && lastLateSync && lastLateSyncSeq !== lateSeqBefore ? lastLateSync : undefined;
    // Is the host still keeping up? We cannot catch the crash itself — the tab
    // dies, there is no rejected promise to handle — so watch what precedes
    // it. In this project's runs the precursors are loud: healthy items land
    // in 2-9 seconds, sick ones take 65-125, and the run that killed the
    // client managed two slides in 458. Any one of these means stop drawing
    // shapes; a finished deck of pictures beats a dead session.
    shapesDrawn += created;
    if (degradedAt === undefined && runOpts.pictureFor) {
      const why =
        ms > SICK_ITEM_MS
          ? `an item took ${Math.round(ms / 1000)}s`
          : lastAddsLost > lostAdsSeen
            ? "the host lost a slide"
            : // …and only for a call THIS run issued. See lastLateSyncOwner.
              lastLateSyncSeq !== lateSeqBefore && lastLateSyncOwner === runId
              ? "the host answered after we gave up waiting"
              : shapesDrawn > (runOpts.shapeBudget ?? Infinity)
                ? `${shapesDrawn} shapes drawn`
                : undefined;
      if (why) {
        degradedAt = i + 1;
        degradeReason = why;
        console.warn(`PowerChart: drawing the rest as pictures — ${why}`);
        trace("demo", "stopped drawing shapes", { fromItem: i + 1, why, shapesDrawn });
      }
    }
    trace("demo", "item finished", {
      i,
      title: items[i].title,
      status,
      created,
      expected: shapeCount,
      grouped,
      // What this run believes it left on the slide, to compare against what
      // the settled readback reports for the same item — see DemoResult.tagged.
      tagged,
      picture: !!degradedPicture,
      attempts,
      ms,
      abandoned,
    });
    lostAdsSeen = lastAddsLost;
    results.push({ created, status, ms, lateOutcome, abandoned, grouped, tagged, attempts });
    onProgress?.(i + 1, items.length);
  }
  const totalMs = Date.now() - runStart;
  const after = await slideCount();
  const slidesAdded = after - before;
  // The deck's size at each boundary. A run once ended with a stamped stray
  // slide PAST this range that nothing owned or cleaned, and finding it took
  // unzipping the .pptx — a count here would have shown it.
  trace("demo", "run finished drawing", {
    before,
    after,
    slidesAdded,
    items: items.length,
    degradedAt,
    totalMs,
  });
  // The adds we actually ISSUED, summed per item rather than inferred from
  // retried/failed — a too-dense item whose stamp sync is refused ends "failed"
  // on ONE add, and inferring a second reported a phantom lost slide. Loss is
  // `addsIssued − slidesAdded`, never `items.length − slidesAdded`: a stray from
  // a retry/fail cancels a lost slide against the latter, hiding corruption
  // (observed on a real run: 2 slides lost, reported 0).
  const addsIssued = results.reduce((n, r) => n + r.attempts, 0);
  // A whole deck lost to HOST errors (not just skipped-as-dense) is a real
  // failure — surface it so the pane says "Failed", not "inserted 0 of N". If
  // everything was merely skipped, there is no error to throw.
  if (items.length && results.every((r) => r.status !== "rendered") && lastError !== undefined) {
    throw lastError instanceof Error ? lastError : new Error(String(lastError));
  }
  // Repair pass. Runs only now, because only now has the host stopped moving:
  // every inline check above races a commit that can land minutes after the
  // timeout that gave up on it, which is how one run shipped two Line charts,
  // two Gantt charts, four false NOT COMPLETE banners and a "BLANK" Agenda
  // slide holding all 13 of its shapes. See `src/core/reconcile.ts`.
  let reconcile: ReconcileOutcome | undefined;
  if (runOpts.reconcile) {
    const expected: ExpectedItem[] = items.map((it, i) => ({
      slot: i,
      title: it.title ?? `Item ${i + 1}`,
      // A DEGRADED item is one picture, not the scene's shape count.
      //
      // `estimateOfficeShapes(it.scene)` is what the chart would be as native
      // shapes, and for everything up to `degradedAt` that is exactly right.
      // Past it the run drew a single picture on purpose — so expecting 253
      // shapes on a Violin slide asks the reconcile to compare a healthy chart
      // against a shape count nothing ever tried to put there, and the honest
      // verdict for that slide is `wreckage`.
      //
      // It has never been SEEN as wreckage, and that is the uncomfortable part:
      // `unmeasured` short-circuits the comparison on this host, so both bugs
      // were invisible because they cancelled. A run on 2026-08-08 drew 36 of
      // its 38 slides as pictures and reported "35 of 38 complete" with per-slide
      // counts up to 253 — numbers describing a chart that was not on the slide.
      shapes: degradedAt !== undefined && i >= degradedAt ? 1 : estimateOfficeShapes(it.scene),
      chart: !!it.tagData,
      skipped: results[i]?.status === "skipped",
      // What this run watched happen — see ExpectedItem.wroteTag.
      wroteTag: results[i]?.tagged,
    }));
    reconcile = await reconcileDeck(
      expected,
      { before, after },
      (slot) => items[slot]?.tagData,
      // Every slide in [before, after) was added by THIS run, so an empty one
      // with no slot tag is our own litter and safe to sweep. The run token
      // rides along regardless: the range is trustworthy here, but a slide the
      // host filed outside it is not, and `runId` is what says so.
      { dropOrphanBlanks: true, run: runId },
    ).catch(() => undefined);
  }
  // Blank readback: find added slides the host kept but left EMPTY. Counted by
  // position, not mapped to items — a scrambled deck (lost/merged/reordered
  // slides, plus retry strays) breaks any positional item mapping, and a blank
  // slide has no tag to identify it anyway. Best-effort: a fault leaves
  // blanksRead false so an empty list is not read as "no blanks".
  //
  // Skipped when the repair pass ran: its snapshot answers the same question
  // from the same settled read, and answers it better — it knows which ITEM a
  // blank slide belonged to, and it has already deleted the empty strays this
  // check would report. Re-running it here would name slides that no longer
  // exist.
  const blanks = reconcile ? blanksFromSnapshots(reconcile) : await findBlankAddedSlides(before, after);
  return {
    run: runId,
    results,
    slidesAdded,
    addsIssued,
    addsLostAtCommit: lastAddsLost,
    blankSlides: blanks.positions,
    blankItems: blanks.items,
    blanksRead: blanks.complete,
    reconcile,
    degradedAt,
    degradeReason,
    totalMs,
  };
}

/**
 * The blank-slide report, taken from the repair pass's own snapshot instead of
 * a second readback: a slide with no content shapes, named by the item its
 * slot tag claims. `complete` is true because the snapshot either read a page
 * or skipped it entirely — a page it could not read contributes no slides, and
 * therefore no false "clean" claim about them.
 */
function blanksFromSnapshots(outcome: ReconcileOutcome): {
  positions: number[];
  items: { position: number; title: string | null }[];
  complete: boolean;
} {
  const deleted = new Set(outcome.plan.actions.filter((a) => a.kind === "delete").map((a) => a.index));
  const blank = outcome.snapshots.filter((s) => !deleted.has(s.index) && s.shapes === 0);
  return {
    positions: blank.map((s) => s.index + 1),
    items: blank.map((s) => ({ position: s.index + 1, title: s.title })),
    complete: true,
  };
}

/** Slides read per sync in the readback — kept modest to stay clear of the web
 *  >50-item load ceiling (office-js#4272), though getCount is a scalar, not a load. */
export const READBACK_PAGE = 20;

/**
 * How long one repair-pass page may take before the run gives up on it.
 *
 * The drawing phase has had a per-batch bound since the first stalled host;
 * everything AFTER it — grouping, tagging, the readback, the repair — had none
 * at all. 75 of this file's 79 syncs were unbounded, and the four that were not
 * are all in the insert path. A run that got past drawing could therefore sit
 * forever on a single unanswered sync with nothing to break the wait: observed,
 * at 1819 seconds and climbing, on a deck run that had to be abandoned by
 * closing the tab.
 *
 * Generous, and deliberately larger than `BATCH_TIMEOUT_MS`: a page here reads
 * twenty slides and their tags, where a draw batch is ten shapes. Its only job
 * is to stop an infinite wait, not to police a slow host — and every caller
 * already has an honest answer for a page it could not read ("unread",
 * "undetermined", "unmeasured"), which is why abandoning one is safe.
 */
let READBACK_TIMEOUT_MS = 90_000;

/** Test-only: shorten the repair-pass page budget so a stall is testable. */
export function _setReadbackTimeoutForTest(ms: number): void {
  READBACK_TIMEOUT_MS = ms;
}

/**
 * The current default budget, for callers that want a SHORTER one.
 *
 * A caller with its own budget must not out-wait this one — a hard-coded ten
 * seconds is shorter than ninety in production and much longer than the
 * milliseconds a test shortens this to, which would make a bounded wait
 * untestable at exactly the site that needed bounding.
 */
export const readbackTimeoutMs = (): number => READBACK_TIMEOUT_MS;

/**
 * How long to wait for the host to draw a slide, which is much less than how
 * long to wait for it to read one.
 *
 * A rasterise either answers quickly or does not answer. There is no middle:
 * every successful `getImageAsBase64` this project has recorded came back in
 * about a second, and PowerPoint on the web has now failed the same call three
 * different ways across three rounds — `GeneralException` at
 * `SlideCollection.getItem`, then taking the call and silently producing
 * nothing, then (2026-08-06) never answering the sync at all.
 *
 * That last one cost a whole round. It burned the full ninety-second readback
 * budget and the tab died moments later on the delete that followed, taking the
 * run's own report with it — for a scenario whose honest verdict is one word,
 * `skipped`. Waiting ninety seconds buys nothing here and costs everything
 * after it.
 *
 * Capped by the readback budget rather than replacing it, for the reason
 * `readbackTimeoutMs` gives: a fixed number would be longer than the
 * milliseconds a test shortens the budget to, and the wait would stop being
 * testable at the one site that most needed bounding.
 */
export const rasteriseTimeoutMs = (): number => Math.min(20_000, READBACK_TIMEOUT_MS);

/**
 * Every PowerPointApi set this host admits to.
 *
 * Half of what the add-in does is gated on these, so a run log or an answer
 * sheet that does not carry them cannot be read: the same verdict means
 * different things on 1.4 and on 1.10.
 */
export function requirementSets(): string[] {
  return ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"].filter((v) => supports(v));
}

/**
 * What this tab is, and how long it has been one.
 *
 * Recorded once per round because half the hypotheses this project has
 * entertained are about the ENVIRONMENT — the tab's age most of all, which has
 * been alive as a candidate for ten rounds with nothing measuring it. Every one
 * of these is free (no host round trip) and constant for the round, so it costs
 * one trace line and can never evict anything that matters.
 *
 * Everything is optional because every one of these APIs is absent somewhere:
 * `Office.context.diagnostics` predates nothing but is not in the fake,
 * `performance.memory` is Chromium-only, and a missing value must read as
 * "not available" rather than as a number.
 */
export interface RoundEnvironment {
  /** Host, platform and version as Office itself reports them. */
  host?: string;
  platform?: string;
  officeVersion?: string;
  /** Requirement sets this host admits to. */
  requirementSets: string[];
  /**
   * Milliseconds since this DOCUMENT was created — the tab's age at round
   * start, which is the number "does a long-lived session degrade" needs and
   * has never had. `performance.timeOrigin` is when the page began; a round
   * that starts twenty minutes in says so.
   */
  tabAgeMs?: number;
  /** JS heap in MB, when the browser exposes it. Chromium does; others do not. */
  heapUsedMb?: number;
  heapLimitMb?: number;
  /** Viewport, because the live canvas is the bottleneck and its size is part of that. */
  viewport?: string;
}

/** Read the round's environment. No host round trip — every source is local. */
export function roundEnvironment(): RoundEnvironment {
  // EVERY read here is in a try, including the requirement sets. This function
  // is evaluated as an argument to `trace("round starting", …)`, so anything it
  // throws does not degrade the environment block — it deletes the round-start
  // line entirely, which is the one line that says which deck a round that dies
  // died on. Found by `pane-host-actions.test.ts` on the first run, in exactly
  // that form: eight tests red because a diagnostic threw where no diagnostic
  // had been before.
  let sets: string[] = [];
  try {
    sets = requirementSets();
  } catch {
    /* a host that will not enumerate its API surface still has a round to log */
  }
  const env: RoundEnvironment = { requirementSets: sets };
  try {
    const d = (Office as unknown as { context?: { diagnostics?: Record<string, unknown> } })?.context?.diagnostics;
    if (d) {
      if (typeof d.host === "string") env.host = d.host;
      if (typeof d.platform === "string") env.platform = d.platform;
      if (typeof d.version === "string") env.officeVersion = d.version;
    }
  } catch {
    /* diagnostics is not everywhere, and a missing environment must not cost a round */
  }
  try {
    const perf = globalThis.performance as
      (Performance & { memory?: { usedJSHeapSize: number; jsHeapSizeLimit: number } }) | undefined;
    if (perf && typeof perf.now === "function") env.tabAgeMs = Math.round(perf.now());
    const mem = perf?.memory;
    if (mem) {
      env.heapUsedMb = Math.round(mem.usedJSHeapSize / 1048576);
      env.heapLimitMb = Math.round(mem.jsHeapSizeLimit / 1048576);
    }
  } catch {
    /* same */
  }
  try {
    const w = globalThis as unknown as { innerWidth?: number; innerHeight?: number };
    if (w.innerWidth && w.innerHeight) env.viewport = `${w.innerWidth}x${w.innerHeight}`;
  } catch {
    /* same */
  }
  return env;
}

/** What a host probe is handed: a live context, and a slide it may wreck. */
export interface ProbeContext {
  /** The deck's slide collection, for questions about lookup and counting. */
  slides: PowerPoint.SlideCollection;
  /**
   * The presentation itself, for the few questions that are not about slides.
   *
   * Handed over explicitly rather than reached for through `slides.context`.
   * That property is part of the real proxy contract and the fake does not
   * model it, so a probe that used it answered `unreadable` in CI while
   * claiming to be about slide masters — a fake divergence invented by the
   * question's own plumbing, which is the one thing this file must never do.
   */
  presentation: PowerPoint.Presentation;
  /**
   * The scratch slide, RE-ACQUIRED on every call — never a proxy to hold.
   *
   * A thunk for the same reason `SlideThunk` is one, and this file had already
   * written down why a thousand lines above: the scratch slide was `add()`ed
   * moments ago, and a handle on a freshly-added slide is only good inside the
   * sync that acquired it. The first version of this handed out the proxy
   * resolved by the liveness check below, which every probe then used AFTER
   * that sync.
   *
   * PowerPoint on the web said what that costs. In the 2026-08-04 answer sheet
   * all six questions that read through a proxy of their own were answered, and
   * all eight that wrote through the held one failed — `GeneralException` on the
   * shape add, or a sync that never came back. Eight questions the sheet then
   * reported as host divergences, none of which had actually been asked.
   */
  scratch: () => PowerPoint.Slide;
  scratchId: string;
  /**
   * A slide that was in the deck BEFORE this run — the control for every
   * question about freshness.
   *
   * Undefined on a deck the run has entirely built, which is a real state and
   * not an error: a probe that needs it reports that it had none rather than
   * silently answering about a slide it added. Nothing may WRITE to it. It is
   * someone's own presentation, and a diagnostic that draws on it is one they
   * stop clicking.
   */
  durableSlideId?: string;
  sync: () => Promise<void>;
}

/**
 * Thrown when a probe never got as far as its question, with WHICH of the two
 * ways that happened.
 *
 * The distinction is the whole finding, and the first version threw it away. A
 * real PowerPoint on the web answered thirteen of fourteen questions with the
 * same sentence — "the host would not resolve the scratch slide" — which is two
 * completely different diagnoses wearing one string:
 *
 * - `gone`: the host says the slide is not there. Something took it away, or it
 *   never durably existed under that id.
 * - `silent`: the host will not say either way. `isLive` reads `isNullObject`,
 *   which `queueNullCheck` populates by loading a REAL property — so this is
 *   also the symptom of that trick not working on this host, and if it does not
 *   work, every `isLive` guard in this file is answering "not live" about live
 *   objects and quietly refusing to act on them.
 *
 * One word separates "your diagnostic littered the deck" from "the add-in's
 * central existence check is broken here". It is worth carrying.
 */
export class ScratchSlideUnavailable extends Error {
  constructor(readonly why: "gone" | "silent") {
    super(
      why === "gone"
        ? "the host says the scratch slide is gone"
        : "the host would not say whether the scratch slide exists",
    );
    this.name = "ScratchSlideUnavailable";
  }
}

/**
 * Run one host probe against a scratch slide, bounded.
 *
 * Its own context per probe, deliberately: a question about proxy staleness or
 * a wedged sync must not carry its damage into the next question, and the
 * sheet's whole value is that it comes back complete from a host that is
 * misbehaving. Bounded for the same reason — one silent question costs a
 * budget, not the run.
 */
export async function withProbeContext<T>(
  scratchId: string,
  budgetMs: number,
  fn: (ctx: ProbeContext) => Promise<T>,
  durableSlideId?: string,
  /**
   * Skip the liveness check for a question that never touches the slide.
   *
   * The check is the single most likely thing to fail on this host — a fresh
   * slide's id resolves once and then stops — so charging it to a question that
   * does not need a slide turns a real answer into `no-scratch-slide`. See
   * `Probe.noSlideNeeded` for the round that lost one that way.
   */
  slideless = false,
): Promise<T> {
  return boundedRun(
    "asking the host a probe question",
    async (context) => {
      if (!slideless) {
        const scratch = context.presentation.slides.getItemOrNullObject(scratchId);
        queueNullCheck(scratch);
        await context.sync();
        const flag = loadedValue(() => scratch.isNullObject);
        if (flag !== false) throw new ScratchSlideUnavailable(flag === true ? "gone" : "silent");
      }
      // One handle per sync-batch, never the `scratch` proxy the liveness check
      // above already holds — that one is a sync old by the time a probe gets
      // here, which is the whole trap. Per BATCH rather than per call because
      // that is the actual rule: a handle is good inside the sync that acquired
      // it, and several uses within one batch are what production does. Per
      // call would also spend host lookups a real host has been seen to ration.
      let handle: PowerPoint.Slide | null = null;
      return fn({
        slides: context.presentation.slides,
        presentation: context.presentation,
        scratch: () => (handle ??= context.presentation.slides.getItemOrNullObject(scratchId)),
        scratchId,
        durableSlideId,
        sync: async () => {
          await context.sync();
          handle = null;
        },
      });
    },
    budgetMs,
  );
}

/** One round of shape adds, and what it cost. */
export interface ShapeRound {
  /** 1-based, in the order the rounds were issued. */
  round: number;
  /** Wall-clock for this round's adds plus the sync that committed them. */
  ms: number;
  /** How many shapes the round queued. Constant by construction — carried so a reader need not assume. */
  shapes: number;
}

/** What a timed series came back with, and why it stopped if it stopped early. */
export interface ShapeRoundSeries {
  rounds: ShapeRound[];
  /** Set when fewer rounds came back than were asked for. */
  cutShort?: string;
}

/**
 * The measurement grid `timeShapeRounds` lays its rectangles out on.
 *
 * Exported so the caller that hands it an origin can check the footprint fits
 * where it is putting it, rather than the two agreeing by coincidence.
 */
export const GRID_COLS = 12;
export const GRID_PITCH = 10;
export const GRID_CELL = 8;

/** How much room `n` shapes need at this pitch. */
export function gridFootprint(n: number): { width: number; height: number } {
  const cols = Math.min(n, GRID_COLS);
  const rows = Math.ceil(n / GRID_COLS);
  return {
    width: Math.max(0, (cols - 1) * GRID_PITCH + GRID_CELL),
    height: Math.max(0, (rows - 1) * GRID_PITCH + GRID_CELL),
  };
}

/**
 * Add shapes to one slide in N timed rounds — either all inside ONE request
 * context, or one context per round.
 *
 * This exists to answer a question three sessions of real-host artefacts could
 * not settle by reasoning: a long run gets slower and eventually wedges, and at
 * least three things could be causing it — the request context accumulating
 * proxies, the DECK growing, or the tab simply having been alive a while. Every
 * artefact this project owns is consistent with all three, because every one of
 * them varies all three at once.
 *
 * Two series taken the same way, differing in exactly one thing, separate them.
 * The comparison that matters is each series' own SLOPE, not the two levels:
 * whichever arm runs second faces a bigger deck and an older tab, so the levels
 * are not comparable and were never meant to be. Within an arm both confounds
 * grow the same way, so a slope difference between the arms is attributable to
 * the one thing that differs — how long a context lives.
 *
 * Deliberately plain rectangles rather than charts. The measurement should be
 * "queue k shapes, sync", the operation every degraded run was doing when it
 * degraded, with no fills, no groups and no tags to make the two arms differ in
 * anything but context lifetime. Each shape is named, in both arms equally, so
 * the deck the experiment leaves behind can be read afterwards.
 *
 * Partial results are the point: a series that wedges at round five returns the
 * four rounds it has and says so. A curve with four points is a finding; an
 * exception is not.
 */
export async function timeShapeRounds(
  slideId: string,
  opts: {
    rounds: number;
    perRound: number;
    oneContext: boolean;
    label: string;
    budgetMs?: number;
    origin?: { left: number; top: number };
  },
): Promise<ShapeRoundSeries> {
  const { rounds, perRound, oneContext, label } = opts;
  const budgetMs = opts.budgetMs ?? READBACK_TIMEOUT_MS;
  // Bottom-left of a 960x540 deck when the caller does not say, which is what
  // this drew before the slot came in. Callers on a real deck pass a slot.
  const origin = opts.origin ?? { left: 20, top: 400 };
  const out: ShapeRound[] = [];
  let cutShort: string | undefined;
  /**
   * Queue one round's shapes.
   *
   * The slide is re-resolved every round even inside the long-lived context,
   * because that is what production does now and what the host demands: only an
   * id crosses a sync, never a handle. So the long arm is not a strawman — it is
   * today's code with its context left open, which is precisely the difference
   * under test.
   */
  const queueRound = (context: PowerPoint.RequestContext, round: number) => {
    const slide = context.presentation.slides.getItemOrNullObject(slideId);
    for (let i = 0; i < perRound; i++) {
      const n = (round - 1) * perRound + i;
      // Small, and out of the way. This used to grid from (20,20) — fine when
      // the experiment took a scratch slide of its own, and not fine since it
      // stopped: it now draws onto a slide the run already owns, and ninety-six
      // squares from the top-left corner sit squarely on that chart's title.
      // The owner opens these decks, so a measurement artefact has to look like
      // one and stay out of the way of what it is measuring beside.
      //
      // 8pt cells on a 10pt pitch: GRID_COLS x GRID_COLS of them fit in
      // 120x120. WHERE that footprint goes is the caller's to decide and is
      // passed in — a hardcoded corner was a guess about the slide, and a guess
      // about the slide is what put a fixed column across a 4:3 chart.
      const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: origin.left + (n % GRID_COLS) * GRID_PITCH,
        top: origin.top + Math.floor(n / GRID_COLS) * GRID_PITCH,
        width: GRID_CELL,
        height: GRID_CELL,
      });
      shape.name = `${label} r${round} #${i}`;
    }
  };
  if (oneContext) {
    try {
      // One deadline for the whole series, because the series IS one context —
      // splitting it would end the context, which is the variable. Rounds land
      // in `out` as they complete, so a deadline that fires still leaves every
      // round that finished before it.
      await boundedRun(
        `timing ${rounds} shape rounds in one context`,
        async (context) => {
          for (let round = 1; round <= rounds; round++) {
            if (isStopRequested()) {
              cutShort = "the run was stopped";
              return;
            }
            const t0 = Date.now();
            queueRound(context, round);
            await context.sync();
            out.push({ round, ms: Date.now() - t0, shapes: perRound });
          }
        },
        budgetMs * rounds,
      );
    } catch (err) {
      cutShort = errorText(err);
    }
  } else {
    for (let round = 1; round <= rounds; round++) {
      if (isStopRequested()) {
        cutShort = "the run was stopped";
        break;
      }
      const t0 = Date.now();
      try {
        await boundedRun(
          "timing one shape round in its own context",
          async (context) => {
            queueRound(context, round);
            await context.sync();
          },
          budgetMs,
        );
      } catch (err) {
        cutShort = errorText(err);
        break;
      }
      out.push({ round, ms: Date.now() - t0, shapes: perRound });
    }
  }
  if (!cutShort && out.length < rounds) cutShort = `only ${out.length} of ${rounds} rounds came back`;
  trace("host", "timed shape rounds", { label, oneContext, got: out.length, of: rounds, cutShort });
  return cutShort ? { rounds: out, cutShort } : { rounds: out };
}

/**
 * `PowerPoint.run`, with a deadline and a name.
 *
 * Same shape as the call it replaces, so a site adopts it by changing the
 * function name and nothing else. The label is what the run log will carry when
 * the deadline fires, and what `lastLateSync` will credit if the host answers
 * afterwards.
 */
function boundedRun<T>(
  what: string,
  fn: (context: PowerPoint.RequestContext) => Promise<T>,
  budgetMs = READBACK_TIMEOUT_MS,
): Promise<T> {
  return withTimeout(PowerPoint.run(fn), budgetMs, what);
}

/**
 * The message `withTimeout` rejects with, recognised by callers that need to
 * tell "the host refused" from "the host never answered".
 *
 * A refusal is a fact about the request; silence is a fact about the host, and
 * only the second one is a reason to report a known platform limitation rather
 * than a defect of ours.
 */
export function isTimeout(err: unknown): boolean {
  return err instanceof Error && /did not respond while/.test(err.message);
}

/**
 * A bounded MUTATION whose result is decided by re-reading the document, not by
 * the promise.
 *
 * On this platform a `context.sync()` that never resolves does not mean the
 * work did not happen. office-js#1650 is exactly that, on exactly the call this
 * add-in leans on hardest: *"the first time `context.sync()` is called the
 * promise resolves, but in subsequent calls the promise doesn't resolve,
 * although **the slide still gets added successfully**."* The same shape is
 * reported for shape work after an image insert (office-js#5022).
 *
 * Every timeout this file had treated silence as failure and threw. For a read
 * that is right — an unread page is unread. For a WRITE it is wrong twice over:
 * it discards work that landed, and it sends the caller off to do the work
 * again, which is how one stalled insert becomes two copies of a chart.
 *
 * So a bounded mutation swallows its own silence and lets the caller's
 * verification decide. Every site that uses this already measures what actually
 * happened — `slideCount()` before and after, a fresh-context re-read — it just
 * never used to reach the measurement. A refusal still throws: the host said no
 * and is still talking, and that is a different fact.
 */
async function withTimeoutOrVerify<T>(p: Promise<T>, ms: number, what: string): Promise<void> {
  try {
    await withTimeout(p, ms, what);
  } catch (err) {
    if (!isTimeout(err)) throw err;
    trace("host", "no answer — asking the document what actually happened", { what, waitedMs: ms });
  }
}

/** Top-level shape counts for slides [start, end), read in one settled sync. */
async function shapeCounts(start: number, end: number): Promise<number[]> {
  return boundedRun(`counting shapes on slides ${start}-${end - 1}`, async (context) => {
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
 * How long to leave the host alone after a structural change to the deck.
 *
 * The only workaround anyone has for office-js#5022 — `context.sync()` running
 * forever after add → delete → re-read — is a pause. The reporter's own words:
 * *"I had better result by adding a timer of 1-2 seconds between the
 * `shape.delete()` and the next `await context.sync()`."* Microsoft has it under
 * investigation and has offered nothing else.
 *
 * A second is a real cost and it is not a guess dressed as one: it is the
 * cheapest known answer to a failure that costs a whole operation, and it is
 * spent only where two structural changes meet back to back.
 */
let SETTLE_MS = 1_000;

/** Test-only: a suite cannot spend a real second per structural change. */
export function _setSettleDelayForTest(ms: number): void {
  SETTLE_MS = ms;
}

/** Leave the host alone for a moment. See `SETTLE_MS`. */
function settle(): Promise<void> {
  return SETTLE_MS > 0 ? new Promise((r) => setTimeout(r, SETTLE_MS)) : Promise.resolve();
}

/**
 * How long to wait before asking a slide for its shapes a SECOND time, when the
 * first answer was short or empty.
 *
 * This is the same remedy as `SETTLE_MS` and `BLANK_REREAD_DELAY_MS` applied to
 * the one read that had never had it — and it is the first thing here that came
 * out of the office-js tracker rather than out of a round.
 *
 * **What the tracker says.** PowerPoint Online does not populate a slide's
 * shapes collection immediately after the slide is added; the workaround
 * everyone converges on is to wait a second or two before reading it
 * (office-js#2903, echoed in #5022 — the same issue `SETTLE_MS` above already
 * quotes for a different symptom).
 *
 * **Why it was not applied here sooner.** This repo had read #2903 and filed it
 * as "upstream has nothing, `sleep(2000)` only". That was a fair reading while
 * the failure looked like a TAG problem. It is the wrong reading now: the
 * failure has been isolated to exactly the state the workaround addresses.
 * Measured over the whole round archive on 2026-08-15:
 *
 *     slide already had shapes   82 chart(s), 81 grouped = 99%
 *     freshly added, empty       74 chart(s),  1 grouped =  1%
 *
 * A chart drawn onto a slide this run has just added gets a pre-grouping re-read
 * that comes back short or empty, so it is not grouped, so its tag falls back to
 * a `created` handle — which this host refuses about seven times in ten.
 *
 * **1500ms, and the cost is bounded and rare.** It is paid per PASS and not per
 * chart, because the whole batch re-reads in one sync — so a deck-wide update
 * that loses five charts waits once, not five times. It is paid ONLY when the
 * first read came back short or empty; a complete read costs nothing at all. And
 * when the second read is short too, everything falls through to the behaviour
 * that was there before, so the worst case is today plus a second and a half.
 */
let REREAD_RETRY_MS = 1_500;

/** Test-only: a suite cannot spend a real 1.5s per short re-read. */
export function _setReReadRetryDelayForTest(ms: number): void {
  REREAD_RETRY_MS = ms;
}

/**
 * How many EXTRA attempts the pre-grouping re-read gets. One.
 *
 * Deliberately not tunable and deliberately not more. A second ask tests the
 * tracker's claim — that the collection fills in shortly after the slide is
 * materialised — and a third would only be waiting on a host that has already
 * answered the same way twice, at a second and a half a go, on the path a
 * deck-wide update runs for every chart.
 */
const REREAD_ATTEMPTS = 1;

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

/** Name `groupAndTagAll` and `rescueGroupAndTag` give the chart's group shape. */
const GROUP_NAME = "PowerChart";

/**
 * Read the added range back as `SlideSnapshot`s — the settled truth a repair
 * pass reasons over.
 *
 * Three passes, each independently fault-tolerant, because a single sync that
 * touches every shape's tags AND every group's children on a struggling host
 * is exactly the kind of call this whole module exists to survive:
 *
 *  A. names + slot tag per slide — gives the shape count, the banner, whether
 *     a group is present, and which item the slide belongs to.
 *  B. the config tag, queued for every shape on slides that have any. Same
 *     shape of sweep `listChartsInDeck` already does deck-wide.
 *  C. the group's child count, one context per grouped slide, since reaching
 *     `Shape.group` on a host below PowerPointApi 1.8 poisons its whole sync.
 *
 * A pass that faults leaves its field unset rather than guessing: an unread
 * group reports `grouped` with no `groupChildren`, which `planReconcile` reads
 * as "do not touch this slide".
 */
export async function snapshotAddedSlides(before: number, after: number): Promise<SlideSnapshot[]> {
  return (await readAddedSlides(before, after)).snapshots;
}

/**
 * The same read, plus how many slides it could NOT see.
 *
 * A page whose read throws is skipped, which is the safe thing to do — an
 * unseen slide is never deleted. But its slides then simply are not in the
 * result, and an item that lives on one comes back `lost`: indistinguishable,
 * in the run summary, from a slide the host really did drop. "Gantt: lost"
 * when Gantt rendered perfectly is exactly the report that makes someone run
 * the insert again. `DemoReport.blanksRead` already draws this distinction for
 * the blank check; this is the same honesty for the repair pass.
 */
export async function readAddedSlides(
  before: number,
  after: number,
  /**
   * Run passes B and C here, over everything pass A read.
   *
   * False lets a caller locate the span FIRST and enrich only that. Pass A has
   * to stay deck-wide — the span is discoverable only by reading slot tags, and
   * assuming this run's slides sit at the tail is a bug this project has already
   * shipped once (`insertSlidesFromBase64` put them at the FRONT on the first
   * real run). Passes B and C are the expensive ones and none of their answers
   * outside the span can ever be acted on: on a 39-slide deck holding earlier
   * work, that is a tag read and two group-count syncs per slide, spent on
   * charts the repair will never touch.
   */
  enrich = true,
): Promise<{ snapshots: SlideSnapshot[]; unread: number }> {
  const snapshots: SlideSnapshot[] = [];
  let unread = 0;
  for (let start = before; start < after; start += READBACK_PAGE) {
    const end = Math.min(start + READBACK_PAGE, after);
    // A stop ends the readback here, and the pages not reached are counted as
    // UNREAD rather than as clean. That distinction is the whole point: an
    // unseen slide is never deleted by the pass that follows, so a stopped
    // read degrades to "we did not look" instead of "there was nothing there".
    if (isStopRequested()) {
      unread += after - start;
      trace("repair", "readback stopped by the user", { from: start, to: after });
      break;
    }
    let page: SlideSnapshot[];
    try {
      page = await boundedRun(`reading slides ${start}-${end - 1} back`, async (context) => {
        const reads = [];
        for (let i = start; i < end; i++) {
          const slide = context.presentation.slides.getItemAt(i);
          slide.shapes.load("items/name");
          const tags = (
            slide as unknown as {
              tags: { getItemOrNullObject(k: string): { isNullObject: boolean; value: string; load(): void } };
            }
          ).tags;
          let tag: { isNullObject: boolean; value: string; load(): void } | null = null;
          try {
            tag = tags.getItemOrNullObject(DEMO_SLOT_TAG);
            tag.load();
          } catch {
            /* no slide tags on this host — the slide stays unclaimed */
          }
          reads.push({ index: i, shapes: slide.shapes, tag });
        }
        await context.sync();
        return reads.map((r) => {
          const names = r.shapes.items.map((s) => (s as unknown as { name: string }).name);
          const slot = r.tag && !r.tag.isNullObject ? parseSlotTag(r.tag.value) : null;
          return {
            index: r.index,
            slot: slot?.i ?? null,
            title: slot?.title ?? null,
            run: slot?.run ?? null,
            shapes: names.length,
            stamped: names.includes(NOT_COMPLETE_NAME),
            grouped: names.includes(GROUP_NAME) || undefined,
            tagged: false,
          } satisfies SlideSnapshot;
        });
      });
    } catch {
      // A page we cannot read is a page we cannot repair. Skipping it is the
      // safe outcome: an unseen slide is never deleted. It IS counted, so the
      // caller can say "could not read N slides" instead of reporting whatever
      // was on them as lost.
      unread += end - start;
      continue;
    }
    snapshots.push(...page);
  }
  if (enrich) await enrichSnapshots(snapshots);
  trace("repair", "read the deck back", {
    range: [before, after],
    read: snapshots.length,
    unread,
    // `withSlotTag`, not `tagged`. This counts DEMO_SLOT_TAG — the harness's own
    // bookkeeping about which generated item a slide is — and says nothing about
    // whether a chart carries its config. Under the old name it read as the
    // second thing, and a real run log showing `tagged=38` next to a tag pass
    // reporting 31 looked like two reads of one deck disagreeing. They were
    // counting different tags. A number that answers a question nobody asked,
    // under the name of the question everybody asks, is worse than no number.
    withSlotTag: snapshots.filter((s) => s.slot !== null).length,
  });
  return { snapshots, unread };
}

/** The `{ i, title }` envelope `insertDemoDeck` writes, parsed defensively. */
function parseSlotTag(raw: string): { i: number; title: string | null; run: string | null } | null {
  try {
    const parsed: unknown = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    const { i, title, run } = parsed as { i?: unknown; title?: unknown; run?: unknown };
    if (typeof i !== "number" || !Number.isInteger(i) || i < 0) return null;
    return {
      i,
      title: typeof title === "string" ? title : null,
      run: typeof run === "string" ? run : null,
    };
  } catch {
    return null;
  }
}

/**
 * A token identifying ONE run, written into every slide that run adds.
 *
 * Slot indices restart at 0 each run and the demo titles are fixed, so nothing
 * else in a slot tag distinguishes this run's Title slide from the one a run an
 * hour ago left in the same deck. See `SlideSnapshot.run` for what a repair
 * pass does when it cannot tell them apart.
 */
export function newRunId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/** Passes B and C over a set of snapshots: are they tagged, and how big are their groups. */
export async function enrichSnapshots(snapshots: SlideSnapshot[]): Promise<void> {
  await markTaggedSlides(snapshots);
  await countGroupChildren(snapshots);
}

/** Pass B: does any shape on the slide carry a config tag (i.e. re-editable). */
async function markTaggedSlides(snapshots: SlideSnapshot[]): Promise<void> {
  if (!supports("1.3")) return;
  // Slides whose tag state this pass could not establish, for one more try.
  let undetermined = snapshots;
  for (let attempt = 0; attempt < 2 && undetermined.length; attempt++) {
    undetermined = await tagPass(undetermined, attempt);
  }
  // Still unknown after a second look. NOT "untagged": see `SlideSnapshot.tagRead`.
  for (const s of undetermined) s.tagRead = false;
  if (undetermined.length)
    trace("repair", "tag state undetermined after re-read", {
      slides: undetermined.length,
      indexes: undetermined.map((s) => s.index),
    });
}

/**
 * One sweep of the tag readback. Returns the slides it could NOT determine.
 *
 * The host has been observed answering this read with collections far shorter
 * than the ones pass A had just counted: on a real 38-slide run, page 2 asked
 * about 19 slides carrying 19 shapes and saw 3. Every shape it did see was
 * interrogated correctly, so the tag lookups are sound — it is the shape
 * collection that arrives hollow.
 *
 * Read naively, a hollow collection has no shapes, therefore no tags,
 * therefore "this chart is not re-editable" — and the repair then rewrites 14
 * tags that were already correct. Pass A's count is the check that catches it:
 * a slide that had shapes a moment ago and reports none now has not been read,
 * whatever else is true.
 */
/**
 * The slide indices in a page that came back with no config tag, capped.
 *
 * Pure and exported so the cap and the overflow marker can be tested without a
 * PowerPoint: the interesting case is a page where EVERY slide is untagged, and
 * reaching that through the fake costs a whole hostile-host run.
 */
export function untaggedIndices(
  page: { index: number }[],
  hit: Set<number>,
  cap = UNTAGGED_NAMED,
): (number | string)[] {
  const missing = page.filter((s) => !hit.has(s.index)).map((s) => s.index);
  return missing.length > cap ? [...missing.slice(0, cap), `…+${missing.length - cap} more`] : missing;
}

/** How many untagged slide indices a single trace line names. See `untaggedIndices`. */
const UNTAGGED_NAMED = 12;

async function tagPass(snapshots: SlideSnapshot[], attempt: number): Promise<SlideSnapshot[]> {
  const undetermined: SlideSnapshot[] = [];
  for (let start = 0; start < snapshots.length; start += READBACK_PAGE) {
    const page = snapshots.slice(start, start + READBACK_PAGE).filter((s) => s.shapes > 0);
    if (!page.length) continue;
    // Stopped: every page not reached is undetermined, which is what a page
    // that threw already reports. "We did not check" and "checked, no tag" are
    // different answers and the repair plan acts on them differently.
    if (isStopRequested()) {
      undetermined.push(...snapshots.slice(start).filter((s) => s.shapes > 0));
      trace("repair", "tag pass stopped by the user", { from: start });
      break;
    }
    try {
      const result = await boundedRun(
        `reading chart tags on slides ${page[0]?.index}-${page[page.length - 1]?.index}`,
        async (context) => {
          const perSlide = page.map((s) => {
            const shapes = context.presentation.slides.getItemAt(s.index).shapes;
            shapes.load("items/id");
            return { want: s, shapes };
          });
          await context.sync();
          const lookups = perSlide.map((p) => {
            // Read the collection ONCE. Asking twice is a second round-trip's
            // worth of answer from a host that has already been observed giving
            // two different ones, and the count below has to describe the same
            // shapes the tags were taken from.
            const items = p.shapes.items;
            return {
              want: p.want,
              seen: items.length,
              tags: items.map((shape) => {
                const t = shape.tags.getItemOrNullObject(CHART_TAG);
                t.load("value");
                return t;
              }),
            };
          });
          await context.sync();
          const hit = new Set<number>();
          const short: SlideSnapshot[] = [];
          let shapesSeen = 0;
          for (const l of lookups) {
            shapesSeen += l.seen;
            if (l.tags.some((t) => !t.isNullObject && !!t.value)) hit.add(l.want.index);
            // Fewer shapes than pass A counted: this slide was not read, so it
            // has told us nothing about its tags either way.
            else if (l.seen < l.want.shapes) short.push(l.want);
            for (const t of l.tags) untrack(t);
          }
          return { hit, short, shapesSeen };
        },
      );
      for (const s of page) if (result.hit.has(s.index)) s.tagged = true;
      undetermined.push(...result.short);
      trace("repair", "tag pass over a page", {
        attempt,
        from: page[0]?.index,
        to: page[page.length - 1]?.index,
        slides: page.length,
        shapesExpected: page.reduce((n, s) => n + s.shapes, 0),
        shapesSeen: result.shapesSeen,
        tagsFound: result.hit.size,
        undetermined: result.short.length,
        // WHICH slides came back without a config, not just how many.
        //
        // `tagsFound=14 slides=18` leaves a subtraction and no way to settle it:
        // four slides with no chart tag are either four harness pages that never
        // had one, or four charts that lost theirs, and those want opposite
        // responses. A reader can only tell by opening the deck and counting —
        // which is what this line saves, because a title page and a broken chart
        // do not look alike once you know which slide to look at.
        //
        // Capped, because this runs per page on a deck that can be hundreds of
        // slides and the whole list would bury the run log it shares.
        withoutTag: untaggedIndices(page, result.hit),
      });
    } catch {
      // A page that threw told us nothing at all — every slide on it is
      // undetermined, not untagged.
      undetermined.push(...page);
    }
  }
  return undetermined;
}

/**
 * Pass C: how many shapes are inside each slide's PowerChart group.
 *
 * PAGED, like passes A and B, and it was the one pass that was not. It opened
 * its own context with two syncs for EVERY grouped slide, which made it the
 * most expensive thing the repair does and made `READBACK_TIMEOUT_MS`
 * meaningless in aggregate: the budget is per slide, so sixty grouped slides is
 * ninety minutes of rope with a full ninety seconds left at every step. On a
 * real 38-slide run this pass accounted for a 49-second gap in the log — 43% of
 * the post-insert wall clock — immediately before the tab died.
 *
 * `getCount` is a scalar rather than a load, so the >50-item ceiling
 * (office-js#4272) that sizes `READBACK_PAGE` does not apply to the counts
 * themselves; the `items/name` loads are what the page size is for.
 *
 * A page whose sync faults falls back to one slide at a time, so a single bad
 * slide costs its own measurement rather than the whole page's. Every failure
 * is swallowed either way — `Shape.group` needs PowerPointApi 1.8, and reaching
 * for it on an older host rejects the sync it was queued in.
 */
async function countGroupChildren(snapshots: SlideSnapshot[]): Promise<void> {
  if (!supports("1.8")) return;
  // Leaves groupChildren unset for anything not measured, which planReconcile
  // already treats as "unmeasured" — see its unmeasured rule. Never as "empty".
  const wanted = snapshots.filter((s) => s.grouped);
  for (let from = 0; from < wanted.length; from += READBACK_PAGE) {
    if (isStopRequested()) break;
    const page = wanted.slice(from, from + READBACK_PAGE);
    try {
      await countGroupChildrenPage(page);
    } catch {
      // The page's sync faulted, which says nothing about any individual slide
      // on it. Retry one at a time so one unreadable group does not cost the
      // other nineteen their measurement — and a measurement missing here is
      // what makes the repair pass "fix" a chart that was never broken.
      for (const s of page) {
        if (isStopRequested()) break;
        await countGroupChildrenPage([s]).catch(() => {});
      }
    }
  }
}

/** One page of pass C: every group's child count in a single context. */
async function countGroupChildrenPage(page: SlideSnapshot[]): Promise<void> {
  const from = page[0]?.index ?? 0;
  const to = page[page.length - 1]?.index ?? from;
  const counts = await boundedRun(`counting chart groups on slides ${from + 1}-${to + 1}`, async (context) => {
    const shapes = page.map((s) => context.presentation.slides.getItemAt(s.index).shapes);
    for (const c of shapes) c.load("items/name");
    await context.sync();
    const queued = shapes.map((c) => {
      const items = loadedItems(c);
      const group = items?.find((sh) => loadedValue(() => (sh as unknown as { name: string }).name) === GROUP_NAME);
      if (!group) return undefined;
      try {
        return (group as unknown as { group: { shapes: { getCount(): { value: number } } } }).group.shapes.getCount();
      } catch {
        return undefined;
      }
    });
    await context.sync();
    return queued.map((c) => (c ? loadedValue(() => c.value) : undefined));
  });
  counts.forEach((n, i) => {
    if (typeof n === "number") page[i].groupChildren = n;
  });
}

/** Delete the NOT COMPLETE banner from a slide. False when there is none. */
async function deleteStamp(slideIndex: number): Promise<boolean> {
  try {
    return await boundedRun(`removing the banner from slide ${slideIndex + 1}`, async (context) => {
      type Named = { name: string; delete(): void };
      const shapes = context.presentation.slides.getItemAt(slideIndex).shapes as unknown as {
        items: (Named & { group?: { shapes: { items: Named[]; load(p: string): void } } })[];
        load(p: string): void;
      };
      shapes.load("items/name");
      await context.sync();
      const stamp = shapes.items.find((s) => s.name === NOT_COMPLETE_NAME);
      if (stamp) {
        stamp.delete();
        await context.sync();
        return true;
      }
      // Not on the slide — look inside the chart's group, where an older
      // rescue may have swept it before that code learned to leave the stamp
      // out. Deleting one member leaves the group itself intact.
      const group = shapes.items.find((s) => s.name === GROUP_NAME);
      const inner = group?.group?.shapes;
      if (!inner) return false;
      inner.load("items/name");
      await context.sync();
      const buried = inner.items.find((s) => s.name === NOT_COMPLETE_NAME);
      if (!buried) return false;
      buried.delete();
      await context.sync();
      return true;
    });
  } catch {
    return false;
  }
}

/**
 * Delete one slide by index, in its own settled context — after checking it is
 * still the slide the plan profiled.
 *
 * A plan is a list of positions, decided from a readback taken some round-trips
 * ago. Positions are not identities: the pane does not lock the deck while a
 * repair runs, and nothing stops the user reordering or deleting slides in
 * PowerPoint's own UI between two of the pass's awaited calls. Every other
 * repair re-checks before it acts (`deleteStamp` matches the shape's name);
 * this one, the only irreversible one, took the index on faith.
 *
 * So re-read the slot tag and the shape count and compare them with what was
 * profiled. A mismatch refuses the delete — reported as a step the host would
 * not do, which is exactly what it is from the user's side.
 */
async function deleteSlide(index: number, expect?: SlideSnapshot): Promise<boolean> {
  try {
    return await boundedRun(`removing slide ${index + 1}`, async (context) => {
      const slide = context.presentation.slides.getItemAt(index);
      if (expect) {
        const shapes = slide.shapes;
        shapes.load("items/name");
        const tag = supports("1.3")
          ? ((
              slide as unknown as {
                tags: {
                  getItemOrNullObject(k: string): { isNullObject: boolean; value: string; load(p?: string): void };
                };
              }
            ).tags.getItemOrNullObject(DEMO_SLOT_TAG) ?? null)
          : null;
        // LOAD it, do not merely resolve it. A getItemOrNullObject proxy nobody
        // loads takes no part in the sync, so `isNullObject` is never populated
        // and reading it throws PropertyNotLoaded — the same trap that made
        // editing an ungrouped chart impossible (see CHART_PARTS_TAG).
        //
        // Here the throw lands inside this function's own catch, so the failure
        // was silent and total: every guarded delete refused, every time, on any
        // host that actually enforces the load. The repair pass could then never
        // remove a duplicate slide it had correctly identified — it reported the
        // step as one the host would not do, and the duplicates stayed.
        //
        // The cast is why this survived: the inline type it was given had no
        // `load` on it, so calling it would not have compiled and its absence
        // read as deliberate. The type now carries `load`.
        tag?.load("value");
        await context.sync();
        const slot = tag && !tag.isNullObject ? parseSlotTag(tag.value) : null;
        // The identity that made this slide deletable in the first place.
        if ((slot?.i ?? null) !== expect.slot || (slot?.run ?? null) !== (expect.run ?? null)) return false;
        // …and that it still holds what it held. A slide that gained shapes
        // since the readback is one somebody has been working on.
        if (shapes.items.length !== expect.shapes) return false;
      }
      (slide as unknown as { delete(): void }).delete();
      await context.sync();
      return true;
    });
  } catch {
    return false;
  }
}

/** What a repair actually managed to do. */
export interface ReconcileOutcome {
  /** The deck as it was READ, before any repair — the run's raw evidence. */
  snapshots: SlideSnapshot[];
  plan: ReconcilePlan;
  /** Banners removed, charts re-grouped, slides deleted — all confirmed. */
  applied: { unstamped: number; regrouped: number; deleted: number };
  /** Actions the host refused. The plan is re-runnable; nothing is corrupted. */
  refused: number;
  /**
   * Slides in range the readback could not see at all. Anything on them is
   * reported `lost` for want of evidence, so a non-zero count here is the
   * difference between "the host dropped it" and "we could not look".
   */
  unread?: number;
}

/**
 * Apply a plan. Deletes run last and in descending index order — a plan is
 * written against the deck as it was READ, and removing slide 5 renumbers
 * every slide after it. `planReconcile` guarantees that ordering; this relies
 * on it and would otherwise delete the wrong slides.
 *
 * Every step is best-effort and independent: a refused delete does not stop
 * the next unstamp. Re-running the pass is always safe, which is what makes it
 * usable as a button.
 */
export async function applyReconcilePlan(
  plan: ReconcilePlan,
  tagFor: (slot: number) => string | undefined,
  origin: { left: number; top: number } = { left: 60, top: 90 },
  snapshots: SlideSnapshot[] = [],
): Promise<ReconcileOutcome> {
  const applied = { unstamped: 0, regrouped: 0, deleted: 0 };
  let refused = 0;
  for (const action of plan.actions) {
    // Between actions, never inside one. Every action so far has committed and
    // the rest have not been touched — the same boundary the draw loop stops
    // at, and the only one where a repair pass that DELETES slides can be
    // interrupted without leaving the deck half-repaired.
    if (isStopRequested()) {
      trace("repair", "repair pass stopped by the user", { applied: { ...applied }, remaining: plan.actions.length });
      break;
    }
    trace("repair", `applying ${action.kind}`, { index: action.index, slot: action.slot, reason: action.reason });
    if (action.kind === "unstamp") {
      if (await deleteStamp(action.index)) applied.unstamped++;
      else refused++;
    } else if (action.kind === "regroup") {
      if (await rescueGroupAndTag(action.index, tagFor(action.slot), origin)) applied.regrouped++;
      else refused++;
    } else if (action.kind === "retag") {
      if (await retagSlideChart(action.index, tagFor(action.slot))) applied.regrouped++;
      else refused++;
    } else {
      // Deletes are checked against the snapshot that authorised them. The
      // snapshots are optional, so a caller that passes none still gets the
      // old unchecked behaviour rather than a repair that silently does
      // nothing.
      if (
        await deleteSlide(
          action.index,
          snapshots.find((s) => s.index === action.index),
        )
      )
        applied.deleted++;
      else refused++;
    }
  }
  return { snapshots, plan, applied, refused };
}

/**
 * The whole repair: read the deck back, decide, act.
 *
 * Called at the end of a demo run — when the host has finally stopped moving,
 * which is the only moment any of this can be measured — and again, on demand,
 * by the pane's repair action, which is how a deck damaged by an EARLIER run
 * gets fixed without re-inserting anything.
 */
export async function reconcileDeck(
  expected: ExpectedItem[],
  range: { before: number; after: number },
  tagFor: (slot: number) => string | undefined,
  opts: ReconcileOptions = {},
): Promise<ReconcileOutcome> {
  const { snapshots, unread } = await readAddedSlides(range.before, range.after);
  const plan = planReconcile(snapshots, expected, opts);
  return { ...(await applyReconcilePlan(plan, tagFor, { left: 60, top: 90 }, snapshots)), unread };
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

/**
 * Whole-deck insert in ONE host call: `Presentation.insertSlidesFromBase64`.
 *
 * Everything else in this file draws a chart the only way Office.js allows —
 * shape by shape, batch by batch, sync by sync — and that is precisely what
 * PowerPoint on the web cannot be trusted with. A 12-item run issues 19 slide
 * adds and hundreds of queued commands, and every one of them is a chance for
 * the host to drop an add, stall a sync past its timeout, or refuse a group.
 * The 2026-07-30 run took 680 seconds and shipped duplicate slides.
 *
 * A .pptx handed over as base64 has none of those failure surfaces. The
 * grouping is in the file. The tags are in the file. There is one call to
 * lose, and a slide-count delta proves whether it landed.
 *
 * The requirement set is deliberately probed rather than asserted: Microsoft's
 * own docs disagree about which set carries this method (the 1.2 "what's new"
 * page announces slide insertion; the API reference cites 1.5), and the
 * manifests admit hosts from 1.4. So: check the method actually exists on the
 * proxy, and treat any host that lacks it as a fallback case rather than an
 * error.
 */
export function canInsertSlidesFromBase64(): boolean {
  return isPowerPointHost() && (supports("1.5") || supports("1.2"));
}

/**
 * How long one deck insert may take before the host counts as stalled. Scales
 * with the deck, because unlike a shape batch this call's size is not capped —
 * handing over 40 slides is legitimately more work than handing over one.
 */
/** Allowance per slide in a one-call deck insert. See `DECK_INSERT_TIMEOUT_MS`. */
let DECK_INSERT_PER_SLIDE_MS = 5_000;

/** Test-only: a suite cannot spend five seconds per slide finding out. */
export function _setDeckInsertPerSlideForTest(ms: number): void {
  DECK_INSERT_PER_SLIDE_MS = ms;
}

const DECK_INSERT_TIMEOUT_MS = (slides: number): number =>
  Math.max(BATCH_TIMEOUT_MS, slides * DECK_INSERT_PER_SLIDE_MS);

/**
 * Insert a whole .pptx and return how many slides the deck actually gained.
 *
 * The count is measured, not assumed: a settled `getCount()` before and after,
 * in their own contexts, for the same reason `addSlides` does it — a host that
 * silently drops the call reports no error, and the delta is the only evidence
 * either way. `expectedSlides` only sizes the timeout.
 *
 * `formatting` defaults to keeping the source's own formatting: PowerChart
 * writes explicit colours and fonts on every shape, so adopting the
 * destination theme would override deliberate choices rather than harmonise
 * them.
 */
export async function insertSlidesFromPptx(
  base64: string,
  expectedSlides: number,
  formatting: "KeepSourceFormatting" | "UseDestinationTheme" = "KeepSourceFormatting",
): Promise<number> {
  const before = await slideCount();
  await withTimeoutOrVerify(
    PowerPoint.run(async (context) => {
      const presentation = context.presentation as unknown as {
        insertSlidesFromBase64(b64: string, opts?: { formatting?: string; targetSlideId?: string }): void;
      };
      if (typeof presentation.insertSlidesFromBase64 !== "function") {
        throw new Error("this host has no insertSlidesFromBase64");
      }
      // APPEND. Without a target the host inserts at the FRONT of the deck,
      // which is not where a demo deck belongs and quietly breaks anything
      // that reasons about "the slides this run added" by position — the first
      // real run put 37 generated slides ahead of the user's own title slide.
      let targetSlideId: string | undefined;
      try {
        const last = context.presentation.slides.getItemAt(before - 1);
        last.load("id");
        await context.sync();
        targetSlideId = last.id;
      } catch {
        /* empty deck, or a host that will not name the slide — front it is */
      }
      presentation.insertSlidesFromBase64(base64, targetSlideId ? { formatting, targetSlideId } : { formatting });
      await context.sync();
    }),
    DECK_INSERT_TIMEOUT_MS(expectedSlides),
    `inserting ${expectedSlides} slide(s) from a generated deck`,
  );
  const after = await slideCount();
  trace("insert", "handed the host a generated deck", {
    expectedSlides,
    landed: after - before,
    base64Bytes: base64.length,
    formatting,
  });
  return after - before;
}

/**
 * Put the view ON a slide, and say whether it worked.
 *
 * The exact inverse of `withSlideDeselected`, and it exists for the same
 * reason: what the user is looking at changes how the host behaves. The
 * live-canvas repaint is the add-in's worst case — a real run died on its
 * first batch with "did not respond while drawing shapes 1-10 of 39" — and
 * every guard written for it has only ever been exercised against a fake. The
 * self-test uses this to put a chart genuinely on screen before redrawing it,
 * because a test that quietly redraws off-screen is testing the easy case.
 *
 * Best-effort on the same terms: `setSelectedSlides` is PowerPointApi 1.5.
 */
export async function showSlide(slideId: string, budgetMs?: number): Promise<boolean> {
  try {
    return await boundedRun(
      "moving the view to a slide",
      async (context) => {
        const presentation = context.presentation as unknown as {
          setSelectedSlides(ids: string[]): void;
        };
        presentation.setSelectedSlides([slideId]);
        await context.sync();
        return true;
      },
      budgetMs,
    );
  } catch {
    return false;
  }
}

/**
 * Whether this host can be told which SHAPE is selected.
 *
 * `Slide.setSelectedShapes` is PowerPointApi **1.5** — the same set this add-in
 * already needs for `getSelectedShapes` and `setSelectedSlides`, so in practice
 * a host that can read the selection can also set it. It is gated anyway
 * because "in practice" is how the last four host assumptions were written.
 */
export const canSelectShapes = (): boolean => supports("1.5");

/**
 * Whether the host will tell us when the user changes the selection.
 *
 * `DocumentSelectionChanged` is a **Common API** event, not a PowerPointApi
 * one — so it is available on hosts far below 1.5, and, more to the point, it
 * does not go through the selection subsystem that a programmatic
 * `setSelectedShapes` wedges. The pane has used it for the "A PowerChart is
 * selected" banner all along.
 */
export function canWatchSelection(): boolean {
  try {
    return typeof Office?.context?.document?.addHandlerAsync === "function";
  } catch {
    return false;
  }
}

/** What a wait for a human click ended in. All three are different findings. */
export interface SelectionWait {
  /** The chart they clicked, when one came back. */
  chart: { configJson: string; target: EditTarget } | null;
  /** Whether the selection changed at all — i.e. whether anyone did anything. */
  sawClick: boolean;
  /** Whether a read was attempted and the host refused or never answered. */
  readFailed: boolean;
}

/**
 * Wait for the user to click a PowerChart, and hand back what they clicked.
 *
 * The one path a real user travels that nothing can script. `setSelectedShapes`
 * is Office.js selecting a shape — the same call in theory as a human clicking
 * one, and on PowerPoint on the web demonstrably not the same in practice: it
 * is taken, and the selection subsystem then stops answering. So the battery
 * cannot reach the pane's most-used read by driving it, and reaching it by
 * asking a person to click is not a workaround, it is the only honest version
 * of the test.
 *
 * Nothing here touches `setSelectedShapes`. It listens, which is what the pane
 * itself does, and reads the selection the host volunteers.
 *
 * Reports whether anyone clicked AT ALL, separately from whether a chart came
 * back. "Nobody clicked" and "you clicked and the host would not tell us what"
 * are different findings, and only the first is the person's doing — blaming
 * them for the second is how a report stops being read.
 *
 * The two budgets are also separate, and conflating them was a bug worth
 * naming: `budgetMs` is how long to wait for a HUMAN, `readBudgetMs` is how
 * long to let the host answer once one has acted. Using the first for both did
 * not merely make a wedged host slow — a click at 29 seconds started a read
 * with 30 seconds of rope, the deadline cut it off one second later, and the
 * scenario reported "nobody clicked" about a person who had just clicked.
 */
export async function awaitSelectedChart(
  budgetMs: number,
  readBudgetMs: number,
  onWaiting?: (secondsLeft: number) => void,
): Promise<SelectionWait> {
  if (!canWatchSelection()) return { chart: null, sawClick: false, readFailed: false };
  return new Promise((resolve) => {
    let done = false;
    let sawClick = false;
    let readFailed = false;
    /** The read a click started, so the deadline can wait for it to land. */
    let reading: Promise<unknown> | null = null;
    const handler = async () => {
      if (done) return;
      sawClick = true;
      const read = (async () => {
        try {
          const found = await loadChartFromSelection(readBudgetMs);
          if (!found || done) return;
          finish(found);
        } catch {
          // The host was asked and did not answer. Recorded rather than
          // swallowed: "you clicked something that is not a chart" and "the
          // host would not tell us what you clicked" are different findings,
          // and only the first is about what the user did.
          readFailed = true;
        }
      })();
      reading = read;
      await read;
      if (reading === read) reading = null;
    };
    const finish = (value: { configJson: string; target: EditTarget } | null) => {
      if (done) return;
      done = true;
      clearInterval(tick);
      clearTimeout(deadline);
      try {
        Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler });
      } catch {
        /* nothing to remove, or a host that will not — the flag stops it anyway */
      }
      resolve({ chart: value, sawClick, readFailed });
    };
    const startedAt = Date.now();
    // A countdown, because a battery that silently waits is indistinguishable
    // from one that has hung — and this one is waiting for a PERSON, who needs
    // to know they are the thing being waited for.
    const tick = setInterval(() => {
      const left = Math.ceil((budgetMs - (Date.now() - startedAt)) / 1000);
      if (left > 0) onWaiting?.(left);
    }, 1_000);
    // A click already being read is not a missed deadline. Let it land first —
    // it is bounded by `readBudgetMs` and cannot extend this indefinitely.
    const deadline = setTimeout(() => void (reading ?? Promise.resolve()).then(() => finish(null)), budgetMs);
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
    } catch {
      finish(null);
    }
  });
}

/** One rung of `selectionLadder` — a single call, and what the host did with it. */
export interface LadderRung {
  /** The call, in the words a report should use. */
  step: string;
  /** `ok` answered, `refused` threw, `silent` never came back at all. */
  outcome: "ok" | "refused" | "silent";
  ms: number;
  /** What was read, or what the refusal said. */
  detail?: string;
}

/** A bare `getSelectedShapes` read — no tags, no targets, just "did it answer". */
async function readSelectionCount(budgetMs: number): Promise<number> {
  return boundedRun(
    "reading the selection",
    async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/id");
      await context.sync();
      // -1 is not "none". It is "the host answered the sync and then would not
      // say what is selected", which is a third outcome and must not be
      // reported as an empty selection.
      return loadedItems(shapes)?.length ?? -1;
    },
    budgetMs,
  );
}

/** One raw `setSelectedShapes`, with nothing swallowed — the point is what it does. */
async function setShapeSelection(slideId: string, ids: string[], budgetMs: number): Promise<void> {
  return boundedRun(
    ids.length ? "selecting a shape" : "clearing the shape selection",
    async (context) => {
      const slide = context.presentation.slides.getItemOrNullObject(slideId);
      queueNullCheck(slide);
      await context.sync();
      if (!isLive(slide)) throw new Error("the host would not resolve the slide");
      (slide as unknown as { setSelectedShapes(x: string[]): void }).setSelectedShapes(ids);
      await context.sync();
    },
    budgetMs,
  );
}

/**
 * Which selection call, exactly, stops PowerPoint on the web from answering.
 *
 * We have two accounts and they name different culprits. This project measured
 * a run where `setSelectedShapes([id])` — the NON-empty select — was taken
 * happily and every selection call after it went silent for a full 90-second
 * budget. office-js#3698 reports that `setSelectedShapes([])` — the EMPTY
 * clear — "causes the `PowerPoint.run` promise to never resolve". Both can be
 * true, but the add-in cannot be built on "probably both", and another blind
 * round would produce the same ambiguity at the same cost.
 *
 * So: a ladder, run in one context after another, from the least invasive call
 * to the most. **It stops at the first rung that goes silent**, and that is the
 * whole design. Once the subsystem is wedged every later call is silent too, so
 * running the rest would produce four timeouts and name nothing — which is
 * precisely the ambiguity being resolved. One silence, with the call that
 * preceded it, is an answer.
 *
 * The rungs interleave writes with reads deliberately: a write that is *taken*
 * and a subsystem that still *answers* are different facts, and the observed
 * failure has the first without the second.
 */
export async function selectionLadder(slideId: string, shapeId: string, budgetMs: number): Promise<LadderRung[]> {
  const rungs: LadderRung[] = [];
  const run = async (step: string, fn: () => Promise<string>): Promise<boolean> => {
    const at = Date.now();
    try {
      const detail = await fn();
      rungs.push({ step, outcome: "ok", ms: Date.now() - at, detail });
      return true;
    } catch (err) {
      const silent = isTimeout(err);
      rungs.push({
        step,
        outcome: silent ? "silent" : "refused",
        ms: Date.now() - at,
        detail: errorText(err),
      });
      // A refusal is informative and survivable — the host said no and is still
      // talking, so the ladder carries on. Silence is terminal: nothing after
      // it can be attributed to anything.
      return !silent;
    }
  };
  const read = (when: string) =>
    run(`read the selection ${when}`, async () => `${await readSelectionCount(budgetMs)} shape(s)`);

  // Rung 0 is the control, and skipping it would invalidate everything: if the
  // host will not answer a selection read BEFORE anything has touched it, the
  // wedge is not ours to have caused and no later rung means a thing.
  if (!(await read("before anything touches it"))) return rungs;
  // The pane does this on its normal path already, so a silence here would be a
  // far bigger finding than the one being chased.
  if (!(await run("setSelectedSlides (the pane's own call)", async () => `${await showSlide(slideId, budgetMs)}`)))
    return rungs;
  if (!(await read("after selecting the slide"))) return rungs;
  // The call this project measured wedging.
  if (
    !(await run(
      "setSelectedShapes([id])",
      async () => (await setShapeSelection(slideId, [shapeId], budgetMs), "taken"),
    ))
  )
    return rungs;
  if (!(await read("after selecting a shape"))) return rungs;
  // The call office-js#3698 says never resolves.
  if (!(await run("setSelectedShapes([])", async () => (await setShapeSelection(slideId, [], budgetMs), "taken"))))
    return rungs;
  await read("after clearing the selection");
  return rungs;
}

/**
 * Select one shape, so the pane's selection-driven paths can be reached without
 * a human clicking.
 *
 * The setter is on **Slide**, not `Presentation` — which is the whole reason
 * the self-test believed for months that this could not be done at all (see the
 * header of `selftest.ts`). `Presentation` has `getSelectedShapes` and
 * `setSelectedSlides` but no `setSelectedShapes`; the shape half lives one class
 * down.
 *
 * The slide is selected first. A shape on a slide nobody is looking at is not a
 * selection the user could have made, and on some hosts the shape set is
 * ignored unless its slide is current.
 */
export async function selectShape(slideId: string, shapeId: string, budgetMs?: number): Promise<boolean> {
  if (!canSelectShapes()) return false;
  await showSlide(slideId, budgetMs);
  try {
    return await boundedRun(
      "selecting a shape",
      async (context) => {
        const slide = context.presentation.slides.getItemOrNullObject(slideId);
        queueNullCheck(slide);
        await context.sync();
        if (!isLive(slide)) return false;
        (slide as unknown as { setSelectedShapes(ids: string[]): void }).setSelectedShapes([shapeId]);
        // The inner bound has to honour the caller's budget too. It did not,
        // and a caller asking for ten seconds still waited the full ninety
        // here — the outer `boundedRun` cannot cut in while this one holds.
        await step("selecting a shape", () =>
          withTimeout(context.sync(), budgetMs ?? READBACK_TIMEOUT_MS, "selecting a shape"),
        );
        return true;
      },
      budgetMs,
    );
  } catch {
    return false;
  }
}

/**
 * Drop the shape selection — by moving the SLIDE selection, and ONLY that.
 *
 * This used to ask for the clear first and re-select the slide as a fallback.
 * The empty-array call is now gone entirely, because reading the tracker
 * properly turned up something the earlier comment here did not know:
 * office-js#3698 reports that `slide.setSelectedShapes([])` on the web "does
 * not clear the selection, causes the `PowerPoint.run` promise to never
 * resolve, and produces no error messages."
 *
 * So the call was doing nothing useful and possibly doing the damage. It cannot
 * clear the selection on the web (office-js#3083, the reason the slide
 * re-select existed as a fallback in the first place), and it is a documented
 * candidate for the silence measured here. Keeping it meant paying a full
 * budget to be told nothing, on the exact host it does not work on.
 *
 * Re-selecting the slide drops the shape selection on every host observed, and
 * is what the fallback already did. What is left is the half that works.
 *
 * The reason any of this matters is downstream: on the web a picture cannot be
 * inserted while another shape is selected (office-js#3698 again), so code that
 * leaves a chart selected breaks the NEXT operation rather than its own — the
 * worst kind of failure to diagnose.
 */
export async function clearShapeSelection(slideId: string, budgetMs?: number): Promise<void> {
  if (!canSelectShapes()) return;
  await showSlide(slideId, budgetMs);
}

/**
 * Drop the shape selection without being told which slide it is on.
 *
 * `clearShapeSelection` needs a slide id, and the everyday insert has none: the
 * user selected a shape, the pane read its bounds, and nothing in that flow
 * ever names the slide. So this asks the host which slide is showing and hands
 * that to the same clear.
 *
 * It exists because of two published PowerPoint-on-the-web bugs that both fire
 * on that flow, in opposite ways, from the same cause — a shape being selected
 * while the add-in draws:
 *
 * - **office-js#2775**: adding a text box DELETES the shape that was selected.
 *   Every chart this add-in draws contains text boxes, and the insert path
 *   deliberately leaves the user's shape selected because that is how it learns
 *   where to put the chart. Selecting a picture and inserting a chart beside it
 *   would silently destroy the picture.
 * - **office-js#3698**: a picture cannot be inserted while another shape is
 *   selected — and the same insert path inserts a picture when a chart is too
 *   dense to draw as shapes.
 *
 * Neither is confirmed on the host the owner actually uses; the self-test's
 * "a selected shape survives an insert" is what would confirm the first. This
 * runs anyway, because it costs one selection call the update path already
 * makes twice, and the failure it prevents is the user losing their own
 * content. Best-effort throughout — a host that will not say which slide is
 * showing leaves the selection exactly as it was, which is today's behaviour.
 */
export async function dropShapeSelection(budgetMs?: number): Promise<boolean> {
  if (!canSelectShapes()) return false;
  try {
    const slideId = await boundedRun(
      "reading which slide is showing",
      async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        slide.load("id");
        await context.sync();
        return loadedValue(() => slide.id);
      },
      budgetMs ?? SELECTION_TIMEOUT_MS,
    );
    if (typeof slideId !== "string" || !slideId) return false;
    await clearShapeSelection(slideId, budgetMs);
    return true;
  } catch {
    return false;
  }
}

/**
 * A PNG of one slide, as the host itself draws it.
 *
 * `Slide.getImageAsBase64` is PowerPointApi **1.8**. This is the only way an
 * add-in can see its own output: every other check in this file counts shapes
 * and reads tags, which says a chart was CONSTRUCTED and says nothing about
 * whether anything is visible. A chart drawn entirely in white, or off the
 * slide, or with every shape at zero size, passes every structural assertion
 * this project has.
 *
 * Undefined rather than throwing on any host that will not do it — the caller
 * reports that as skipped, which is a different answer from a failure.
 */
export async function slideImageBase64(slideId: string, width?: number): Promise<string | undefined> {
  if (!supports("1.8")) {
    trace("host", "no rasteriser on this host", { slideId, need: "PowerPointApi 1.8" });
    return undefined;
  }
  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemOrNullObject(slideId);
      queueNullCheck(slide);
      await context.sync();
      if (!isLive(slide)) {
        trace("host", "cannot rasterise: the host would not resolve the slide", { slideId });
        return undefined;
      }
      // Rasterise through a FRESH handle, never the one just resolved.
      //
      // Resolving a `getItemOrNullObject` proxy is what makes Office.js rewrite
      // its object path to `getItem(id)`, and a freshly-added slide's id does
      // not round-trip through `getItem` on the web — the same trap `SlideThunk`
      // exists for. It cost the last scenario of a real self-test run: the
      // visibility check adds a scratch slide, rasterises it, and PowerPoint on
      // the web answered "GeneralException | errorLocation: SlideCollection.-
      // getItem | statement: var slide = slides.getItem(...); slide.getImageAs-
      // Base64(...)" — the liveness check above having passed a moment earlier
      // on the very same id, through the very same proxy.
      const img = (
        context.presentation.slides.getItemOrNullObject(slideId) as unknown as {
          getImageAsBase64(options?: { width?: number }): { value: string };
        }
      ).getImageAsBase64(width ? { width } : undefined);
      // Bounded, and this is the one that most needed it: rasterising a whole
      // slide is the heaviest single call the add-in makes, it is brand new, and
      // it had no deadline and no stop check. A self-test that wedges here
      // cannot be cancelled and cannot be timed out — which is exactly what a
      // real host did at 1819 seconds.
      await step("rasterising a slide", () => withTimeout(context.sync(), rasteriseTimeoutMs(), "rasterising a slide"));
      const v = loadedValue(() => img.value);
      if (typeof v === "string" && v.length) return v;
      // Took the call, raised nothing, produced nothing.
      //
      // The quietest of the four ways this returns undefined, and the one that
      // had no line at all. A real run's visibility scenario went straight from
      // "rasterising the empty slide" to "removing the scratch slide" with
      // nothing in between — a skipped verdict whose reason existed nowhere.
      // The round before, the same call at least said `GeneralException`; the
      // fresh handle stopped it throwing without making it work, and traded a
      // loud wrong answer for a silent one. Both are failures; only one can be
      // diagnosed.
      trace("host", "the host took the rasterise and returned nothing", {
        slideId,
        width,
        read: typeof v,
      });
      return undefined;
    });
  } catch (err) {
    trace("host", "rasterising a slide threw", { slideId, error: errorText(err) });
    return undefined;
  }
}

/**
 * One slide's picture, or the reason there isn't one — for every id asked about.
 *
 * Never a shorter array than it was given. A run that comes back with four
 * pictures for six slides leaves the reader guessing which two are missing and
 * why, and "the host has no rasteriser" and "that slide would not resolve" are
 * different findings that a gap in an array cannot tell apart.
 */
export interface SlideShot {
  slideId: string;
  /** Base64 PNG, absent when the host would not produce one. */
  png?: string;
}

/**
 * Ask the host to draw the slides a diagnostic round touched.
 *
 * The point is to stop asking a person to take screenshots. Several real bugs
 * this project has fixed were visible only in a picture — a chart drawn where
 * nothing else could see it, a slide that ended up blank, 41 shapes that became
 * 79 — and every one of them cost a round trip because the run log carried
 * counts and the screenshot lived in a different message.
 *
 * Bounded three ways, because it is the heaviest call the add-in makes and this
 * is a diagnostic, not a feature: a cap on how many slides, a width, and
 * `slideImageBase64`'s own deadline per slide. Over the cap, the EXTRA IDS ARE
 * STILL RETURNED, without pictures and marked — a capped run that silently
 * showed the first twelve of twenty would read as a deck of twelve.
 */
export async function slideShots(
  slideIds: string[],
  opts: { max?: number; width?: number } = {},
): Promise<SlideShot[]> {
  const max = opts.max ?? 12;
  const out: SlideShot[] = [];
  // Three different reasons a slide comes back without a picture, counted
  // apart. They used to share one line — `slides the host would not draw
  // {asked: 22, drew: 12, max: 12}` — which reads as a host refusing ten
  // slides, and every real round said exactly that while the host had refused
  // nothing at all: ten slides were over OUR cap and never asked about. That is
  // the same "never asked looks like answered no" mistake the contract gate
  // used to make, in the message a reader reaches for first.
  let overCap = 0;
  let stopped = 0;
  for (const slideId of slideIds) {
    if (out.length >= max) {
      overCap++;
      out.push({ slideId });
      continue;
    }
    if (isStopRequested()) {
      stopped++;
      out.push({ slideId });
      continue;
    }
    const png = await slideImageBase64(slideId, opts.width ?? 480);
    if (!png) out.push({ slideId });
    else out.push({ slideId, png });
  }
  const drew = out.filter((s) => s.png).length;
  const refused = out.length - drew - overCap - stopped;
  if (refused > 0) trace("host", "slides the host would not draw", { asked: out.length, drew, refused });
  if (overCap > 0 || stopped > 0)
    trace("host", "slides never asked about", { asked: out.length, drew, overCap, stopped, max });
  return out;
}

/** Every slide id in the deck, in order — or undefined if the host would not list them. */
export async function deckSlideIds(): Promise<string[] | undefined> {
  return slideIds();
}

/** Where a slide-size answer came from — recorded so a wrong placement can be
 *  traced to the rung that produced the number. */
export type SlideSizeSource = "pageSetup" | "exportedSlide" | "documentFile" | "assumed";

export interface SlideSize {
  /** Points. */
  width: number;
  /** Points. */
  height: number;
  source: SlideSizeSource;
}

/**
 * 16:9, the overwhelmingly common default, used when nothing can be read.
 *
 * Only ever the floor of the ladder below, and always labelled `assumed` so no
 * caller mistakes it for a measurement.
 */
const ASSUMED_SLIDE_SIZE: SlideSize = { width: 960, height: 540, source: "assumed" };

/**
 * Cached because two of the three rungs below are expensive, and slide size is
 * a property of the deck rather than of any one operation.
 *
 * Not permanent: a user can change slide size mid-session from PowerPoint's own
 * Design tab, and a cached 16:9 would then place charts off the edge of a deck
 * that is now 4:3. `slideSize({ refresh: true })` re-reads, and the cheap rung
 * re-reads anyway (one property load costs a fraction of the sync it rides on).
 */
let cachedSlideSize: SlideSize | undefined;

/** Drop the cached slide size — for tests, and for a deck whose setup changed. */
export function _resetSlideSizeCache(): void {
  cachedSlideSize = undefined;
}

/** Read `<p:sldSz>` out of a base64 .pptx, in points. */
async function slideSizeFromPptxBase64(base64: string): Promise<SlideSize | null> {
  try {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(base64, { base64: true });
    const part = zip.file("ppt/presentation.xml");
    if (!part) return null;
    const emu = parseSlideSizeEmu(await part.async("string"));
    if (!emu) return null;
    return {
      width: emu.cx / EMU_PER_POINT,
      height: emu.cy / EMU_PER_POINT,
      source: "exportedSlide",
    };
  } catch {
    return null;
  }
}

/**
 * The presentation's slide dimensions, in points.
 *
 * This add-in spent its life assuming them. `placeChart` still says why that
 * was survivable — every standard slide is 540pt tall, 4:3 and 16:9 alike — but
 * it was only ever true of the HEIGHT, and width is the dimension that decides
 * whether a second chart can sit beside the first or has to go underneath it.
 * The assumption is also what makes the generated-deck fast path rescale on a
 * 4:3 deck, since the file it builds declares 16:9.
 *
 * Three rungs, cheapest and most exact first:
 *
 *  1. **`presentation.pageSetup`** (PowerPointApi 1.10) states both dimensions
 *     in points. One property load on a sync we are opening anyway.
 *  2. **`slide.exportAsBase64()`** (1.8) hands back the slide as its own .pptx,
 *     and that file's `ppt/presentation.xml` carries the SOURCE deck's
 *     `<p:sldSz>`. Exact, and jszip is already a dependency, but it is a whole
 *     file round-trip for two numbers.
 *  3. **`Office.context.document.getFileAsync`** (Common API, no PowerPointApi
 *     gate at all) returns the entire deck in 4MB slices. Same parse, and the
 *     only rung that works below 1.8 — but it copies the whole presentation,
 *     which on a large deck is seconds and tens of MB, so it is genuinely last.
 *
 * Anything that fails falls through to the next rung, and a total failure
 * answers 16:9 marked `assumed` rather than throwing: a wrong-but-labelled
 * width degrades placement to what it did before, while a throw would take down
 * an insert over a layout hint.
 */
export async function slideSize(opts: { refresh?: boolean } = {}): Promise<SlideSize> {
  if (!opts.refresh && cachedSlideSize) return cachedSlideSize;

  // Rung 1 — the direct read.
  if (supports("1.10")) {
    try {
      // Bounded: this is rung ONE of a ladder whose whole design is "fall
      // through to the next rung", and it runs on the Insert click path. An
      // unbounded first rung turns a graceful degradation into a hang — the
      // ladder never reaches rungs 2, 3 or 4, and the pane never draws.
      const got = await boundedRun(
        "reading the slide size",
        async (context) => {
          const setup = (context.presentation as unknown as { pageSetup: { slideWidth: number; slideHeight: number } })
            .pageSetup;
          (setup as unknown as { load(p: string): void }).load("slideWidth,slideHeight");
          await context.sync();
          return { width: setup.slideWidth, height: setup.slideHeight };
        },
        SELECTION_TIMEOUT_MS,
      );
      if (Number.isFinite(got.width) && Number.isFinite(got.height) && got.width > 0 && got.height > 0) {
        cachedSlideSize = { ...got, source: "pageSetup" };
        trace("host", "slide size read", { ...cachedSlideSize });
        return cachedSlideSize;
      }
    } catch {
      /* 1.10 advertised but the property is not there — try the next rung */
    }
  }

  // Rung 2 — export one slide and read what it declares.
  if (supports("1.8")) {
    try {
      const base64 = await PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        const out = (slide as unknown as { exportAsBase64(): { value: string } }).exportAsBase64();
        await context.sync();
        return out.value;
      });
      const got = base64 ? await slideSizeFromPptxBase64(base64) : null;
      if (got) {
        cachedSlideSize = got;
        trace("host", "slide size read", { ...cachedSlideSize });
        return cachedSlideSize;
      }
    } catch {
      /* empty deck, or a host that will not export — try the next rung */
    }
  }

  // Rung 3 — the whole document, through the Common API.
  const fromFile = await slideSizeFromDocumentFile();
  if (fromFile) {
    cachedSlideSize = { ...fromFile, source: "documentFile" };
    trace("host", "slide size read", { ...cachedSlideSize });
    return cachedSlideSize;
  }

  trace("host", "slide size unavailable — assuming 16:9", { ...ASSUMED_SLIDE_SIZE });
  cachedSlideSize = ASSUMED_SLIDE_SIZE;
  return cachedSlideSize;
}

/**
 * The deepest rung: pull the whole presentation through the Common API.
 *
 * `Office.context.document.getFileAsync` predates the PowerPointApi requirement
 * sets entirely, so this is the only read available on a 1.4 host. It is also
 * by far the most expensive — the entire deck, in 4MB slices, copied to reach
 * two numbers in the first part — which is why nothing calls it until the two
 * cheaper rungs have declined.
 *
 * The file handle MUST be closed. An add-in that leaves one open holds the
 * host's copy of the document alive, and the docs are explicit that a leaked
 * handle can block later `getFileAsync` calls outright.
 */
async function slideSizeFromDocumentFile(): Promise<{ width: number; height: number } | null> {
  type Slice = { data: number[] | Uint8Array };
  type OfficeFile = {
    size: number;
    sliceCount: number;
    getSliceAsync(i: number, cb: (r: { status: string; value: Slice }) => void): void;
    closeAsync(cb?: () => void): void;
  };
  let file: OfficeFile | undefined;
  try {
    const doc = (
      Office as unknown as {
        context?: {
          document?: {
            getFileAsync(t: unknown, o: unknown, cb: (r: { status: string; value: OfficeFile }) => void): void;
          };
        };
      }
    ).context?.document;
    const fileType = (Office as unknown as { FileType?: { Compressed?: unknown } }).FileType?.Compressed;
    if (!doc?.getFileAsync || fileType === undefined) return null;

    file = await new Promise<OfficeFile>((resolve, reject) => {
      doc.getFileAsync(fileType, { sliceSize: 4194304 }, (r) =>
        r.status === "succeeded" ? resolve(r.value) : reject(new Error("getFileAsync failed")),
      );
    });

    const parts: Uint8Array[] = [];
    for (let i = 0; i < file.sliceCount; i++) {
      const slice = await new Promise<Slice>((resolve, reject) => {
        file!.getSliceAsync(i, (r) =>
          r.status === "succeeded" ? resolve(r.value) : reject(new Error("slice failed")),
        );
      });
      parts.push(slice.data instanceof Uint8Array ? slice.data : Uint8Array.from(slice.data));
    }
    const total = parts.reduce((n, p) => n + p.length, 0);
    const bytes = new Uint8Array(total);
    let at = 0;
    for (const p of parts) {
      bytes.set(p, at);
      at += p.length;
    }
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(bytes);
    const part = zip.file("ppt/presentation.xml");
    if (!part) return null;
    const emu = parseSlideSizeEmu(await part.async("string"));
    if (!emu) return null;
    return { width: emu.cx / EMU_PER_POINT, height: emu.cy / EMU_PER_POINT };
  } catch {
    return null;
  } finally {
    try {
      file?.closeAsync();
    } catch {
      /* nothing useful to do — the host will reclaim it */
    }
  }
}

/**
 * Append a blank slide and return its id, or null if it did not land.
 *
 * Exists so `withSlideDeselected` has somewhere to look on a deck that offers
 * nowhere. Verified the way `addSlides` verifies: a settled `slideCount()` in a
 * fresh context, because a count queued in the add's own sync reports the
 * PRE-add total on PowerPoint web, and because the web host drops `slides.add()`
 * outright under load. A scratch slide that did not land must report null
 * rather than hand back an id nobody can select or delete.
 *
 * The id is found by DIFFING the deck's ids before and after, from contexts
 * either side of the add, and then proved by resolving it once more.
 *
 * Both of those replaced weaker steps, and a real host is why. The first
 * version counted slides and read `getItemAt(count - 1).id` — "the last slide
 * is the one I just made". That is an assumption about where `slides.add()`
 * puts a slide, made by a function whose whole output is later passed to
 * `deleteSlideById`; if the host ever appends anywhere but the end, this hands
 * back one of the user's own slides to be deleted. A diff cannot make that
 * mistake: it names a slide that was not in the deck a moment ago.
 *
 * The second version stopped at reading the id, and the id was the problem. A
 * real PowerPoint on the web answered thirteen of fourteen host-probe questions
 * with "the host would not resolve the scratch slide" — the same id, read this
 * way, resolved once and then never again. Reading an id is not the same as
 * having a usable handle on a slide, and the only way to tell the two apart is
 * to go back and ask.
 *
 * **Position is back, and the warning above still stands — read both.** The
 * diff alone stopped being enough: this host renumbers an existing slide when
 * it appends one, so the deck grows by exactly ONE while TWO ids read as new,
 * eleven times across five rounds with no exception. The diff then names two
 * candidates, cannot choose, and refuses — which cost the probe's whole second
 * pass on 2026-08-10.
 *
 * So the fallback claims the LAST slide, and it is not the first version
 * returning. That one asked position ALONE, which is the thing this comment
 * warns about. This one may only claim a slide the diff has already proved was
 * not in the deck a moment ago, and only when the deck grew by exactly one — so
 * a host that appends somewhere other than the end hands back nothing rather
 * than one of the user's slides. `test/host-probe.test.ts` arms
 * `renumbersOnAdd` and `addsAtFront` together to prove that second half: drop
 * the freshness check and the probe deletes a slide it never added.
 *
 * `budgetMs` bounds the add for callers who cannot afford the default. The
 * probe is one: `READBACK_TIMEOUT_MS` is ninety seconds, sized for a
 * twenty-slide repair page, and a per-question budget of eight seconds means
 * nothing while the slide the question needs can take ninety on its own. One
 * question on 2026-08-10 took 95.6 seconds against that eight-second budget.
 * Measured, the choice is easy: successful adds in that run ran 0.21s to 4.0s
 * and failures took the full ninety, so the two are nowhere near each other.
 */
export async function addScratchSlide(
  budgetMs?: number,
  /**
   * Called for a slide this run added, could not use, and could not take back
   * out — i.e. one it has LEFT BEHIND.
   *
   * The clean-up sweep deletes by position and clamps at "no more than this run
   * added", so its notion of what this run added has to include every slide the
   * run put in the deck — including the ones it then refused to use. It did not:
   * a slide that lands and will not resolve is removed here, and when that
   * removal fails (which on this host is every time — delete-by-id cannot work)
   * the function returns null and the caller never hears the id. The 2026-08-11
   * round left exactly two slides behind for that reason: the deck grew by 70
   * while the run could account for 68, so the clamp correctly refused the
   * other two.
   *
   * Reported here rather than in the return value because the return value
   * means "an id you may use", and these are precisely the ids you may not.
   * They are still slides this run created and left in the deck, which is the
   * question the sweep is asking.
   */
  onAdded?: (id: string) => void,
  /**
   * Called for a slide that LANDED and could not be named.
   *
   * Separate from `onAdded` because there is no id to give — that is the whole
   * condition. The clean-up needs the COUNT regardless: its positional sweep
   * clamps at "no more than this run added", and a slide nobody counted is one
   * the clamp will refuse to remove. See the `!id` branch for the measurement
   * that found this.
   */
  onLeftBehindUnnamed?: () => void,
): Promise<string | null> {
  try {
    const before = await slideIds();
    // Bounded, and then ASKED — the same treatment the other three slide-adds
    // in this file get, and the only one that was missing it.
    //
    // office-js#1650 verbatim: "the promise doesn't resolve, although the slide
    // still gets added successfully". A raw `await context.sync()` here hangs
    // `runHostProbes` forever on a host that does that, and the whole promise
    // of the answer sheet is that it comes back from a host that is
    // misbehaving. Nothing is lost by not waiting: the before/after id diff
    // below is what names the slide, and it never trusted the promise anyway.
    await withTimeoutOrVerify(
      PowerPoint.run(async (context) => {
        const layoutId = await blankLayoutId(context);
        context.presentation.slides.add(layoutId ? { layoutId } : undefined);
        await context.sync();
      }),
      budgetMs ?? readbackTimeoutMs(),
      "adding a scratch slide",
    );
    const after = await slideIds();
    // A host that will not list its slides cannot be diffed. Nothing was
    // necessarily lost — the add may well have landed — but nothing here can
    // name what landed, and a scratch slide nobody can name is litter.
    if (!before || !after) {
      trace("host", "scratch slide: the deck would not list its slides", { before: !!before, after: !!after });
      return null;
    }
    const known = new Set(before);
    const fresh = after.filter((id) => !known.has(id));
    // Exactly one new slide, or give up. None means the host swallowed the add;
    // more than one USED to mean something else was adding slides at the same
    // time, and neither case leaves a slide this function can claim to own —
    // deleting one later would then delete the user's work.
    //
    // On this host "more than one" means something else entirely, and it is
    // common. Seven observations across four rounds, every single one the same
    // arithmetic: the deck grows by exactly ONE and TWO ids read as new.
    //
    //   before=20 after=21 fresh=2      before=37 after=38 fresh=2
    //   before=21 after=22 fresh=2      before=3  after=4  fresh=2
    //   before=18 after=19 fresh=2      before=19 after=20 fresh=2
    //   before=20 after=21 fresh=2
    //
    // One id has to leave the list for that to add up: appending a slide makes
    // this host RENUMBER an existing one. Nothing else is adding slides — the
    // probe is the only caller — so refusing both cost the question and left the
    // slide behind anyway. On 2026-08-10 it took the probe's whole second pass:
    // three attempts, three `fresh=2`, five questions never re-asked.
    //
    // So claim by POSITION rather than by set difference, and only in the case
    // the evidence actually covers. `slides.add()` APPENDS, so the slide we just
    // added is the last one; a renumbered pre-existing slide keeps the place it
    // always had. Requiring the deck to have grown by exactly one, and the last
    // id to be one of the new ones, is a stricter test than "pick a fresh id"
    // and strictly safer than the old rule in the direction that matters —
    // it never claims a slide that was already in the deck.
    let id = fresh.length === 1 ? fresh[0] : undefined;
    if (!id && after.length - before.length === 1) {
      const last = after[after.length - 1];
      if (!known.has(last)) {
        id = last;
        trace("host", "claimed the appended slide though the id list churned", {
          before: before.length,
          after: after.length,
          fresh: fresh.length,
        });
      }
    }
    if (!id) {
      // GREW BUT UNNAMEABLE IS NOT "DID NOT LAND", AND THE DIFFERENCE IS A LEAK.
      //
      // This branch is reached when the id diff cannot pick out which slide
      // arrived, and that happens for two completely different reasons while
      // reporting one of them for both: the add genuinely failed, or the add
      // worked and this host will not say which slide it was — the same
      // renumbering that makes delete-by-id useless here.
      //
      // In the second case a slide sits in the deck that the run has no record
      // of, because `onAdded` fires only on the branch below. The clean-up's
      // sweep clamps at "no more than this run added", so a slide it was never
      // told about is one it will not remove.
      //
      // Measured rather than reasoned: logging the deck either side of every
      // call in one run showed two of fifteen adds grow the deck by one, return
      // null, and report nothing. It is very likely also the already-recorded
      // "the deck grew by 70 while the run could account for 68" — the clamp was
      // right, and the count it clamped was two short.
      const landed = after.length > before.length;
      trace("host", landed ? "scratch slide landed but could not be named" : "scratch slide did not land", {
        before: before.length,
        after: after.length,
        fresh: fresh.length,
      });
      if (landed) onLeftBehindUnnamed?.();
      return null;
    }
    // NO settle here, and the reason is measured rather than argued.
    //
    // office-js#2903's workaround is to wait a couple of seconds after adding a
    // slide before touching it, and on 2026-08-07 this function did exactly
    // that. The next real-host round answered **1 of 25** questions, against 19
    // of 26 on the build before it: `addScratchSlide` returned null every time,
    // so every question came back `no-scratch-slide` and the only row with an
    // answer was the cleanup's own.
    //
    // So this host is not the host that issue describes. Its freshly-added
    // slide does not need time — it resolves once and is refused ever after
    // (the behaviour `deleteSlideById` and `SlideThunk` are already built
    // around), and spending two seconds first only moves the one resolution
    // later, by which point the id is gone. Waiting made it strictly worse.
    //
    // Left as a comment rather than a switched-off constant: a knob nobody
    // should turn is worse than a paragraph saying why.
    // Prove the id is worth handing out. Every caller's next move is to resolve
    // it in a context of its own, so do that once here, where the answer is
    // still cheap and a `null` is survivable.
    const usable = await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemOrNullObject(id);
      queueNullCheck(slide);
      await context.sync();
      return isLive(slide);
    });
    if (!usable) {
      trace("host", "scratch slide landed but its id will not resolve", { id });
      // Take it back out. It landed — the diff named it — so refusing to hand
      // out the id without also removing the slide would leave a blank one in
      // the deck on every attempt, which is the litter this whole function is
      // careful about everywhere else. `deleteSlideById` has a path that does
      // not need the id to resolve, which is exactly the case here.
      if (!(await deleteSlideById(id))) {
        trace("host", "could not remove the unusable scratch slide", { id });
        // Still in the deck, still this run's. Told to the caller here and only
        // here: a slide that was successfully taken back out is not left
        // behind, and counting it would make the clean-up owe a delete for a
        // slide already gone.
        onAdded?.(id);
      }
      return null;
    }
    return id;
  } catch {
    return null;
  }
}

/**
 * Every slide id in the deck, or undefined when the host would not say.
 *
 * Undefined rather than an empty array: "the deck has no slides" and "the deck
 * would not answer" are different facts, and `addScratchSlide` diffs two of
 * these — an unanswered read silently read as "empty" would make every existing
 * slide look brand new.
 */
async function slideIds(): Promise<string[] | undefined> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await boundedSync(context, "listing the deck's slides", READBACK_TIMEOUT_MS);
    const items = loadedItems(slides);
    if (!items) return undefined;
    const ids = items.map((s) => loadedValue(() => s.id));
    return ids.every((id): id is string => typeof id === "string" && !!id) ? ids : undefined;
  });
}

/** Delete one slide by id, best-effort. True when it is gone (or already was). */
export async function deleteSlideById(slideId: string): Promise<boolean> {
  try {
    const deleted = await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemOrNullObject(slideId);
      queueNullCheck(slide);
      await context.sync();
      // Not a verdict, only a failed handle: the host either says it is gone or
      // will not say, and neither is a reason to stop trying to clean up. The
      // deck's own slide list settles it below.
      if (!isLive(slide)) return false;
      slide.delete();
      await boundedSync(context, "deleting a slide", READBACK_TIMEOUT_MS);
      return true;
    });
    // Verify from a FRESH read, and answer from the deck rather than from the
    // fact that nothing threw.
    //
    // This used to `return true` straight after the delete's sync, which is not
    // a report — it is an assumption wearing one. A queued `delete()` the host
    // accepts and does not perform raises nothing, and this project has the
    // receipts for exactly that shape of failure: `slides.add()` silently
    // dropped under load, whole decks taken and never landed, tag writes
    // acknowledged and absent. The self-test's visibility scenario is the one
    // that made it visible: it borrows a slide and gives it back, and it could
    // not tell a returned slide from a leaked one.
    if (deleted && (await slideIsGone(slideId))) return true;
  } catch {
    /* fall through — a context that failed has not proved the slide is still there */
  }
  return deleteSlideByPosition(slideId);
}

/**
 * Whether the slide is confirmed absent — never "probably".
 *
 * The deck's own list is asked first because it is the stronger question: it is
 * one read of one collection, and a real host went on listing an id it had
 * stopped resolving individually. The single-object read is the fallback rather
 * than the primary for exactly that reason, but it stays, because a host that
 * will not describe its whole deck may still answer about one slide — and
 * dropping it would trade one blind spot for another.
 *
 * Both roads end at false when nobody will answer. A caller told "done" stops
 * looking, and every caller of this is cleaning something up.
 */
async function slideIsGone(slideId: string): Promise<boolean> {
  const ids = await slideIds().catch(() => undefined);
  if (ids) return !ids.includes(slideId);
  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemOrNullObject(slideId);
      queueNullCheck(slide);
      await context.sync();
      return loadedValue(() => slide.isNullObject) === true;
    });
  } catch {
    return false;
  }
}

/**
 * Delete `count` slides ending at the deck's end, by POSITION, highest first.
 *
 * The only clean-up that survives 2026-08-11's finding. Delete-by-id cannot
 * work on this host — `the deck still lists 0 of 62 of these ids`, so every
 * by-id delete reports "already gone" and removes nothing — and position needs
 * no id at all: `slides.add()` appends, so a run's own slides are the last N.
 *
 * WHICH slides is decided by `positionalSweepPlan`, away from any host call and
 * under test, because that decision is the safety and this function is only the
 * hands. It never chooses for itself.
 *
 * Highest index first, so removing one cannot shift the index of another still
 * to go. Each delete is queued into one batch and committed together: on a host
 * that stops answering mid-sweep this leaves the deck consistent, and the
 * caller's own before/after count is what says how much actually went.
 *
 * Returns how many the DECK lost, never how many deletes were issued. This file
 * has now twice shipped a clean-up that reported work it had not done, and both
 * times the deck was the thing that knew.
 */
export async function deleteTrailingSlides(from: number, count: number): Promise<number> {
  if (count <= 0 || from < 0) return 0;
  const before = (await slideIds().catch(() => undefined))?.length;
  try {
    await PowerPoint.run(async (context) => {
      for (let i = from + count - 1; i >= from; i--) {
        context.presentation.slides.getItemAt(i).delete();
      }
      await boundedSync(context, `deleting ${count} trailing slide(s)`, READBACK_TIMEOUT_MS);
    });
  } catch {
    /* fall through to the count — a throw does not say nothing was removed */
  }
  const after = (await slideIds().catch(() => undefined))?.length;
  if (before === undefined || after === undefined) return 0;
  return Math.max(0, before - after);
}

/**
 * Take out a slide the host will not hand back by id, by finding it in the
 * deck's own list and deleting it where it stands.
 *
 * Needed because `getItemOrNullObject` is not the last word on whether a slide
 * exists. A real PowerPoint on the web resolved a freshly-added slide's id once
 * and then refused it — while still listing that same id in
 * `slides.load("items/id")`. The slide was plainly there; only the lookup was
 * broken. `deleteSlideById` treated the refusal as "gone, nothing to do" and
 * reported success, so a host-probe run left fourteen blank slides in the deck
 * and said it had cleaned up after itself. `withSlideDeselected` has the same
 * cleanup on the same call, on the user's own deck.
 *
 * Positional deletes are how an add-in destroys someone's work, so the index is
 * re-read and matched against the id before anything is deleted. One extra
 * round trip is the price of never deleting the wrong slide.
 */
async function deleteSlideByPosition(slideId: string): Promise<boolean> {
  const ids = await slideIds().catch(() => undefined);
  if (!ids) return false;
  const index = ids.indexOf(slideId);
  // NOT in the deck's list, which this used to read as "already gone" — the one
  // reading that came from the deck rather than from a proxy that would not
  // answer. On this host that reading is false, and 2026-08-11 (`756682e`)
  // measured it rather than arguing it: of the 62 scratch slides a probe run
  // was deleting, `the deck still lists 0 of 62 of these ids`. Zero. Every one
  // of those deletes took this branch, returned true, and deleted nothing; the
  // deck went from 65 slides to 65 while the run reported a clean sweep.
  //
  // Both id lists come from the SAME `slideIds()` projection minutes apart —
  // the add captured its id from it, this reads it back — so "the id is not
  // there" cannot mean the slide is not there. It means the deck is answering
  // about the same slide under a different name.
  //
  // So: UNKNOWN, not gone. This makes the function claim less and delete no
  // more, which is the only safe direction on a call that removes slides from
  // someone's presentation. The caller's count check (`slidesActuallyReturned`)
  // is what turns the honest false into an honest report.
  if (index < 0) return false;
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(index);
      slide.load("id");
      await context.sync();
      if (loadedValue(() => slide.id) !== slideId) return;
      slide.delete();
      await boundedSync(context, "deleting a slide by position", READBACK_TIMEOUT_MS);
    });
  } catch {
    return false;
  }
  return slideIsGone(slideId);
}

/**
 * Run `fn` with the given slides NOT selected, then put the selection back.
 *
 * The live canvas is the bottleneck, not PowerPoint. This renderer already
 * knows it: `SHAPES_PER_SYNC` is 10 for the slide the user is looking at and
 * 40 for one they are not, because the web client repaints as shapes arrive
 * and stops answering somewhere past twenty. Updating a chart in place is
 * therefore the worst case in the whole add-in — a redraw of every shape,
 * onto the one slide guaranteed to be on screen. A real run died on the FIRST
 * batch: "did not respond while drawing shapes 1-10 of 39".
 *
 * Looking away is free. Select any other slide, redraw off-screen at the
 * bigger batch size, select back.
 *
 * A one-slide deck has nowhere to look, so this MAKES somewhere: a blank slide
 * appended for the duration and deleted in the `finally`. That case is not the
 * rare one it sounds like — it is a user building their first chart, and it was
 * the worst-served path in the add-in. Every edit ran on the live canvas at
 * batch 10, which is the exact configuration a real run died in ("did not
 * respond while drawing shapes 1-10 of 39"), and the deck with the fewest
 * slides to look away to is the one most likely to have only the chart on it.
 * Two extra undo entries is a real cost and the reason this is not done when a
 * slide to park on already exists; a redraw that kills the tab is a worse one.
 *
 * Entirely best-effort throughout: `setSelectedSlides` needs PowerPointApi 1.5,
 * the host can swallow the scratch add, and it can refuse the selection. Every
 * one of those falls back to running `fn` on the live canvas exactly as before,
 * having first put back anything it managed to add. `deselected` tells the
 * caller which happened, so it knows whether the wait budget it just used was
 * the off-screen one.
 */
export async function withSlideDeselected<T>(slideIds: string[], fn: (deselected: boolean) => Promise<T>): Promise<T> {
  if (!supports("1.5")) return fn(false);
  let restore: string[] | null = null;
  /** The slide this function parked the view on, so it can tell later whether
   *  the user has since moved away from it. */
  let parkedOn: string | null = null;
  /** A slide this function ADDED purely to have somewhere to look. Removed in
   *  the `finally`; never shown to `fn`, and never the user's own. */
  let scratchId: string | null = null;
  try {
    // Bounded, on the selection budget. This is the FIRST thing an in-place
    // update does, and it was two raw `PowerPoint.run`s sandwiching draw work
    // that is bounded per batch — so the one part of the update that could hang
    // forever was the part before anything was drawn. `showSlide` bounds the
    // identical `setSelectedSlides` call thirty lines away.
    //
    // A host that goes quiet here does not throw: the promise never settles,
    // this function never returns, `doInsert`'s `finally` never runs, and the
    // pane's busy counter stays up for the rest of the session — dead selection
    // banner, frozen status strip, an auto-update timer re-arming forever. The
    // failure is silence, so only a deadline can see it.
    const moved = await boundedRun(
      "looking away from the slide being redrawn",
      async (context) => {
        const presentation = context.presentation as unknown as {
          getSelectedSlides(): { items: { id: string }[]; load(p: string): void };
          setSelectedSlides(ids: string[]): void;
          slides: PowerPoint.SlideCollection;
        };
        const selected = presentation.getSelectedSlides();
        selected.load("items/id");
        const all = context.presentation.slides;
        all.load("items/id");
        await context.sync();
        const previous = selected.items.map((s) => s.id);
        const elsewhere = all.items.map((s) => s.id).find((id) => !slideIds.includes(id));
        // Nowhere to look YET. Report the selection anyway — the caller below
        // makes a slide to look at, and it still has to know where to put the
        // user back afterwards. Returning null here (as this did) threw that away.
        if (!elsewhere) return { previous, parkedOn: null };
        presentation.setSelectedSlides([elsewhere]);
        await context.sync();
        return { previous, parkedOn: elsewhere };
      },
      SELECTION_TIMEOUT_MS,
    );
    if (moved.parkedOn === null) {
      // Every slide in the deck is one we are about to draw on. Make a blank
      // one at the end, look at that instead, and take it away again after.
      scratchId = await addScratchSlide();
      // The host swallowed the add: nothing landed, nothing to clean up, and
      // the live canvas is still the only surface available.
      if (!scratchId) return await fn(false);
      if (!(await showSlide(scratchId))) {
        // It landed but the host will not look at it — an unusable scratch
        // slide is litter, so take it back out before falling through.
        await deleteSlideById(scratchId);
        scratchId = null;
        return await fn(false);
      }
      trace("host", "parked on a scratch slide", { scratchId });
    }
    restore = moved.previous;
    parkedOn = moved.parkedOn ?? scratchId;
  } catch {
    // A host that will not tell us what is selected, or will not change it —
    // draw on the live canvas as before rather than refusing to draw at all.
    // Anything already added has to come back out on the way past.
    if (scratchId) await deleteSlideById(scratchId);
    return fn(false);
  }
  try {
    return await fn(true);
  } finally {
    // Put the user back where they were — unless they have moved themselves.
    //
    // An off-screen redraw can run for tens of seconds, and the user is free
    // to click through the deck while it does. Restoring unconditionally then
    // snapped them back to wherever they happened to be standing when the
    // redraw started, discarding their navigation with no notice. So restore
    // only from the slide this function parked them on: if the view is still
    // there, nobody has touched it and putting them back is right; if it has
    // moved, the move was theirs and it wins.
    //
    // Failing to restore is not worth surfacing: they are one click from it.
    if (restore?.length) {
      // Bounded too, and it matters more here than the comment above suggests.
      // This sits in a `finally`, so a sync that never settles does not merely
      // skip the restore — it stops the `finally` from completing, which means
      // the caller's own `finally` never runs either. The pane's busy counter
      // is decremented in one of those. Failing to restore the view is cheap;
      // never returning from the update that did it is not.
      await boundedRun(
        "putting the view back",
        async (context) => {
          const presentation = context.presentation as unknown as {
            getSelectedSlides(): { items: { id: string }[]; load(p: string): void };
            setSelectedSlides(ids: string[]): void;
          };
          const now = presentation.getSelectedSlides();
          now.load("items/id");
          await context.sync();
          const ids = now.items.map((s) => s.id);
          if (ids.length !== 1 || ids[0] !== parkedOn) return;
          presentation.setSelectedSlides(restore);
          await context.sync();
        },
        SELECTION_TIMEOUT_MS,
      ).catch(() => {});
    }
    // Take the scratch slide back out — AFTER the restore above, never before.
    // Deleting the slide the view is currently on leaves the host to choose
    // where the user lands, which is the one outcome this whole function exists
    // to avoid. Restoring first means the delete happens off-view.
    //
    // Unconditional: it runs even when the restore was skipped because the user
    // navigated away themselves. A blank slide the add-in left at the end of
    // their deck is litter whether or not anyone is looking at it.
    if (scratchId) {
      const gone = await deleteSlideById(scratchId);
      if (!gone) {
        trace("host", "scratch slide could not be removed", { scratchId });
        console.warn(
          "PowerChart: could not remove the blank slide used to redraw off-screen — " +
            "it is the last slide in the deck and safe to delete by hand.",
        );
      }
    }
  }
}

/**
 * How long one demo item may take before the run stops trusting the host.
 *
 * Healthy items in this project's real runs land in 2-9 seconds. Sick ones
 * take 65-125. Thirty sits well clear of both, so this fires on a host that
 * is genuinely struggling rather than on one that is merely busy.
 */
const SICK_ITEM_MS = 30_000;

/** Batch size to use once the target slide is off-screen — see
 *  `SHAPES_PER_SYNC_OFFSCREEN`, which this deliberately mirrors. */
export const OFFSCREEN_BATCH = 40;

/**
 * The names of every shape on a slide, or `null` when the host would not say.
 *
 * Two answers, kept apart on purpose. "The slide holds these shapes" and "the
 * host did not answer for this slide's shapes" are different facts, and this
 * host produces the second constantly — the shape collection is the one thing
 * PowerPoint on the web reliably refuses (`the re-read before grouping came
 * back empty drew=24`), while the slide itself answers fine. A caller that
 * collapses them into a list reads a refusal as an empty slide.
 *
 * `null` covers all of it: the slide would not resolve, the collection did not
 * load, or the list could not be corroborated against the slide's own count.
 * That last one is what catches a collection answering SHORT without throwing
 * — observed on a stalled host as `shapesExpected=19 shapesSeen=15`, and
 * indistinguishable by name alone from a slide that really is that bare.
 * `getCount` is a scalar rather than a load, so it does not count against the
 * >50-item load ceiling (office-js#4272); it can throw outright on a host that
 * does not offer it, which costs the corroboration and nothing else.
 */
export async function slideShapeList(slideId: string): Promise<{ id: string; name: string }[] | null> {
  try {
    return await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemOrNullObject(slideId);
      slide.load("id");
      await context.sync();
      if (!isLive(slide)) return null;
      slide.shapes.load("items/id,items/name");
      let count: { value: number } | undefined;
      try {
        count = slide.shapes.getCount();
      } catch {
        count = undefined;
      }
      await context.sync();
      const items = loadedItems(slide.shapes);
      if (!items) return null;
      const read = items.map((s) => ({
        id: loadedValue(() => (s as unknown as { id: string }).id),
        name: loadedValue(() => (s as unknown as { name: string }).name),
      }));
      const n = count ? loadedValue(() => count.value) : undefined;
      if (typeof n !== "number" || n !== read.length) {
        trace("insert", "the host would not corroborate a slide's shapes", { slideId, seen: read.length, count: n });
        return null;
      }
      return read.map((s) => ({
        id: typeof s.id === "string" ? s.id : "",
        name: typeof s.name === "string" ? s.name : "",
      }));
    });
  } catch {
    return null;
  }
}

/** The names alone, for callers that do not need to tell one shape from another. */
export async function slideShapeNames(slideId: string): Promise<string[] | null> {
  return (await slideShapeList(slideId))?.map((s) => s.name) ?? null;
}

/**
 * True when the slide holds ONE shape and that shape is a PowerChart group —
 * i.e. replacing the whole slide loses nothing but the chart itself.
 *
 * The gate on `replaceSlideWithDeck`. A slide swap is the fastest possible
 * update (one host call instead of a redraw), but the new slide is a NEW
 * slide: speaker notes, transitions and anything else the old one carried do
 * not come with it. So it is only ever offered where there is demonstrably
 * nothing else to lose.
 *
 * An unreadable slide answers NO. This gate authorises DELETING the user's
 * slide — their logo, their title, their footnote, replaced by a generated one
 * that carries none of it and no speaker notes either — and it is consulted
 * only AFTER the host has already stalled, which is exactly when the shape
 * collection stops answering. It is the one gate in the file where being wrong
 * cannot be undone by anything but Ctrl-Z, so "I could not check" is not a
 * licence.
 *
 * An EMPTY slide counts — but only on a slide THIS RUN has not drawn on. The
 * question is "would replacing this slide lose anything the user put here", and
 * for a bare slide the answer is no; the gate used to insist on seeing exactly
 * one chart group, so the moment a failed redraw had deleted that group (and
 * its litter had been swept) the swap disqualified itself on the evidence of
 * its own damage, which left the one case this fallback exists for as the one
 * case it refused.
 *
 * The qualification is round `957aca0`, which broke the premise the empty case
 * rested on. `slideShapeNames` corroborates the collection against
 * `getCount()`, and that is what catches a hollow read — but on that round BOTH
 * answered zero for a slide holding a shape named `PowerChart`, so an empty
 * answer and an empty slide were byte-for-byte the same answer. The fake can
 * now produce it (`faults.slideReadsEmptyAfterPicture`) and no amount of
 * reading tells them apart.
 *
 * `shapesDrawnOn` is the one signal outside the collection, and it separates
 * them along the pathology's own documented line: this host does not list the
 * shapes A RUN JUST ADDED. So a zero on a slide this run has added shapes to is
 * not credible — it is the refusal, and believing it deletes the user's logo,
 * title and speaker notes for a slide that still holds them. A zero on a slide
 * this run never touched is credible, and that is exactly the case the empty
 * branch was added for: a chart whose delete landed and whose adds did not
 * leaves nothing drawn and nothing on the slide.
 *
 * Costs no host call — the count is bookkeeping this run already keeps.
 *
 * The count being NET is what keeps the fallback alive rather than merely
 * safe. `deleteShapesById` gives the slide's count back what it sweeps, and the
 * caller sweeps the stalled redraw's litter immediately before asking this
 * question — so a run that put shapes on the slide and took all of them off
 * again reads zero here and its empty read is believed, which is exactly the
 * "failed redraw, litter swept" case the empty branch was added for. A run
 * whose shapes are still sitting there unlisted reads non-zero and is refused.
 * Both halves come from the same subtraction.
 */
export async function slideHoldsOnlyChart(slideId: string): Promise<boolean> {
  const names = await slideShapeNames(slideId);
  if (!names) return false;
  if (names.length === 0) {
    const drew = shapesDrawnOn(slideId);
    if (drew > 0) {
      trace("insert", "not offering the slide swap: the slide read EMPTY after this run drew on it", {
        slideId,
        drew,
      });
      return false;
    }
    return true;
  }
  return names.length === 1 && names[0] === GROUP_NAME;
}

/**
 * Replace one slide with the single slide of a generated deck.
 *
 * Insert after the target, then delete the target — the order matters, since
 * deleting first would leave nothing to insert after and drop the new slide at
 * the front of the deck. Verified by slide count: the deck must end the same
 * size it started, or the swap half-happened and the caller is told so.
 *
 * Three outcomes, not two. "Nothing landed" and "the new slide landed but the
 * old one would not go" both used to answer `false`, and the caller treated
 * them identically — falling through to its picture fallback, which drew a
 * picture over the ORIGINAL chart while the new slide sat there holding the
 * same chart in native shapes. One stall, two charts, and a success message.
 * `duplicated` is that case named, so a caller can stop instead of making it
 * worse.
 */
export type SwapOutcome = "swapped" | "failed" | "duplicated";

export async function replaceSlideWithDeck(slideId: string, base64: string): Promise<SwapOutcome> {
  const before = await slideCount();
  try {
    // Silence is not failure here — office-js#1650 is this exact call not
    // answering while the slide lands anyway. The count check below is what
    // decides, and it used to be unreachable on the runs that needed it most.
    await withTimeoutOrVerify(
      PowerPoint.run(async (context) => {
        const presentation = context.presentation as unknown as {
          insertSlidesFromBase64(b64: string, opts?: { formatting?: string; targetSlideId?: string }): void;
        };
        presentation.insertSlidesFromBase64(base64, {
          formatting: "KeepSourceFormatting",
          targetSlideId: slideId,
        });
        await context.sync();
      }),
      DECK_INSERT_TIMEOUT_MS(1),
      "replacing a slide from a generated deck",
    );
  } catch {
    return "failed";
  }
  if ((await slideCount()) !== before + 1) return "failed";
  // Let the host settle before deleting. office-js#5022's only known workaround
  // is "a timer of 1-2 seconds between the shape.delete() and the next
  // context.sync()", and this is the same shape of thing one step up: a
  // structural change to the deck immediately after another one. A second is
  // cheap against a swap that already costs several, and against the failure it
  // avoids — a sync that never comes back and a duplicate slide left on screen.
  await settle();
  // Through `deleteSlideById`, which knows that a refused lookup is not a
  // verdict.
  //
  // This used to open its own context, call `getItemOrNullObject(slideId)`, and
  // throw "original slide is gone" the moment the flag was not `false` — on the
  // very id `insertSlidesFromBase64` had just accepted as `targetSlideId`, with
  // the count check above already proving the deck had grown by one. So the
  // original was demonstrably still there, and the user was told to sort out two
  // identical slides by hand. This host does exactly that: it resolved a slide's
  // id once and refused it ever after, while still listing it among the deck's
  // slides. `deleteSlideById` answers from that list and falls back to deleting
  // by position, which is the whole reason it was written.
  await withTimeoutOrVerify(
    deleteSlideById(slideId),
    DECK_INSERT_TIMEOUT_MS(1),
    "removing the slide a generated deck replaced",
  ).catch(() => {});
  // The deck's own answer, and now the only one — for the refusal case as much
  // as for the success case. A delete whose sync went unanswered still lands
  // often enough that reading the count is strictly better information than
  // reading the promise, and the old code returned "duplicated" out of its catch
  // WITHOUT re-counting, so a delete that did land was still reported as a
  // duplicate.
  if ((await slideCount()) === before) return "swapped";
  trace("insert", "slide swap left the original behind", { slideId });
  return "duplicated";
}

/**
 * Write the host's identity and capabilities into the trace.
 *
 * Called once when tracing starts. Every investigation in this project has
 * begun by guessing at these — which requirement sets the host admits to,
 * whether the one-call insert was even available — and guessing wrong at
 * least once. They cost one line each and answer the first three questions
 * anyone asks of a log.
 */
export function traceEnvironment(build: string): void {
  const d = (() => {
    try {
      return Office.context?.diagnostics;
    } catch {
      return undefined;
    }
  })();
  trace("host", "environment", {
    build,
    host: d?.host,
    platform: d?.platform,
    version: d?.version,
    // Whether the host will tell us what it actually ran. See `enableExtendedErrorLogging`.
    extendedErrors: enableExtendedErrorLogging(),
    // Every published set, not only the ones this add-in gates on. A log that
    // reports "1.5, 1.8" leaves a reader unable to tell a host that stops at
    // 1.5 from one that has 1.9 and simply is not asked — and the gap between
    // "what the host has" and "what we use" is where the next `setSelectedShapes`
    // is hiding. It cost months to find that one, and the log could have said.
    requirementSets: requirementSets(),
    canInsertSlidesFromBase64: canInsertSlidesFromBase64(),
    canInsertPicture: canInsertPicture(),
  });
  // Asynchronous, so it lands as its own trace line rather than holding up the
  // one above — and the SOURCE goes in it, because "960x540" read off PageSetup
  // and "960x540" assumed because nothing answered are the same two numbers
  // with completely different weight behind them. A run log that cannot tell
  // them apart cannot explain a mis-placed chart.
  void slideSize()
    .then((s) => trace("host", "slide size", { ...s }))
    .catch(() => {});
}

/**
 * Ask Office.js to record the statements it actually sent, not a summary.
 *
 * Every `debugInfo` in every real-host log this project owns ends the same way:
 * `"fullStatements":["Please enable config.extendedErrorLogging to see full
 * statements."]`. What survives without it is `surroundingStatements`, a
 * pretty-printed excerpt — and that excerpt is not enough to answer the
 * question the last four rounds have turned on.
 *
 * The question, concretely. Office.js rewrites a resolved
 * `getItemOrNullObject(id)` proxy's object path to `getItem(id)`, so a
 * statement reading `slides.getItem(...)` normally means "a handle held across
 * a sync" — the thing this whole file is built to avoid. But run 9 printed
 * exactly that for `settleAndTagChart`, whose slide and shape are both resolved
 * fresh inside a first sync of their own, which cannot have been rewritten.
 * Either the printer collapses the two forms under some condition, or that
 * batch was not the one it appeared to be. Both readings fit; neither can be
 * argued to a conclusion from an excerpt, and one of them says the settle is
 * repairable while the other says it never ran.
 *
 * So: stop reasoning, and make the host answer. `fullStatements` names the real
 * call. It costs one assignment and, bounded by `trimDebugInfo`, a few lines
 * per error in the log.
 *
 * Returns what happened, for the environment line — an older host without
 * `OfficeExtension.config` is a perfectly ordinary outcome, and a reader
 * looking at a log with no full statements in it should be able to tell "the
 * host would not" from "nobody asked".
 */
export function enableExtendedErrorLogging(): boolean {
  try {
    const cfg = (globalThis as { OfficeExtension?: { config?: { extendedErrorLogging?: boolean } } }).OfficeExtension
      ?.config;
    if (!cfg) return false;
    cfg.extendedErrorLogging = true;
    return cfg.extendedErrorLogging === true;
  } catch {
    return false;
  }
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
export const SHAPES_PER_SYNC = 10;

/**
 * Did this chart take more than one sync to draw?
 *
 * If it did, the proxies from every batch but the last are older than the
 * sync that is about to group them, and PowerPoint on the web answers
 * `InvalidParam passed to GetItem(id)` for exactly those. One re-read of the
 * slide's shapes costs a round trip; not doing it costs the chart its group
 * AND its config tag, which is the difference between a chart the pane can
 * re-open and a pile of shapes.
 */
function spansBatches(created: PowerPoint.Shape[], opts: InsertOptions): boolean {
  return created.length > (opts.shapesPerSync ?? SHAPES_PER_SYNC);
}

/**
 * Should this chart re-read the slide's shapes before it is grouped and tagged?
 *
 * **YES FOR ANY CHART THAT WILL BE GROUPED**, and until 2026-08-16 it was yes
 * only for a chart that spanned batches. That gate is what made a small chart
 * un-groupable on PowerPoint for the web, and the archive is unambiguous — 41
 * rounds, 537 draws:
 *
 *     spanned batches   452 draw(s),  353 grouped = 78%
 *     one batch only    214 draw(s),   49 grouped = 23%
 *
 * (First reported as 333/333 = 100%, which was a bug in the pooling function
 * rather than a fact: it dropped every `not grouping` decline from both arms.
 * The separation survived the correction; the absolutes did not.)
 *
 * A single-batch chart never refreshed, so `addGroup` got the raw `created`
 * proxies, and this host refuses those: `InvalidParam passed to GetItem(id)`,
 * 5010. The failed group then takes the tag with it — `target.tags` comes back
 * undefined, 155 times out of 155 across the archive, every one immediately
 * after a 5010 group.
 *
 * `spansBatches` was never wrong about its own case; it was too narrow. Its
 * reasoning is about proxies aged across a sync boundary, and it missed that
 * this host refuses a creation proxy for `addGroup` **whatever its age**.
 *
 * **WHY THE OLD GATE EXISTED, and why the objection has weakened.** Asking for a
 * re-read a chart did not need was "a way to LOSE a group, not gain one",
 * because this host answers a re-read short or empty and an empty answer used to
 * mean "group nothing". Two things protect that now: a short or empty answer is
 * asked AGAIN after a settle delay (`REREAD_RETRY_MS`), and if it is still empty
 * a chart whose proxies never crossed a sync falls back to `created` rather than
 * declining (`chooseGroupMembers`). The widening cannot leave a small chart
 * worse off than it was.
 *
 * **The cost is a round trip per single-batch chart on the live insert path**,
 * which is the path a user waits on. Taken deliberately: a chart that cannot be
 * grouped loses its config three times in four, and an insert 0.1s slower beats
 * one the pane cannot re-open.
 *
 * `spansBatches` stays in the OR for the chart that is NOT groupable but still
 * crosses a sync boundary — a single degraded picture is the case — because the
 * stale-proxy trap is a property of using a proxy across a sync, and tagging
 * does that too.
 */
export function needsPreGroupRefresh(created: PowerPoint.Shape[], opts: InsertOptions, hasPicture = false): boolean {
  // The same test `groupAndTagAll` uses to decide what is groupable, so the two
  // cannot drift apart — a chart that will be grouped is exactly the chart that
  // needs fresh handles to be grouped WITH.
  const willGroup = opts.group !== false && created.length > 1;
  return willGroup || spansBatches(created, opts) || hasPicture;
}

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
  /**
   * Filled with every shape this call creates, as it creates them.
   *
   * The return value only exists if the whole render succeeded, and a render
   * that stalls half way has still COMMITTED every batch before the one that
   * died — those shapes are on the slide and the caller has no other handle on
   * them. An update that deleted the old chart first then left that partial
   * chart behind as unreachable litter, under whatever it drew next. Passing a
   * sink is how the caller gets to clean up what it started.
   */
  sink?: PowerPoint.Shape[],
): Promise<PowerPoint.Shape[]> {
  if (wantsPicture(opts, scene)) {
    // Report the picture as a single unit of work so a caller's progress bar
    // doesn't sit at zero and then jump — the picture IS the whole chart.
    onBatch?.(1, 1);
    const picture = await renderPictureShape(context, getSlide, scene, opts);
    if (picture.length) {
      sink?.push(...picture);
      return picture;
    }
    // Refused: fall through and draw the nodes instead, in this same context.
  }
  const { dx: left, dy: top } = frameOrigin(opts);
  const batchSize = opts.shapesPerSync ?? SHAPES_PER_SYNC;
  // The sink IS the accumulator when one was passed, so a throw leaves the
  // caller holding exactly what got drawn — no copying, nothing to keep in step.
  const created: PowerPoint.Shape[] = sink ?? [];
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
  // Which batch this is. `created` can arrive as a non-empty sink, so its length
  // is not a batch counter.
  let batchNo = 0;
  /** The previous batch's sync duration — the number differencing could not give. */
  let prevBatchMs: number | undefined;
  /**
   * How much THIS call has banked, and where — so the sentinel's first-batch
   * shapes can be moved onto the real slide key once the host names it, without
   * moving another draw's shapes that happen to sit under the same sentinel.
   */
  let bankedHere = 0;
  let bankedUnder: string | undefined;
  while (s < steps.length) {
    // The stop check, and the only place it can honestly go: batches already
    // committed are on the slide and a sync in flight cannot be recalled, so
    // "stop" means "queue nothing further". Throwing here rather than returning
    // early is deliberate — every caller already has a path for a render that
    // did not finish (it sweeps what landed), and a stop leaves exactly the
    // same partial chart a stall does. Returning `created` would instead report
    // a half-drawn chart as a successful render, group it, and tag it.
    throwIfStopped();
    // Fresh slide proxy per batch: a proxy held across the previous sync may have
    // been rewritten to an unusable getItem(id) — see SlideThunk.
    const shapes = getSlide().shapes;
    batchNo++;
    const before = created.length;
    // Always draw at least one step, then keep going while the batch stays under
    // budget — a single indivisible node (a 16-triangle wedge fan) is its own
    // floor, and drawing it alone is the smallest batch available.
    do {
      created.push(...steps[s++](shapes));
    } while (s < steps.length && created.length - before < batchSize);
    // Read each shape's id back on the batch's OWN sync — free, because the
    // sync is happening anyway, and it is what lets a later re-fetch find these
    // exact shapes again. Without ids the only way to identify a chart's shapes
    // in a fresh read is "the last N on the slide", which is true of a blank
    // slide this run added and false of the one the user is looking at.
    //
    // EVERY shape, the anchor included. Running this one shape behind — to leave
    // the anchor unresolved for the tag write — was built, shipped, and measured
    // at no effect on the host it was for; see the note beside `CHART_ORIGIN_TAG`
    // before rebuilding it.
    for (let k = before; k < created.length; k++) created[k].load("id");
    // TRIED HERE AND REVERTED: writing the config tag on `created[0]` in this
    // same sync. Recorded because the reasoning behind it is sound and somebody
    // will have it again.
    //
    // The evidence for it is real. Every refused tag write in the 2026-08-07 log
    // is a proxy several batches old — `shapes.getItem("27") /* originally
    // addTextBox(...) */`, where 27 is the title box drawn in the FIRST batch and
    // written to in the LAST — and the probe answers the other half:
    // `tags-on-fresh-shape: yes`. A shape has its tags collection the moment it
    // is added; what it does not survive is the round trip that rewrites its
    // object path. So this sync looks like the one window the web host honours.
    //
    // What it breaks is the case where grouping WORKS. The group is tagged as
    // well, and both tags are then findable, so the deck scan counts one chart
    // twice and an in-place update can pick the wrong one — eleven tests, and
    // among them "follows a chart the user has DRAGGED" going 277pt wrong. A
    // chart that teleports on every desktop edit is a worse bug than one that is
    // not re-editable on the web.
    //
    // The fix this wants is a SECOND key that only the recovery path reads, or
    // dropping the anchor's tag once the group's lands. Both are real changes
    // with their own guards, not a line in this loop.
    sent += created.length - before;
    const upTo = Math.min(sent, total);
    // Reported BEFORE the sync, and deliberately: the sync is where a bad host
    // stops answering, so this is the number that has to be on screen WHILE we
    // wait. Reporting after would leave the pane naming the previous phase and
    // blaming the wrong one for the stall.
    onBatch?.(upTo, total);
    // The idle gap, on the FIRST batch only — the baseline the stall record has
    // no meaning without. Every draw stall on record is a first batch, and the
    // first one to name its predecessor said `rasterising a slide, idleMs: 1`.
    // Whether either number is remarkable depends entirely on what the batches
    // that SURVIVE report.
    //
    // `idleMs` got its baseline first and DIED of it — survivors span 1ms to
    // 2182ms, two of them at exactly 1ms, so the gap separates nothing. The
    // predecessor's NAME was then left in the same condition for two more
    // rounds: recorded on stalls, never on successes. Round 9 made that
    // untenable by producing two stalls naming two DIFFERENT calls (`selecting
    // a shape` and `rasterising a slide`) while thirteen draws in the same round
    // survived and said nothing about what they followed.
    //
    // So both halves are recorded here, together, and one round decides whether
    // the identity of the preceding call discriminates or goes the way the gap
    // did. First batch only: a later batch's predecessor is always the batch
    // before it, which says nothing and would cost three lines a chart.
    //
    // ISSUED, not committed — and the name is load-bearing. This line was called
    // `batch committed` for as long as it existed, while being written one
    // statement BEFORE the sync it describes, so every stall on record left a
    // line claiming the batch it killed had committed. Two hand analyses of the
    // same data died on it: one paired the lines with draws and reported 0
    // stalls in 32, the other counted them as successes and manufactured a 6x
    // rasterise effect that was not there. Both are written up in
    // `scripts/triage.mjs`, which compensates by ignoring this message entirely
    // and reading `gave up waiting` instead. A comment in the reader is not the
    // fix; a truthful name in the writer is. The commit has no line of its own
    // on purpose — the next batch's `issued` implies it, and the last one's is
    // the draw returning.
    // What this run had already put on THIS slide before the batch, and how long
    // the previous batch took. Both recorded on every batch, stalled or not,
    // because a value written down only when something goes wrong cannot be
    // compared against anything — the mistake this file has now made four times.
    //
    // `onSlide` is the input to the quadratic per-slide cost and has never been
    // recorded; `prevBatchMs` is the batch duration, which until now had to be
    // differenced out of consecutive `issued` timestamps and so silently
    // included every bit of inter-batch work.
    // Keyed on the slide when the caller named one, and on a sentinel when it
    // did not — because `opts.slideId` is absent on most draws and this field
    // was populated on SEVEN of forty-six batches in its first real round. A
    // number that is missing five times out of six cannot answer the question
    // it was added for, which is the cost curve's own input.
    //
    // `(visible)` is honest rather than convenient: an unnamed target is the
    // slide the user is looking at, which is exactly where the repeated-draw
    // case lives (editing a chart in place redraws every shape onto it). The
    // key is emitted beside the count, so a reader can see when a total is
    // accumulating across what may be more than one slide instead of having to
    // assume it is not.
    //
    // That escape hatch earned its keep immediately, and then had to be closed.
    // The round of `4feb5be` redrew eight charts on eight DIFFERENT slides
    // through `updateChartsInSlides`, which never filled `slideId` in, so every
    // one of them keyed on `(visible)` and the counter climbed to 260 — a
    // number that describes no slide in the deck. `onSlideKey` said so on every
    // line, which is the only reason it was caught rather than plotted; the
    // update path names the slide now (`it.target.slideId`), so the counter is
    // per-slide there as it always claimed to be.
    //
    // …and the INSERT path was still on the sentinel, which is worse than
    // untidy. `slideHoldsOnlyChart` reads this counter to decide whether an
    // empty read of a slide is believable, and it authorises deleting the
    // user's slide. A chart inserted the ordinary way — the pane's Insert
    // button, no id — banked its twenty-four shapes under `(visible)`, so
    // `shapesDrawnOn(realId)` answered ZERO for a slide this run had just
    // filled, and the guard would have believed the host's empty read on
    // exactly the slide it was written to refuse. Round `393e6e4` is where that
    // showed: every batch of `insert onto a slide that already has content`
    // keyed on the sentinel while every other scenario named its slide.
    //
    // The id is free. `insertSceneIntoSlide` already queues `slide.load("id")`
    // before the first batch, so from the second batch on the host has answered
    // it — `slideKeyFor` reads it through `loadedValue`, because a caller that
    // did not queue the load throws `PropertyNotLoaded` on the property access
    // rather than returning undefined. The first batch is retagged rather than
    // lost: its shapes are moved onto the real key the moment one is known, so
    // the total is right even though the first line is written before the host
    // could say.
    const slideKey = slideKeyFor(opts, getSlide);
    if (bankedUnder !== undefined && bankedUnder !== slideKey && bankedHere > 0) {
      // Only this call's own shapes move. The sentinel is shared, so taking its
      // whole total would steal another draw's count.
      forgetShapesDrawnOn(bankedUnder, bankedHere);
      shapesDrawnOnSlide.set(slideKey, (shapesDrawnOnSlide.get(slideKey) ?? 0) + bankedHere);
      trace("draw", "moved this draw's shape count onto the slide the host finally named", {
        from: bankedUnder,
        to: slideKey,
        shapes: bankedHere,
      });
    }
    trace("draw", "batch issued", {
      upTo,
      total,
      onSlide: shapesDrawnOnSlide.get(slideKey) ?? 0,
      onSlideKey: slideKey,
      ...(prevBatchMs === undefined ? {} : { prevBatchMs }),
      ...(batchNo === 1
        ? {
            idleMs: Math.round(idleSinceLastAnswer()),
            afterAnswering: lastAnsweredCall ?? "nothing yet",
            afterAnsweringMs: lastAnsweredMs,
          }
        : {}),
    });
    const batchStarted = Date.now();
    // Budget per BATCH, not per chart: a stalled host must still be caught, but
    // the limit now measures a batch we know the host can swallow.
    await withTimeout(
      context.sync(),
      BATCH_TIMEOUT_MS,
      `drawing shapes ${upTo - (created.length - before) + 1}-${upTo} of ${total}`,
    );
    prevBatchMs = Date.now() - batchStarted;
    shapesDrawnOnSlide.set(slideKey, (shapesDrawnOnSlide.get(slideKey) ?? 0) + (created.length - before));
    bankedHere += created.length - before;
    bankedUnder = slideKey;
  }
  return created;
}

/** One chart's committed shapes, awaiting grouping and tagging. */
interface Grouping {
  getSlide: SlideThunk;
  created: PowerPoint.Shape[];
  opts: InsertOptions;
  /**
   * Re-load the slide's shape collection right before addGroup and re-resolve
   * our shapes out of it BY ID. The Shape proxies returned by earlier syncs'
   * add*() calls have their object paths rewritten to getItem(id) by the time
   * this sync runs, and the web host can't round-trip those ids —
   * addGroup(theseStaleProxies) silently drops the group.
   *
   * THIS IS NOT AN OWNERSHIP CLAIM, and it used to read as one. The doc here
   * promised "the caller can guarantee the target N shapes are the last N on
   * the slide", which described the ORIGINAL positional rule; the matcher has
   * been id-based since, and no call site ever set this field from ownership —
   * all three set it from `spansBatches()`, i.e. "this chart's proxies crossed
   * a sync boundary". Contract and use had diverged, and anyone reaching for a
   * real ownership guarantee would have found this field, read that sentence,
   * and been wrong. The honest reading is the only one the code implements:
   *
   *   set it when the created proxies may be stale — nothing more.
   *
   * There is deliberately no ownership flag to reach for instead. Only the
   * self-test and the demo path add their own blank slides; the pane's Insert
   * and the deck-wide update both draw onto the USER's slides, so a rule gated
   * on ownership would improve the measurement and not the product.
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
/**
 * Where a chart's tagged shape ended up — PLAIN VALUES, read off the proxy
 * while it was still readable.
 *
 * Never the proxy itself, and that distinction is the bug this type exists to
 * make impossible. `groupAndTagAll` used to hand back `PowerPoint.Shape`, whose
 * `id`/`left`/`top` are only populated by the `load` queued in the TAGGING
 * sync. When that sync failed — `InvalidParam passed to GetItem(id)`, the web
 * host refusing a stale proxy, which is a documented and CAUGHT outcome here —
 * the catch logged it and the function returned the proxy anyway. Every caller
 * then read `t.left` off a shape that had never loaded and threw
 * `PropertyNotLoaded` at `Shape.left`, out of a call whose charts were on the
 * slide and whose only real loss was re-editability.
 *
 * So a deck-wide rescale in PowerPoint on the web reported a crash for work
 * that had, on the slide, succeeded — and the same read is on the everyday
 * insert path, one sync away from the same fate.
 */
interface TargetRef {
  id: string;
  left: number;
  top: number;
}

/** A tagged shape's position, or undefined when the host never answered for it. */
function targetRef(shape: PowerPoint.Shape | undefined): TargetRef | undefined {
  if (!shape) return undefined;
  const id = loadedValue(() => shape.id);
  const left = loadedValue(() => shape.left);
  const top = loadedValue(() => shape.top);
  if (typeof id !== "string" || !id) return undefined;
  if (!Number.isFinite(left) || !Number.isFinite(top)) return undefined;
  return { id, left: left as number, top: top as number };
}

/**
 * Which handles `addGroup` may be given — or none, meaning do not group.
 *
 * A shape proxy carries its PARENT's object path, and that is the whole
 * problem. `freshMembers` comes out of `collections[k].items`, whose parent is
 * the slide handle the re-read sync resolved — so Office.js has since rewritten
 * that path to `slides.getItem(id)`, which PowerPoint on the web refuses for a
 * slide this run added. `it.created` carries the same wound for the same
 * reason: those proxies were made in the drawing batches, off a handle that has
 * been resolved since.
 *
 * A real host listed the statements it refused, which is how the shape of it is
 * known at all:
 *
 *   var itemOrNullObject = slides.getItemOrNullObject(...);   // fresh
 *   var slide            = slides.getItem(...);               // rewritten
 *   var shapes1          = slide.shapes;
 *   var shape            = shapes1.getItem(...);              // refused, 5010
 *
 * Only an ID may cross a sync. Where every member can be named, they are
 * re-resolved off a handle taken in the grouping batch and grouping is safe.
 * Where they cannot, the answer is to group NOTHING — because handing a refused
 * handle to `addGroup` does not merely fail to group, it throws, and the throw
 * takes the batch's tagging with it. Five charts in one run lost their group
 * AND their config that way, where grouping nothing loses at most one chart's.
 *
 * THAT COMPARISON USED TO END "…than an ungrouped chart that is still
 * re-editable", and the archive refutes the second half: an ungrouped chart
 * keeps its config about one time in three, not as a rule (97 charts on an
 * established slide, 96 grouped; 84 on a freshly added one, 1 grouped —
 * `npm run rounds`). Grouping nothing is still the better of the two, because a
 * throw costs every chart in the batch rather than one, but it is a choice
 * between two losses and not a safe fallback. The way out is upstream: the
 * members cannot be named because the slide was added by this run, and that is
 * where a fix belongs. See `docs/BACKLOG.md`.
 *
 * The one case that may still use `created`: a chart that never asked for a
 * refresh. Nothing has been re-resolved behind it, so `created` is the only
 * handle there has ever been — the small-chart path that has always worked.
 */
export type GroupMembers = { use: "ids"; ids: string[] } | { use: "created" } | { use: "none" };

export function chooseGroupMembers(o: {
  /** Ids read off the re-read members, or undefined when the re-read produced nothing. */
  refreshedIds?: (string | undefined)[];
  askedForRefresh: boolean;
}): GroupMembers {
  if (o.refreshedIds) {
    const named = o.refreshedIds.filter((id): id is string => typeof id === "string" && !!id);
    // Every member or none. A partial list would group part of the chart and
    // leave the rest loose beside it, which is worse than not grouping at all —
    // the loose remainder is then deleted by the next in-place update.
    return named.length === o.refreshedIds.length && named.length > 0 ? { use: "ids", ids: named } : { use: "none" };
  }
  // Asked for a refresh and did not get one. That is the answer of a host which
  // has just refused to list the slide's shapes at all — this one answered
  // `shapes-items-count-honest` with `short-0`, twenty-four drawn and none
  // reported — and it is exactly the host that will refuse `created` too.
  //
  // FALLING BACK TO `created` HERE WAS TRIED ON 2026-08-16 AND IS WRONG, twice
  // over. Every groupable chart now asks for the refresh, so it seemed only fair
  // that a small chart whose proxies never crossed a sync should drop back to
  // what it used before rather than be made worse by asking.
  //
  // It cannot work, and it would hurt if it did. **The re-read itself costs a
  // sync**, so by the time `addGroup` runs those proxies are a sync old and this
  // host rejects them anyway — `strictGroup` in the fake models exactly that,
  // and it is not a modelling artifact: on the real host a single-batch chart
  // grouping through same-sync creation proxies succeeded 49 times in 204.
  //
  // And a doomed attempt is worse than none. A refused `addGroup` does not just
  // fail — it takes the tag with it: `target.tags` comes back undefined 155
  // times out of 155 across the archive, every one immediately after a 5010
  // group. An ungrouped chart still gets a tag write attempted; a failed group
  // loses the group AND the tag. So declining is strictly better than trying.
  return o.askedForRefresh ? { use: "none" } : { use: "created" };
}

async function groupAndTagAll(
  context: PowerPoint.RequestContext,
  items: Grouping[],
): Promise<{ target?: TargetRef; partIds?: string[]; grouped?: boolean; tagged?: boolean }[]> {
  const tagTargets = items.map((it) => it.created[0] as PowerPoint.Shape | undefined);
  /**
   * WHERE each tag target came from, which is the one thing the failure trace
   * could never say.
   *
   * Six rounds lost half a deck of config tags to `InvalidParam passed to
   * GetItem(id)` here, and every one of them recorded the error and the count
   * and nothing about the handle it was written through. That is the difference
   * between the four routes below, and it decides the fix:
   *
   *   created    the proxy that drew the shape, never loaded — the host takes
   *              writes through this one however old it is
   *              (`tag-the-creation-proxy-a-sync-later: yes`, four rounds)
   *   refreshed  the pre-grouping re-read, which a `load()` has RESOLVED, and
   *              which Office.js has therefore rewritten to `shapes.getItem(id)`
   *   group      the group made in the grouping batch, also never loaded
   *   by-id      an explicit `getItemOrNullObject(id)` lookup
   *
   * `refreshed` and `by-id` are the resolved pair and are what this host
   * refuses; `created` and `group` are not. A round that says which was used
   * turns "the tag write failed" into "the tag write failed THROUGH A RESOLVED
   * HANDLE", which is a fix rather than another investigation.
   */
  const targetFrom = items.map(() => "created");
  /** `["created","refreshed","created"]` → `"created×2, refreshed×1"`. */
  const countBy = (xs: string[]) =>
    [...xs.reduce((m, x) => m.set(x, (m.get(x) ?? 0) + 1), new Map<string, number>())]
      .map(([k, n]) => `${k}×${n}`)
      .join(", ");
  // Which charts actually ended up as one shape. The rest hang everything the
  // group would have carried off their first shape instead — see below.
  const grouped = new Set<number>();
  // Grouping is PowerPointApi 1.8+; skip entirely where unsupported.
  const groupable = items
    .map((it, i) => ({ it, i }))
    .filter(({ it }) => it.opts.group !== false && it.created.length > 1);
  // Re-load shape collections for items that asked to refresh — one sync for
  // all of them, then re-resolve our own shapes out of each collection BY ID.
  // See Grouping.refreshShapes: bypasses the stale-proxy trap the web host
  // trips on.
  //
  // NOT just the groupable ones. A degraded picture is a single shape, so it is
  // not groupable, and it went to the tag step holding the proxy it was created
  // with. The web host then rejected that proxy outright:
  //
  //   InvalidParam passed to GetItem(id) | code=5010
  //   errorLocation: ShapeCollection.getItem
  //   statement: var shape = shapes.getItem(...); var tags = shape.tags;
  //
  // Twenty-eight times in one 38-item run, which is where that run's missing
  // tags went. The stale-proxy trap is not a property of GROUPING; it is a
  // property of using a proxy across a sync boundary, and tagging does that
  // too.
  // Every item that ASKED to refresh, not just the groupable ones — but only
  // those that asked, because a re-read costs a round trip and this host answers
  // it short or empty often enough that asking for one a chart does not need is
  // a way to LOSE a group, not gain one (see the `needsRefresh` call site).
  //
  // This filter used to be justified by "the last N shapes on the slide is only
  // true of a slide this run added blank" — the positional rule, gone since the
  // matcher went id-based below. Matching by id is safe on any slide, so what
  // remains is a cost argument, not a correctness one.
  const refresher = items.map((it, i) => ({ it, i })).filter(({ it }) => it.refreshShapes);
  const freshMembers = new Map<number, PowerPoint.Shape[]>();
  if (refresher.length) {
    try {
      // Inside the try, all of it. Resolving the collections and queueing their
      // loads can throw SYNCHRONOUSLY — `getSlide()` reaches into the host, and
      // a host that hands back something without a `.load` (seen on the web as
      // "e.load is not a function") took down the whole PowerPoint.run from out
      // here. That failed the entire update of a chart whose shapes had already
      // committed: Same Scale across the deck reported one thrown TypeError for
      // work that was, on the slide, done. Everything this block does is a
      // best-effort refresh — the catch below already says so — so it all
      // belongs where the catch can reach it.
      // ASKED TWICE, and the second time after a pause. See `REREAD_RETRY_MS`:
      // this host does not populate a freshly materialised slide's shape
      // collection straight away, and a chart drawn onto a slide this run just
      // added is grouped 1% of the time against 99% on a slide that already had
      // shapes. `pending` carries the charts whose answer was short or empty;
      // anything that matched is out of the loop after the first pass and pays
      // nothing.
      //
      // A FRESH slide handle per attempt, from `it.getSlide()` — not the
      // collection from last time. Office.js will have rewritten the previous
      // attempt's slide path to `slides.getItem(id)` by now, and this host
      // refuses shapes hanging off that; the draw loop takes a fresh proxy per
      // batch for the same reason (see SlideThunk).
      let pending = refresher;
      for (let attempt = 0; pending.length && attempt <= REREAD_ATTEMPTS; attempt++) {
        const lastAttempt = attempt === REREAD_ATTEMPTS;
        if (attempt > 0) {
          trace("group", "re-reading the slide's shapes again after a settle delay", {
            charts: pending.length,
            waitedMs: REREAD_RETRY_MS,
            slides: countBy(pending.map(({ it }) => slideKeyFor(it.opts, it.getSlide))),
          });
          if (REREAD_RETRY_MS > 0) await new Promise((r) => setTimeout(r, REREAD_RETRY_MS));
        }
        const collections = pending.map(({ it }) => it.getSlide().shapes);
        const retry: typeof pending = [];
        // A bare `items`, and it is the only collection load in this file that does
        // not name its properties. That was changed to `items/id` on 2026-08-07 on
        // the reasoning that `id` is the only property this block reads and
        // Office.js populates what was asked for — and it was changed BACK, because
        // the reasoning rests on a fact nobody here has: whether `load("items")`
        // populates the items' scalar properties or only the collection.
        //
        // The suite answered a narrower question and answered it clearly. Under
        // `applyWebProfile` the switch broke `still gets its degraded pictures
        // tagged`: `hollowReads` models a short collection read and keys on
        // `items/id`, so widening this projection pulled the grouping re-read into
        // a blindness the bare form had been escaping — the re-read came back
        // empty, the single-shape re-fetch below never ran, and the degraded
        // picture's tag went through the stale proxy the host refuses. That is not
        // the fake being unfair: a real host that reads short does not care which
        // projection was asked for.
        //
        // So the change trades a guess about `load("items")` for a measured
        // regression, which is the wrong way round. Ask the host instead — the
        // three traces below now separate "came back empty" from "threw" from
        // "matched nothing", and a run log that says which will settle it.
        for (const c of collections) c.load("items/id");
        await boundedSync(
          context,
          attempt === 0
            ? "re-reading the slide's shapes before grouping"
            : "re-reading the slide's shapes after a settle delay",
        );
        pending.forEach((entry, k) => {
          const { it, i } = entry;
          const items = collections[k].items;
          // A slide we JUST drew onto cannot be empty, so an empty answer is the
          // host reading short — the `hollowReads` behaviour, seen for real when
          // one readback asked about 19 shapes and was told 3. Worth its own line
          // because the alternative readings of `refreshed=0` want different
          // fixes, and from a run log they are indistinguishable.
          if (!items?.length && it.created.length) {
            // COUNTED AND TRACED ONLY ONCE THE RETRY HAS RUN OUT. An empty first
            // answer that a settled re-read then fills in is not a fault the
            // round should carry — counting it would report the host as failing
            // on a chart that came out fine, and `emptyReReads` is read as a
            // per-round friction number.
            if (!lastAttempt) {
              retry.push(entry);
              return;
            }
            hostFriction.emptyReReads += 1;
            trace("group", "the re-read before grouping came back empty", {
              index: i,
              drew: it.created.length,
              // WHICH SYNC OF THIS CONTEXT. See `syncsPerContext`: the deck-wide
              // update runs every chart through one context and the re-read decays
              // chart by chart, so the count is the x-axis that says whether the
              // context is what wears out.
              contextSyncs: syncsOf(context),
              // AND THAT THE PAUSE DID NOT SAVE IT, which is the one thing the
              // line could not say before. An empty read here has now survived a
              // settle delay, so it is not the host still catching up.
              afterRetry: true,
            });
            return;
          }
          // By ID, which is exact and works on any slide. The old rule was "the
          // chart's shapes are the LAST N on the slide" — true of a blank slide
          // this run added, and false of the slide the user is looking at, which
          // is why the ordinary insert path could not use this at all and lost
          // its grouping and its config tag on every chart big enough to need
          // more than one batch.
          const byId = new Map<string, PowerPoint.Shape>();
          for (const sh of items) if (sh.id) byId.set(sh.id, sh);
          const matched = it.created.map((sh) => (sh.id ? byId.get(sh.id) : undefined)).filter(Boolean);
          if (matched.length === it.created.length) {
            // Same order as `created`, so index 0 stays the chart's anchor and
            // `ungroupedFallback`'s "everything after it is a part" holds.
            freshMembers.set(i, matched as PowerPoint.Shape[]);
          } else if (matched.length) {
            // ASK AGAIN FIRST. A partial match is the OTHER shape of the same
            // fault the empty read is — chart 4 of 8 matched 20 of its 24 shapes
            // in every round for five rounds running, which is a read that has not
            // finished rather than four shapes that are missing. Only once a
            // settled re-read has said the same thing twice is the trade below a
            // real trade.
            if (!lastAttempt) {
              retry.push(entry);
              return;
            }
            // A PARTIAL match is thrown away, and this used to be kept.
            //
            // The argument for keeping it was that every shape in it is provably
            // ours, where the positional rule below is a guess. True, and beside
            // the point: what it produces is a chart split into a group plus a
            // remainder that does not move with it. The user drags the chart and
            // leaves its baseline behind — and it LOOKS like one object, so
            // nothing warns them. `4feb5be` left exactly that on a real slide,
            // `grouped the chart's shapes charts=1 partial=1 left=0:4`, with
            // `label-1-3`, `baseline`, `series-label-0` and `series-label-1`
            // stranded inside the chart's own box. The round before it left the
            // same four, and neither round could say why until the success path
            // learned to speak.
            //
            // Grouping nothing keeps the chart WHOLE. It is still deleted
            // correctly on the next update — the parts tag, not the group, is what
            // carries a chart's membership. Ugly beats silently destructible: the
            // same reasoning this file already applies to a chart that loses its
            // group but keeps its config, and the same conclusion
            // `chooseGroupMembers` reaches when one member cannot be named at all.
            //
            // "IT IS STILL TAGGED, STILL RE-EDITABLE" USED TO BE THE NEXT WORDS,
            // and on this host they are false. Measured 2026-08-15 over the whole
            // round archive, joining each chart's grouping outcome to its tag:
            //
            //     grouped      64 chart(s),  1 lost the tag =  2%
            //     NOT grouped  62 chart(s), 41 lost the tag = 66%
            //
            // A group is a handle made in the grouping batch and never resolved,
            // so the config lands on it. Without one the tag falls back to a
            // `created` handle and this host refuses it two times in three. So
            // declining to group is not the free, conservative choice it reads as
            // here — it is the choice that costs the chart its config most of the
            // time.
            //
            // The branch is LEFT AS IT IS for now, deliberately: the alternative
            // (group the subset) strands the unmatched shapes, which is the
            // destructible failure this rule was written to prevent, and choosing
            // between two harms is not a call to make from a trace. What the
            // measurement changes is where to look — chart 4 of 8 matches 20 of
            // its 24 shapes in EVERY round, which is a re-read problem with a
            // fixed shape, not a coin toss. Fix the re-read and neither harm has
            // to be chosen. See `docs/BACKLOG.md`.
            //
            // Deliberately NOT falling through to the positional rule below. That
            // branch is safe only when NOTHING matched by id, because a slide
            // holding the user's own shapes can satisfy `items.length >=
            // created.length` while the chart itself read short — and "the last
            // N" would then reach past the chart into the user's content and
            // group it in, to be deleted with the chart on the next update.
            hostFriction.shortReReads += 1;
            trace("group", "the re-read matched only some of the chart's shapes", {
              index: i,
              drew: it.created.length,
              matched: matched.length,
              // See the empty-re-read line above: the sync count is what turns a
              // decay curve into a measurement.
              contextSyncs: syncsOf(context),
              // Survived a settle delay, so this is not the host still catching up.
              afterRetry: true,
            });
          } else if (items.length >= it.created.length) {
            // No ids to match on at all — a host that would not read them back.
            // The positional rule is still right for a slide this run added
            // blank, which is every slide the demo path draws on.
            freshMembers.set(i, items.slice(items.length - it.created.length));
          }
        });
        // WHAT THE PAUSE BOUGHT, counted in production. Everything that came
        // into this attempt and is not going round again resolved on it — so on
        // a retry attempt that is exactly the set of charts a settled re-read
        // repaired. It is the only number that can say whether the tracker's
        // settling-slide theory holds on the real host; the fake can only say
        // that the code does what it was written to do.
        if (attempt > 0) hostFriction.reReadsRepaired += pending.length - retry.length;
        pending = retry;
      }
    } catch (err) {
      // Say so. This catch was silent, and silence here is the mistake this
      // project has now paid for twice: `refreshed=0` in a run log means
      // "freshMembers was never set", and a refresh that THREW and a refresh
      // that matched nothing produce exactly the same zero. They want different
      // fixes — one is a host refusing a read, the other is a match rule that
      // does not fit the slide — and no deck can tell them apart. Same lesson as
      // the settle's "never asked" versus "asked and failed".
      trace("group", "the re-read before grouping faulted", {
        charts: refresher.length,
        error: errorText(err),
      });
    }
  }
  // A refreshed single shape is the tag target from here — the whole point of
  // refreshing it. Grouped items get their target replaced by the group below.
  for (const [i, fresh] of freshMembers)
    if (fresh.length === 1) {
      tagTargets[i] = fresh[0];
      targetFrom[i] = "refreshed";
    }
  if (groupable.length && supports("1.8")) {
    const took: { index: number; drew: number; took: number; by: string }[] = [];
    try {
      for (const { it, i } of groupable) {
        // Members RE-RESOLVED in this batch, off a slide handle taken in this
        // batch — never the proxies the re-read handed back a sync ago.
        //
        // A shape proxy carries its parent's object path. `freshMembers` came
        // out of `collections[k].items`, whose parent is the slide handle the
        // re-read sync resolved — so by now Office.js has rewritten that path to
        // `slides.getItem(id)`, and every member's path with it. A real host
        // named this exactly, listing the statements it refused:
        //
        //   var itemOrNullObject = slides.getItemOrNullObject(...);   ← fresh
        //   var slide = slides.getItem(...);                          ← the re-read's, rewritten
        //   var shapes1 = slide.shapes;
        //   var shape = shapes1.getItem(...);                         ← refused, 5010
        //
        // Both handles are in the same batch, and it is the rewritten one the
        // members hang off that fails. The members were never too old; their
        // PARENT was. Re-resolving by id costs nothing — the ids came from a
        // re-read one sync ago, so the shapes are there — and it is the same
        // access `settleAndTagChart` uses on the path that carries real charts.
        const refreshed = freshMembers.get(i);
        const choice = chooseGroupMembers({
          refreshedIds: refreshed?.map((s) => loadedValue(() => s.id)),
          askedForRefresh: !!it.refreshShapes,
        });
        // Nothing safe to group with. Leave the chart loose and let the tag
        // path have its turn — an ungrouped chart that carries its config is
        // re-editable; a grouped one that lost its config is not, and grouping
        // through a refused handle costs BOTH: `addGroup` throws, and the throw
        // takes the batch's tagging with it.
        if (choice.use === "none") {
          trace("group", "not grouping: no member handle this host will accept", {
            index: i,
            refreshed: refreshed?.length ?? 0,
            // WHICH SLIDE, because a lead is sitting in the archive that cannot
            // be tested without it. Rounds 043-045 each lost a chart's config,
            // and the one line that names a slide named `257#0` every time — an
            // id whose second half is `0`, which is not the shape this host
            // gives a slide it has finished adding. Two such ids appear among
            // every round's added slides.
            //
            // So the question is whether the charts that lose their tag are the
            // ones sitting on those slides, and today it can only be inferred
            // from a single settle-pass line. Putting the slide on the three
            // lines that decide a chart's fate — no group, tag refused, tag
            // never queued — makes it a join instead of a guess, and costs a
            // string already in hand.
            slide: slideKeyFor(it.opts, it.getSlide),
          });
          continue;
        }
        const members =
          choice.use === "ids" ? choice.ids.map((id) => it.getSlide().shapes.getItemOrNullObject(id)) : it.created;
        const group = (
          it.getSlide().shapes as unknown as { addGroup(items: PowerPoint.Shape[]): PowerPoint.Shape }
        ).addGroup(members);
        group.name = "PowerChart";
        // Accessible alt text on the group, queued in this same grouping sync so
        // a screen reader announces the chart — the description the engine built.
        applyAltText(group, it.opts);
        tagTargets[i] = group;
        targetFrom[i] = "group";
        grouped.add(i);
        took.push({ index: i, drew: it.created.length, took: members.length, by: choice.use });
      }
      await boundedSync(context, "grouping the chart's shapes");
      // AFTER the sync, so the name is an outcome the line knows — the mistake
      // `batch committed` made, and the reason there is no commit line here
      // either. Written unconditionally whenever anything grouped.
      //
      // This is the fifth failure-only field this repo has built. Grouping spoke
      // only when it refused, so the 2026-08-11 round left a slide carrying one
      // `PowerChart` group PLUS four loose shapes — `label-1-3`, `baseline`,
      // `series-label-0`, `series-label-1`, all inside the chart's own box, all
      // with lower ids than the group — and the log could not say whether the
      // group took a subset or four shapes from an earlier draw survived it.
      // It said so on its very next outing (`partial=1 left=0:4`), which is
      // what got the partial match thrown away rather than kept.
      //
      // `partial` therefore reports an INVARIANT now, not an outcome: the
      // re-read hands over whole matches only, so it should read 0 forever.
      // Kept because a zero that is checked is worth more than a field that was
      // removed once the bug it found was fixed — if a future change puts a
      // short match back into `freshMembers`, this is the line that says so,
      // and the alternative is finding out from someone's deck again.
      if (took.length) {
        const partial = took.filter((t) => t.took < t.drew);
        trace("group", "grouped the chart's shapes", {
          charts: took.length,
          partial: partial.length,
          ...(partial.length ? { left: partial.map((t) => `${t.index}:${t.drew - t.took}`).join(",") } : {}),
          by: [...new Set(took.map((t) => t.by))].join(","),
        });
      }
    } catch {
      /* grouping failed — shapes stay ungrouped, charts are already on the slide */
      // RE-RESOLVED by id off a slide handle taken now, which is the same trick
      // the grouping loop above uses and the one this recovery forgot.
      //
      // It used to hand back `freshMembers[0]`, and that is the proxy whose
      // PARENT was rewritten to `slides.getItem(id)` — the exact handle whose
      // refusal just failed the grouping. So the recovery reached for the one
      // thing guaranteed to fail again, and a real host said so twice in the
      // same breath: `InvalidParam passed to GetItem(id)` on the group, then
      // `Cannot read properties of undefined (reading 'add')` on the tag.
      // Five charts in one run lost their config that way.
      //
      // Only the id crosses; the handle never does. Where the id cannot be read
      // there is nothing better to offer than the proxy itself, and the settle
      // pass gets the last word either way.
      for (const { i } of groupable) {
        const fresh = freshMembers.get(i)?.[0];
        const id = fresh ? loadedValue(() => fresh.id) : undefined;
        const byId = typeof id === "string" && id;
        tagTargets[i] = byId
          ? (items[i].getSlide().shapes.getItemOrNullObject(id) as PowerPoint.Shape)
          : (fresh ?? items[i].created[0]);
        targetFrom[i] = byId ? "by-id" : fresh ? "refreshed" : "created";
      }
      grouped.clear();
    }
  }
  const partsJson = await ungroupedFallback(context, items, tagTargets, grouped, freshMembers);
  // Tags are PowerPointApi 1.3+; keep the chart re-editable where supported.
  const taggable = items
    .map((it, i) => ({ it, i, target: tagTargets[i] }))
    .filter((t) => t.it.opts.tagData && t.target);
  // Which items' config tag COMMITTED. Not "was queued": the queue is where
  // this silently failed before, and the caller reports this number.
  const taggedOk = new Set<number>();
  if (taggable.length && supports("1.3")) {
    /** Charts whose tag writes were successfully QUEUED — see the loop below. */
    const queued: typeof taggable = [];
    try {
      for (const t of taggable) {
        const { it, i, target } = t;
        // Per chart, because one bad target used to cost every chart after it.
        //
        // A real host answered `target.tags` as UNDEFINED here — "Cannot read
        // properties of undefined (reading 'add')", four times in one run,
        // each time on a chart whose grouping had just been refused with
        // InvalidParam 5010. That throw is SYNCHRONOUS, so it escaped the loop
        // and took the whole batch's tagging with it: the charts after it lost
        // their config without ever being attempted. `updateChartsInSlides`
        // was made per-chart resilient for exactly this shape of failure; this
        // loop never was.
        try {
          target!.tags.add(CHART_TAG, it.opts.tagData!);
          // The rest of an ungrouped chart travels with the tagged shape, so an
          // in-place update can delete all of it — see CHART_PARTS_TAG.
          if (partsJson[i]) target!.tags.add(CHART_PARTS_TAG, partsJson[i]!);
          // What this chart was drawn from, so its next update can tell a scene
          // it can diff from one it only thinks it can. See CHART_SCENE_TAG.
          if (it.opts.sceneTag) target!.tags.add(CHART_SCENE_TAG, it.opts.sceneTag);
          target!.load("id,left,top");
          queued.push(t);
        } catch (err) {
          trace("group", "a chart's tag could not even be queued", {
            index: i,
            from: targetFrom[i],
            // See the `not grouping` line: the slide is what turns a lead about
            // `NNN#0` ids into a join.
            slide: slideKeyFor(it.opts, it.getSlide),
            error: errorText(err),
          });
        }
      }
      if (!queued.length) throw new Error("no chart's tag could be queued");
      await boundedSync(context, "writing the chart's config tag");
      // The config tag is on the slide from here. The origin tag below is a
      // separate sync and a separate risk — losing it costs drag tracking,
      // not re-editability — so `tagged` is decided at THIS line, not after.
      for (const { i } of queued) taggedOk.add(i);

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
      //
      // ITS OWN try, because losing it is not what the catch below describes.
      // The config tag committed one sync ago and `taggedOk` already says so:
      // these charts ARE re-editable, and reporting "not re-editable until
      // repaired" over a failed origin write is a false alarm that sends a
      // reader looking for lost configs there are none of.
      //
      // The two writes are separate syncs, so either can fail without the other
      // — and this is the pair that fails asymmetrically, because the config tag
      // commits FIRST. A host that starts refusing mid-pass therefore produces
      // config-lands-origin-fails as its good case, and the old shared catch
      // reported it as "charts are not re-editable until repaired". That is the
      // lie worth spending a try/catch to avoid: it is false, and it is the
      // outcome a degrading host reaches most often.
      //
      // Arrived alongside the `tagAnchorIndex` experiment, which has since been
      // reverted for want of any measured effect. This did not go with it: it
      // describes the order of two writes, which the anchor never touched.
      try {
        for (const { it, target } of queued) {
          target!.tags.add(
            CHART_ORIGIN_TAG,
            JSON.stringify([it.opts.left ?? 60, it.opts.top ?? 90, target!.left, target!.top]),
          );
        }
        await boundedSync(context, "writing the chart's origin tag");
      } catch (err) {
        // Named for what it costs: an update re-renders at the recorded origin
        // and cannot shift it by how far the user has dragged the chart since.
        // The chart is re-editable; it may just walk if it has been moved.
        trace("group", "origin tag lost — the chart is re-editable but an update may not follow a drag", {
          charts: queued.length,
          from: countBy(queued.map((t) => targetFrom[t.i])),
          error: errorText(err),
        });
      }
    } catch (err) {
      // The charts are on the slide but carry no config, so they are not
      // re-editable — and this used to be entirely silent, which is why a real
      // run shipped 19 untagged charts with nothing in the log to say when or
      // why. No retry here on purpose: a retry against a host that just
      // dropped a sync is how this project got its duplicate slides. The
      // settled repair pass re-reads the deck and plans a `retag`, which is
      // the same job done with evidence.
      trace("group", "tagging failed — charts are not re-editable until repaired", {
        charts: queued.length || taggable.length,
        // See `targetFrom`. Without this the trace says a write failed and not
        // which of four handles it went through, and the two resolved ones are
        // the whole question.
        from: countBy((queued.length ? queued : taggable).map((t) => targetFrom[t.i])),
        // The slides this batch was writing to, deduplicated — the batch covers
        // several charts, so one id would be a guess about which. Same purpose
        // as on the `not grouping` line above: it turns the `NNN#0` lead into a
        // join rather than an inference off a single settle-pass line.
        slides: [
          ...new Set((queued.length ? queued : taggable).map((t) => slideKeyFor(t.it.opts, t.it.getSlide))),
        ].join(","),
        error: errorText(err),
      });
    }
  }
  // Whatever the tagging sync did not resolve, ask for once more, on its own.
  //
  // Only the taggable items got `load("id,left,top")` above, and only if that
  // sync landed. Everything else — an untagged chart, a host below 1.3, and
  // every chart in a run whose tagging threw — reaches here holding a proxy
  // that has never been loaded. That is not a reason to lose the target: the
  // shapes are on the slide and the caller wants to know where. One extra
  // round-trip, paid only when something is actually missing, and best-effort
  // like everything else after the shapes commit.
  const unresolved = items.map((_, i) => i).filter((i) => tagTargets[i] && !targetRef(tagTargets[i]));
  if (unresolved.length) {
    try {
      for (const i of unresolved) tagTargets[i]!.load("id,left,top");
      await boundedSync(context, "reading back where the charts landed");
    } catch {
      /* the host would not say where they are — the caller keeps its old target */
    }
  }
  return items.map((_, i) => ({
    // Resolved to plain values HERE, inside the context that can still read
    // them. A caller handed the proxy would read it after this run is gone.
    target: targetRef(tagTargets[i]),
    partIds: partsJson[i] ? (JSON.parse(partsJson[i]!) as string[]) : undefined,
    grouped: grouped.has(i),
    tagged: taggedOk.has(i),
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
  freshMembers?: Map<number, PowerPoint.Shape[]>,
): Promise<(string | undefined)[]> {
  const partsJson: (string | undefined)[] = items.map(() => undefined);
  const loose = (i: number) => !grouped.has(i) && tagTargets[i];
  // Only a re-editable chart (one that gets a config tag, so a host with tags)
  // needs its parts written down; the update path is what reads them.
  const hasTags = supports("1.3");
  /**
   * The re-read's members where there are any, the drawing proxies otherwise.
   *
   * `it.created` are the proxies `addGeometricShape` handed back, and on the web
   * Office.js has by now rewritten their object paths to `shapes.getItem(id)` —
   * which is the call this host refuses. That is not a theory: `reading back an
   * ungrouped chart's shape ids` failed three times in the 2026-08-07 run with
   * `InvalidParam passed to GetItem(id)` at `errorLocation:
   * ShapeCollection.getItem`, and each failure cost that chart its
   * CHART_PARTS_TAG.
   *
   * Which matters more here than the name "fallback" suggests. PowerPoint on the
   * web ungroups every chart it cannot group, so on that host this path carries
   * EVERY chart — and a chart with no parts list is one whose in-place update
   * deletes its anchor and redraws all 24 shapes, leaving the other 23 behind.
   * The chart grows by a whole chart on every edit.
   *
   * `freshMembers` came out of a collection read, in `created` order by
   * construction (see the match above), so index 0 is still the anchor and
   * "everything after it is a part" still holds. A member of an `items` read is
   * the proxy pattern this host does honour — the same one `settleByCollection-
   * Read` relies on.
   */
  const parts = (it: Grouping, i: number): PowerPoint.Shape[] => (freshMembers?.get(i) ?? it.created).slice(1);
  const siblings = items.map((it, i) => (hasTags && loose(i) && it.opts.tagData ? parts(it, i) : []));
  const alt = items.map((it, i) => ({ it, i })).filter(({ it, i }) => loose(i) && wantsAltText(it.opts));
  if (!alt.length && !siblings.some((s) => s.length)) return partsJson;
  try {
    // Queueing counts as best-effort too. `s.load("id")` and `applyAltText` are
    // ordinary property access on host proxies, and on the web those can throw
    // synchronously — from out here that rejected the whole PowerPoint.run and
    // failed an update whose shapes were already committed. The catch was
    // always meant to cover this: the comment below is the contract, and it is
    // just as true of a queue that faulted as of a sync that did.
    for (const s of siblings.flat()) s.load("id");
    for (const { it, i } of alt) applyAltText(tagTargets[i]!, it.opts);
    // Labelled and bounded, like every other sync in this phase. This one was
    // neither — a bare `context.sync()` with no `step`, so an error escaping it
    // carried no `at=`, and no deadline, so a host that went quiet here hung the
    // whole insert with nothing on screen. It runs on EVERY web-host insert:
    // grouping is refused there, which is what puts every chart down this path.
    await boundedSync(context, "reading back an ungrouped chart's shape ids");
  } catch {
    /* no alt text or id read-back here — the chart is on the slide regardless */
    return partsJson;
  }
  siblings.forEach((shapes, i) => {
    // Read past a sibling the host did not answer for, rather than out of the
    // whole insert. This runs AFTER the sync above and outside its catch, so a
    // raw read here threw straight through `groupAndTagAll` into a caller whose
    // chart was already drawn — and the parts list it was building is the one
    // thing that makes an ungrouped chart deletable as a unit.
    const ids = shapes
      .map((s) => loadedValue(() => s.id))
      .filter((id): id is string => typeof id === "string" && id.length > 0);
    if (ids.length) partsJson[i] = JSON.stringify(ids);
  });
  return partsJson;
}

function getTargetSlide(context: PowerPoint.RequestContext, slideId?: string): PowerPoint.Slide {
  if (slideId) {
    // getItem, not getItemOrNullObject: this returns synchronously, before any
    // sync could tell us whether the id resolved, so there is no null to check
    // — a dead id throws here and the selected slide is the honest fallback.
    try {
      return context.presentation.slides.getItem(slideId);
    } catch {
      /* the slide is gone — fall through to whatever the user is looking at */
    }
  }
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
  // `dashKind` answers `none` for an array that specifies no dash, which is what
  // the preview has always drawn for one. Guarding on `s.dash` being TRUTHY put
  // a dotted line in the deck where the preview drew it solid — `[]` is truthy.
  const dashStyle = dashKind(s.dash);
  const setDash = (shape: PowerPoint.Shape) => {
    if (dashStyle === "none") return;
    try {
      // Map to the nearest native line style: a dotted array (e.g. waterfall
      // carry connectors) stays dotted instead of flattening to a generic dash.
      shape.lineFormat.dashStyle =
        dashStyle === "dot" ? PowerPoint.ShapeLineDashStyle.roundDot : PowerPoint.ShapeLineDashStyle.dash;
    } catch {
      /* dash style unsupported on this host */
    }
  };

  // A rotated rectangle cannot carry a dash style, so a dashed line has to be a
  // real line shape. Same question as `setDash` asks, for the same reason: `[]`
  // is truthy and was forcing this branch for a line that is not dashed.
  if (w < 0.5 || h < 0.5 || dashStyle !== "none") {
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
 * dropped. Colour NAMES the CSS table knows go through verbatim, since Office
 * knows the same ones; everything else — including a bare word that is not a
 * colour at all — normalises to `#RRGGBB`.
 */
/**
 * A point size PowerPoint will accept, enforced where it is WRITTEN.
 *
 * The third sink again, and the same argument the skill's `fontPt` already
 * makes about its own bytes: `font.size` was assigned straight from the scene
 * node here, unchecked, where the pptx sink clamps to OOXML's `ST_TextFontSize`
 * (1pt to 4000pt) and the SVG sink at least substitutes for a non-finite value.
 *
 * It is the worst of the three places to be missing it. A rejected property
 * THROWS, and a throw inside a draw batch takes the batch's other work with it
 * — on this host that includes the config tag, so the user gets a chart that is
 * not re-editable, or no chart at all. A clamped label is merely small.
 *
 * The engine clamps `style.fontSize` on the way in and its layouts now floor
 * every fitted label, so nothing should arrive out of range. That is exactly
 * what was true of the other two sinks when each of them was found to need one.
 */
const officeFontPt = (v: number): number => (Number.isFinite(v) ? Math.min(4000, Math.max(1, v)) : 12);

const officeHex = (color: string): string => {
  // The THIRD colour sink, and the one that runs in a real PowerPoint. The
  // other two — `src/core/color.ts` and the skill's `pptx-paint.mjs` — each had
  // the same hole independently: a palette of NUMBERS threw
  // `color.trim is not a function`. Here it would take down a live insert, on
  // the path a user is actually standing on, for a config that came out of the
  // JSON box or a shape tag written in another deck.
  //
  // `toHex6` handles anything unrecognised already, so a non-string just needs
  // to reach it rather than die on the way.
  const raw = typeof color === "string" ? color.trim() : "";
  // `transparent` is a NAMED colour by the test below and Office.js is not one
  // of the things that knows it: `setSolidColor("transparent")` is rejected, and
  // the shape is left with whatever fill it had. It carries alpha 0 through
  // `alphaOf`, so the concrete value here is never seen — it only has to be
  // something the host accepts.
  if (/^transparent$/i.test(raw)) return "#ffffff";
  // A word Office actually KNOWS, not merely a word.
  //
  // The test used to be `/^[a-zA-Z]+$/` — any run of letters at all — on the
  // reasoning that Office.js knows the CSS names and `toHex6` would flatten
  // them to grey. Half of that is no longer true: this sink's `toRgb` carries
  // the CSS table now, so a known name survives `toHex6` intact and only an
  // UNKNOWN word becomes grey.
  //
  // The other half was never true. `style.palette: ["banana"]` — or a series
  // colour of `constructor`, which the pane's own template store made reachable
  // — is a run of letters and is not a colour, and `setSolidColor` on a name
  // the host does not know is rejected. That rejection lands inside a draw
  // batch, so one bad word did not degrade one shape: it took the whole batch,
  // and with it the chart.
  return isNamedColor(raw) ? raw : toHex6(raw);
};

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

/**
 * Redraw a chart by writing only what changed, or say no.
 *
 * The add-in's normal update deletes every shape and adds every shape back. On
 * PowerPoint for the web that costs ~50 seconds for a 24-shape chart, and the
 * measured diffs say almost none of it is needed: a retitle changes one node of
 * twenty-four, a single edited data point changes two.
 *
 * Returns true only when the chart is now correct on the slide. Every other
 * answer is false, and false means the caller deletes and redraws exactly as it
 * always has — so this can only ever make an update faster, never wrong. It
 * refuses when:
 *
 * - the chart carries no stored config, so there is no old scene to diff;
 * - the stored config does not parse, or renders to a different node count than
 *   the shapes the chart actually has;
 * - the scene fingerprint is missing or does not match what the stored config
 *   renders to NOW — the engine has changed since this chart was drawn, and the
 *   scene being diffed is not the one on the slide;
 * - `planSceneUpdate` refuses, or the change is too much of the chart to be
 *   worth it;
 * - any shape it would have to write to is one the host will not confirm.
 *
 * The tag writes ride in the same sync as the property writes, deliberately. A
 * chart whose picture is new and whose config is old is the worst outcome
 * available here — re-opening it would silently revert the edit — so if the
 * host refuses the batch, nothing has been deleted, this returns false, and the
 * redraw does the whole job.
 */
async function tryInPlaceUpdate(
  context: PowerPoint.RequestContext,
  entry: {
    it: { scene: Scene; target: EditTarget; opts?: InsertOptions };
    old: PowerPoint.Shape;
    parts: PowerPoint.Shape[];
  },
  opts: InsertOptions,
  tags: { config?: string; scene?: string },
  step: <T>(what: string, run: () => Promise<T>) => Promise<T>,
): Promise<boolean> {
  const { it, old, parts } = entry;
  // Say WHY, every time it declines.
  //
  // The first real round on a build carrying this path produced not one line —
  // no success, no refusal — which is indistinguishable from the code not being
  // there at all, and left the reason to be reasoned out from grouping traces
  // and a deck inventory. That is the sixth failure-only field this repo has
  // built, and the first one written by the same session that wrote the rule
  // down. The reasons are not interchangeable: "this chart is grouped, so there
  // is no node-to-shape mapping" is a fact about the design, "the host would not
  // read the shape ids back" is a fact about the host, and "the fingerprint did
  // not match" means an engine change is being handled correctly. Three
  // different next steps, one silent `return false`.
  const no = (why: string, extra?: Record<string, unknown>): false => {
    trace("draw", "not updating in place — redrawing instead", { why, ...extra });
    return false;
  };
  if (!it.opts?.tagData) return no("this update carries no config to write");
  // A picture is not in the scene, so the scene cannot decide this update.
  //
  // `render: "image"` on the config does NOT produce a picture — the renderer
  // takes that path only when handed `pictureBase64` — so collapsing a chart to
  // a picture builds the SAME scene it already has. The differ compared 24
  // nodes to 24 identical nodes, answered "nothing changed", wrote nothing, and
  // reported success. Round `89675b6` is the case: `updated only the shapes
  // that changed {changed: 0, of: 24}` on the collapse, and the slide still
  // held its 24 native shapes afterwards.
  //
  // Costly beyond the self-test, because the auto-picture fallback is what the
  // add-in reaches for when this host has ALREADY failed to draw shapes — so
  // the one path that exists to rescue a struggling host was the one being
  // skipped, silently, with a success reported to the user.
  //
  // Refused rather than taught to handle it: the fast path writes a closed set
  // of `rect` and `text` properties, and a picture fill is neither.
  if (it.opts.pictureBase64) return no("this update draws a picture, which is not in the scene the differ compares");
  if (!tags.config) return no("the chart carries no stored config to diff against");
  // Grouped charts land here: their shapes are inside the group and the parts
  // tag does not list them, so there is nothing to write to. Named separately
  // from a missing fingerprint because it is permanent for that chart, not a
  // one-off.
  if (!parts.length) return no("the chart has no parts list, so its nodes cannot be mapped to shapes");
  if (!tags.scene) return no("the chart carries no scene fingerprint — it was drawn by an older build");
  let prev: Scene;
  try {
    prev = buildChart(JSON.parse(tags.config) as ChartConfig);
  } catch {
    return no("the stored config will not parse");
  }
  // The shapes have to line up with the nodes ONE FOR ONE, because that is how
  // they are found: the anchor is node 0 and the parts tag lists the rest in
  // drawing order. A chart whose parts tag is short — one the host would not
  // read back when it was drawn — has no usable mapping at all.
  if (prev.nodes.length !== parts.length + 1)
    return no("the parts list does not match the scene one for one", {
      parts: parts.length + 1,
      nodes: prev.nodes.length,
    });
  if (sceneFingerprint(prev) !== tags.scene)
    return no("the stored config no longer renders to the scene that was drawn");
  const plan = planSceneUpdate(prev, it.scene);
  if (!plan) return no("the two scenes are not node-compatible");
  if (!worthUpdating(plan, it.scene.nodes.length))
    return no("too much of the chart changed to be worth writing shape by shape", {
      changed: plan.changed.length,
      of: it.scene.nodes.length,
    });
  const shapes = [old, ...parts];
  // Only shapes the host CONFIRMED. `isLive` is the same test the delete path
  // uses before touching anything, and for the same reason: writing to a shape
  // the host would not answer for is how an update comes to edit something that
  // is not ours.
  if (plan.changed.some((n) => !isLive(shapes[n])))
    return no("the host would not confirm every shape that had to change", { changed: plan.changed.length });
  const { dx, dy } = frameOrigin(opts);
  try {
    for (const n of plan.changed) applyNodeInPlace(shapes[n], it.scene.nodes[n], dx, dy, opts);
    old.tags.add(CHART_TAG, it.opts.tagData);
    old.tags.add(CHART_SCENE_TAG, sceneFingerprint(it.scene));
    await step("updating the shapes a chart actually changed", () => context.sync());
  } catch (err) {
    trace("draw", "in-place update refused — redrawing instead", {
      changed: plan.changed.length,
      of: it.scene.nodes.length,
      error: errorText(err),
    });
    return false;
  }
  trace("draw", "updated only the shapes that changed", {
    changed: plan.changed.length,
    of: it.scene.nodes.length,
    saved: it.scene.nodes.length * 2 - plan.changed.length,
  });
  return true;
}

/**
 * Where a scene's origin lands on the slide.
 *
 * One reader, two callers: the draw places every node at `dx + n.x`, and the
 * in-place update has to write the identical number onto shapes that are
 * already there. If the two ever disagree, an edit silently teleports every
 * changed shape by the difference — a defect no test that only ever draws could
 * see, and the reason the pair of defaults lives here instead of being spelled
 * twice.
 */
function frameOrigin(opts: InsertOptions): { dx: number; dy: number } {
  return { dx: opts.left ?? 60, dy: opts.top ?? 90 };
}

/**
 * Write a changed node onto the shape that already draws it.
 *
 * The other half of `planSceneUpdate`: the planner decides which shapes a
 * change can be applied to, and this applies it. Only `rect` and `text` reach
 * here, because they are the only kinds whose whole drawn appearance is a
 * closed set of property writes — everything else has geometry baked at
 * creation with no freeform path to edit afterwards.
 *
 * MUST set every property `addNode` sets for the same kind, and that is not a
 * style note. A property the adder writes and this one does not is a chart that
 * looks right when first drawn and wrong after an edit — a bar that keeps its
 * old colour, a label that keeps its old font — with nothing in any log to say
 * so, because from the host's point of view the update succeeded. The two are
 * held in lockstep by a source scan in `test/office-render.test.ts`; if you add
 * a line to one of the adders, that test will name the one you forgot.
 *
 * Queues only. The caller owns the sync, so a whole chart's changed shapes go
 * out in one batch and a refusal takes the fallback rather than half a chart.
 */
function applyNodeInPlace(shape: PowerPoint.Shape, n: SceneNode, dx: number, dy: number, opts: InsertOptions): void {
  if (n.kind === "rect") {
    shape.left = dx + n.x;
    shape.top = dy + n.y;
    shape.width = Math.max(0.2, n.w);
    shape.height = Math.max(0.2, n.h);
    if (n.fill === "none") shape.fill.clear();
    else solidFill(shape.fill, n.fill);
    if (n.stroke && (n.strokeWidth ?? 0) > 0) {
      strokeColor(shape.lineFormat, n.stroke);
      shape.lineFormat.weight = n.strokeWidth ?? 1;
    } else {
      shape.lineFormat.visible = false;
    }
    if (n.name) shape.name = n.name;
    return;
  }
  if (n.kind !== "text") return;
  shape.left = dx + n.x;
  shape.top = dy + n.y;
  shape.width = Math.max(4, n.w);
  shape.height = Math.max(4, n.h);
  shape.fill.clear();
  shape.lineFormat.visible = false;
  const tf = shape.textFrame;
  // The string itself. `addTextBox` takes it as an argument, so this is the one
  // write with no counterpart in the adder's body — and it is the whole point
  // of the fast path: a retitle changes one node, and this is the line that
  // applies it.
  tf.textRange.text = n.text;
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
  font.size = officeFontPt(n.fontSize);
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
      // outline. symbolPreset names are GeometricShapeType keys.
      const geo = (PowerPoint.GeometricShapeType as unknown as Record<string, PowerPoint.GeometricShapeType>)[
        symbolPreset(n.shape)
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
  font.size = officeFontPt(n.fontSize);
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
