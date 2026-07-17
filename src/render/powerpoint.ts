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
import { polar, arrowheadBox, wedgeFanSteps, wedgeFanChord, SYMBOL_PRESET } from "../core/geometry";
import type { Scene, SceneNode, TextNode, WedgeNode } from "../core/scene";

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
}

/** Tag key under which the chart's serialized config is persisted. */
export const CHART_TAG = "POWERCHART_CONFIG";

const DEFAULT_FONT = "Segoe UI";

/** Where an existing PowerChart lives on the deck, for in-place update. */
export interface EditTarget {
  slideId: string;
  shapeId: string;
  left: number;
  top: number;
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

let lateSubscriber: ((msg: string) => void) | null = null;

/** Be told when a call we already gave up on finally settles. */
export function onLateSync(cb: (msg: string) => void): void {
  lateSubscriber = cb;
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
 */
function withTimeout<T>(p: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  let abandoned = false;
  const started = Date.now();
  const describe = (outcome: string) => {
    if (!abandoned) return;
    lastLateSync = `${what}: ${outcome} after ${Math.round((Date.now() - started) / 1000)}s`;
    lateSubscriber?.(lastLateSync);
  };
  p.then(
    () => describe("the host eventually SUCCEEDED"),
    (err: unknown) => describe(`the host eventually FAILED — ${errorText(err)}`),
  );
  return Promise.race([
    p.finally(() => clearTimeout(timer)),
    new Promise<never>((_, reject) => {
      timer = setTimeout(() => {
        abandoned = true;
        reject(new Error(`PowerPoint did not respond while ${what} (${ms / 1000}s)`));
      }, ms);
    }),
  ]);
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
const BATCH_TIMEOUT_MS = 45_000;

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
    await groupAndTagAll(context, [{ getSlide, created, opts }]);
    onPhase?.("done");
  });
}

/** Replace an existing PowerChart group with a re-rendered scene, in place. */
export async function updateChartInSlide(scene: Scene, target: EditTarget, opts: InsertOptions = {}): Promise<void> {
  await updateChartsInSlides([{ scene, target, opts }]);
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
): Promise<void> {
  if (!items.length) return;
  await PowerPoint.run(async (context) => {
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
    if (!live.length) return;
    const withOld = live.map(({ it, slide }) => ({ it, slide, old: slide.shapes.getItemOrNullObject(it.target.shapeId) }));
    await context.sync();

    // 2. Drop the old shapes — one sync for all of them.
    for (const { old } of withOld) if (!old.isNullObject) old.delete();
    await context.sync();

    // 3. Redraw each chart in batches. One of these charts is on the slide the
    //    user is looking at, and a live canvas will not take a whole chart in
    //    one sync — so the batching is not an optimisation here, it is the only
    //    way the shapes arrive at all. Per chart, because a chart's shapes must
    //    all reach the same slide.
    const rendered: Grouping[] = [];
    for (const { it, slide } of withOld) {
      const opts: InsertOptions = { ...it.opts, left: it.target.left, top: it.target.top };
      // An existing slide's proxy is stable across syncs — hold it. Only a
      // freshly-added slide needs a per-batch fresh proxy; see SlideThunk.
      const getSlide: SlideThunk = () => slide;
      rendered.push({ getSlide, created: await renderShapesChunked(context, getSlide, it.scene, opts), opts });
    }

    // 4-5. Group, then tag — one sync each, however many charts.
    await groupAndTagAll(context, rendered);
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

    const tags = shapes.items.map((s) => {
      const tag = s.tags.getItemOrNullObject(CHART_TAG);
      tag.load("value");
      return tag;
    });
    await context.sync();

    for (let i = 0; i < shapes.items.length; i++) {
      if (!tags[i].isNullObject && tags[i].value) {
        const s = shapes.items[i];
        return {
          configJson: tags[i].value,
          target: { slideId: slide.id, shapeId: s.id, left: s.left, top: s.top },
        };
      }
    }
    return null;
  });
}

/**
 * Bounds of the currently selected shape when it is NOT a PowerChart —
 * used to insert a new chart into a selected placeholder/frame.
 */
export async function getSelectionBounds(): Promise<{ left: number; top: number; width: number; height: number } | null> {
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
    const tags = shapes.items.map((s) => {
      const tag = s.tags.getItemOrNullObject(CHART_TAG);
      tag.load("value");
      return tag;
    });
    await context.sync();
    return shapes.items
      .map((s, i) => ({ s, tag: tags[i] }))
      .filter(({ tag }) => !tag.isNullObject && tag.value)
      .map(({ s, tag }) => ({
        configJson: tag.value,
        target: { slideId: slide.id, shapeId: s.id, left: s.left, top: s.top },
      }));
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

    const lookups: { slideId: string; shape: PowerPoint.Shape; tag: PowerPoint.Tag }[] = [];
    for (const slide of perSlide) {
      for (const shape of slide.shapes.items) {
        const tag = shape.tags.getItemOrNullObject(CHART_TAG);
        tag.load("value");
        lookups.push({ slideId: slide.id, shape, tag });
      }
    }
    await context.sync();

    return lookups
      .filter((l) => !l.tag.isNullObject && l.tag.value)
      .map((l) => ({
        configJson: l.tag.value,
        target: { slideId: l.slideId, shapeId: l.shape.id, left: l.shape.left, top: l.shape.top },
      }));
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
 * Append `count` blank slides and return a fresh-proxy thunk for each.
 *
 * By index, off a `getCount()` taken in its OWN sync BEFORE the adds —
 * `slides.add()` always appends to the end, so the new slides are
 * `start .. start+count-1`. The adds then get their own commit sync, and the
 * positional thunks need nothing more.
 *
 * There is deliberately NO post-add count check. `getCount()` queued in the SAME
 * sync as the adds returns the PRE-add total on PowerPoint web — the adds are
 * queued but not yet reflected in the count — so a delta assertion there throws
 * "added 0 of N" while the slides are in fact appearing. `start` is read cleanly
 * before any add is queued, which is the only count reading this host reports
 * reliably.
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
  return Array.from({ length: count }, (_, i) => () => context.presentation.slides.getItemAt(start + i));
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
      await renderShapesChunked(context, slideThunks[i], scenes[i], { left: 0, top: 0, group: false, tagData: undefined });
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
 * How many NATIVE shapes a scene becomes on the host — the number the budget
 * actually cares about, which is NOT the node count. A wedge fans out into
 * `wedgeFanSteps` shapes (+2 stroke edges); a polygon draws one line per edge.
 * So a 10-node pie is ~50 shapes and a 10-node violin ~250 — counting nodes waved
 * both straight past the budget and the host choked, which is exactly what the
 * self-check caught. Everything else is one shape.
 */
export function estimateOfficeShapes(scene: Scene): number {
  let total = 0;
  for (const n of scene.nodes) {
    if (n.kind === "wedge") {
      const span = n.endAngle - n.startAngle;
      total += wedgeFanSteps(n.r, span).steps + (n.stroke && span < 359.9 ? 2 : 0);
    } else if (n.kind === "polygon") {
      total += n.points.length; // one line per edge, closed
    } else {
      total += 1;
    }
  }
  return total;
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
async function stampSlide(context: PowerPoint.RequestContext, getSlide: SlideThunk, title: string, detail: string): Promise<void> {
  const box = (getSlide().shapes as unknown as {
    addTextBox(text: string, box: { left: number; top: number; width: number; height: number }): PowerPoint.Shape;
  }).addTextBox(`${title} — ${detail}`, { left: 24, top: 12, width: 912, height: 46 });
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
}

/** A demo-deck insert's self-verification report. */
export interface DemoReport {
  results: DemoResult[];
  /**
   * How much the deck ACTUALLY grew, read back from the host. It SHOULD equal
   * `results.length` (one slide per item, even skipped/failed ones). A shortfall
   * means the host silently lost slides mid-run — an otherwise-invisible
   * corruption that a regression run must surface, not hide.
   */
  slidesAdded: number;
}

/** The current slide count, read in its own settled sync (reliable on web). */
async function slideCount(): Promise<number> {
  return PowerPoint.run(async (context) => {
    const c = context.presentation.slides.getCount();
    await context.sync();
    return c.value;
  });
}

export async function insertDemoDeck(
  items: { scene: Scene; tagData?: string }[],
  onProgress?: (done: number, total: number) => void,
): Promise<DemoReport> {
  const results: DemoResult[] = [];
  let lastError: unknown;
  // The blank-layout id is a plain string valid across contexts, so it is looked
  // up once (on the first slide's context) and reused for the rest.
  let layoutId: string | undefined;
  let layoutResolved = false;
  // Bracket the run with a settled slide count, so a regression run can prove the
  // deck grew by exactly one slide per item — the lost-slide check.
  const before = await slideCount();
  for (let i = 0; i < items.length; i++) {
    const shapeCount = estimateOfficeShapes(items[i].scene);
    const tooDense = shapeCount > DEMO_SHAPE_BUDGET;
    let created = 0;
    let status: DemoResult["status"] = "rendered";
    try {
      await PowerPoint.run(async (context) => {
        if (!layoutResolved) {
          layoutId = await blankLayoutId(context);
          layoutResolved = true;
        }
        const [getSlide] = await addSlides(context, 1, layoutId);
        if (tooDense) {
          // Do NOT attempt it: this many shapes will not land on web, and trying
          // burns the timeout and pushes the host toward a crash. Leave a stamped
          // placeholder so the slide count and order still line up.
          await stampSlide(context, getSlide, "NOT COMPLETE", `Too dense for this host — ${shapeCount} shapes, not rendered`);
          status = "skipped";
          return;
        }
        const opts: InsertOptions = { left: 60, top: 90, group: true, tagData: items[i].tagData };
        const drawn = await renderShapesChunked(context, getSlide, items[i].scene, opts);
        created = drawn.length;
        // Shapes are committed by now, so grouping cannot roll them back.
        await groupAndTagAll(context, [{ getSlide, created: drawn, opts }]);
      });
    } catch (err) {
      // One chart the host would not draw does not sink the rest of the deck.
      lastError = err;
      status = "failed";
      // Mark the half-rendered slide so a partial chart is not mistaken for a
      // real one. A fresh context, because the failed render poisoned its own.
      await stampLastSlide("NOT COMPLETE", "PowerPoint stopped responding while drawing this chart").catch(() => {});
    }
    results.push({ created, status });
    onProgress?.(i + 1, items.length);
  }
  const after = await slideCount();
  // A whole deck lost to HOST errors (not just skipped-as-dense) is a real
  // failure — surface it so the pane says "Failed", not "inserted 0 of N". If
  // everything was merely skipped, there is no error to throw.
  if (items.length && results.every((r) => r.status !== "rendered") && lastError !== undefined) {
    throw lastError instanceof Error ? lastError : new Error(String(lastError));
  }
  return { results, slidesAdded: after - before };
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
 * Render a scene onto a slide, committing in small batches.
 *
 * This forfeits all-or-nothing, which is a real loss: a failure now strands a
 * partial chart instead of leaving the slide clean. It buys the only thing that
 * matters more — the chart arriving at all — and it is why `created` is
 * returned even on failure, so a caller can clean up what landed.
 *
 * The batches also make progress REAL: shapes committed over shapes total is a
 * fact, not the estimate a single opaque sync would force us to invent.
 */
async function renderShapesChunked(
  context: PowerPoint.RequestContext,
  getSlide: SlideThunk,
  scene: Scene,
  opts: InsertOptions,
  onBatch?: (sending: number, total: number) => void,
): Promise<PowerPoint.Shape[]> {
  const left = opts.left ?? 60;
  const top = opts.top ?? 90;
  const created: PowerPoint.Shape[] = [];
  const total = scene.nodes.length;
  for (let i = 0; i < total; i += SHAPES_PER_SYNC) {
    // Fresh slide proxy per batch: a proxy held across the previous sync may have
    // been rewritten to an unusable getItem(id) — see SlideThunk.
    const shapes = getSlide().shapes;
    for (const n of scene.nodes.slice(i, i + SHAPES_PER_SYNC)) {
      created.push(...addNode(shapes, n, left, top, opts));
    }
    const upTo = Math.min(i + SHAPES_PER_SYNC, total);
    // Reported BEFORE the sync, and deliberately: the sync is where a bad host
    // stops answering, so this is the number that has to be on screen WHILE we
    // wait. Reporting after would leave the pane naming the previous phase and
    // blaming the wrong one for the stall.
    onBatch?.(upTo, total);
    // Budget per BATCH, not per chart: a stalled host must still be caught, but
    // the limit now measures a batch we know the host can swallow.
    await withTimeout(context.sync(), BATCH_TIMEOUT_MS, `drawing shapes ${i + 1}-${upTo} of ${total}`);
  }
  return created;
}

/** One chart's committed shapes, awaiting grouping and tagging. */
interface Grouping {
  getSlide: SlideThunk;
  created: PowerPoint.Shape[];
  opts: InsertOptions;
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
async function groupAndTagAll(context: PowerPoint.RequestContext, items: Grouping[]): Promise<void> {
  const tagTargets = items.map((it) => it.created[0] as PowerPoint.Shape | undefined);
  // Grouping is PowerPointApi 1.8+; skip entirely where unsupported.
  const groupable = items
    .map((it, i) => ({ it, i }))
    .filter(({ it }) => it.opts.group !== false && it.created.length > 1);
  if (groupable.length && supports("1.8")) {
    try {
      for (const { it, i } of groupable) {
        // Fresh slide proxy: grouping runs a sync after the render, by which time
        // a held proxy to a new slide could be stale — see SlideThunk.
        const group = (it.getSlide().shapes as unknown as { addGroup(items: PowerPoint.Shape[]): PowerPoint.Shape }).addGroup(it.created);
        group.name = "PowerChart";
        tagTargets[i] = group;
      }
      await context.sync();
    } catch {
      /* grouping failed — shapes stay ungrouped, charts are already on the slide */
      for (const { i } of groupable) tagTargets[i] = items[i].created[0];
    }
  }
  // Tags are PowerPointApi 1.3+; keep the chart re-editable where supported.
  const taggable = items.map((it, i) => ({ it, target: tagTargets[i] })).filter((t) => t.it.opts.tagData && t.target);
  if (taggable.length && supports("1.3")) {
    try {
      for (const { it, target } of taggable) target!.tags.add(CHART_TAG, it.opts.tagData!);
      await context.sync();
    } catch {
      /* tags unavailable — charts are inserted but not re-editable */
    }
  }
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
      shape.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.dash;
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
    line.lineFormat.color = s.stroke;
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
  rect.fill.setSolidColor(s.stroke);
  rect.lineFormat.visible = false;
  try {
    rect.rotation = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
  } catch {
    /* rotation unsupported — line renders horizontally */
  }
  if (s.name) rect.name = s.name;
  return rect;
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
      shape.fill.setSolidColor(n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        shape.lineFormat.color = n.stroke;
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
      else shape.fill.setSolidColor(n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        shape.lineFormat.color = n.stroke;
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
      shape.fill.setSolidColor(n.fill);
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
      shape.fill.setSolidColor(n.fill);
      if (n.stroke && (n.strokeWidth ?? 0) > 0) {
        shape.lineFormat.color = n.stroke;
        shape.lineFormat.weight = n.strokeWidth ?? 1;
      } else {
        shape.lineFormat.visible = false;
      }
      if (n.name) shape.name = n.name;
      return [shape];
    }
    case "wedge":
      return addWedgeFan(shapes, n, dx, dy);
    case "polygon": {
      // No freeform paths in Office.js: draw the outline as connected line
      // segments (translucent fills degrade to outline-only in PowerPoint).
      // These go through addSegment like any other line — passing each edge's
      // bounding box straight to addLine mirrored every up-right edge and gave
      // horizontal ones a zero-height box.
      const created: PowerPoint.Shape[] = [];
      const pts = n.points;
      for (let i = 0; i < pts.length; i++) {
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
      shape.fill.setSolidColor(n.fill);
      shape.lineFormat.visible = false;
      try {
        // Geometric 'triangle' points up (= -90° in scene terms); rotation is
        // exposed from PowerPointApi 1.10 — best effort on older hosts.
        (shape as unknown as { rotation: number }).rotation = box.rotation;
      } catch {
        /* rotation unsupported — arrowhead stays axis-aligned */
      }
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
  font.color = n.color;
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
function addWedgeFan(
  shapes: PowerPoint.ShapeCollection,
  n: WedgeNode,
  dx: number,
  dy: number,
): PowerPoint.Shape[] {
  const created: PowerPoint.Shape[] = [];
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
      shape.fill.setSolidColor(n.fill);
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
        edge.fill.setSolidColor(n.stroke);
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
  return (
    typeof Office !== "undefined" &&
    typeof PowerPoint !== "undefined" &&
    !!Office.context?.host
  );
}
