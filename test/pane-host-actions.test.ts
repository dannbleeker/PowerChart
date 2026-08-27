// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "fs";
import type { ChartConfig } from "../src/core/types";
import { buildChart } from "../src/core/chart";
import { sceneToSvg } from "../src/render/svg";

/**
 * Pane ↔ host command handlers. `pane-state.test.ts` boots app.ts under jsdom
 * with no `Office` global, so it exercises everything EXCEPT the buttons that
 * talk to PowerPoint — Insert, Insert-new, Load-selection, Same-scale. Those
 * handlers (`doInsert`, `doLoadSelection`, `doSameScale`) were the pane's
 * largest untested surface: the branch that chooses a selected placeholder's
 * bounds over the tiled offset, the edit-in-place path, the union-extent maths
 * behind Same scale, and the "not an SSF chart" / "needs two charts" guards.
 *
 * The renderer primitives those handlers call (`insertSceneIntoSlide` et al.)
 * are covered against a fake host in `office-render.test.ts`; here the module is
 * mocked to spies so the test asserts the pane's ORCHESTRATION — which handler
 * fires, with what config — rather than re-testing the shape emitter.
 */

/** Shared mailbox the mocked renderer writes to; reset before each boot. */
/**
 * What `isPowerPointHost()` answers, and WHEN.
 *
 * Hoisted beside `host` because the `vi.mock` factory closes over it. Default
 * `always: true` keeps every existing test unchanged; `lateOffice` reproduces
 * the real ordering — false while app.ts is evaluating at module scope, true by
 * the time `Office.onReady` fires.
 */
const hostReadiness = vi.hoisted(() => {
  let mode: "always" | "late" = "always";
  let ready = false;
  return {
    answer: () => (mode === "always" ? true : ready),
    reset: () => {
      mode = "always";
      ready = false;
    },
    lateOffice: () => {
      mode = "late";
      ready = false;
    },
    officeArrives: () => {
      ready = true;
    },
  };
});

/** Where the mocked `traceEnvironment` sends its line. Pointed at the real `trace`. */
const traceRef = vi.hoisted(() => ({
  fn: (_scope: string, _message: string, _data?: Record<string, unknown>) => {},
}));

const host = vi.hoisted(() => ({
  selectionBounds: null as null | { left: number; top: number; width: number; height: number },
  /** What is already on the slide the next insert would draw onto. */
  /**
   * What `getSlideShapeBounds` answers — `null` for a host that would not read
   * the slide at all, which is a different fact from an empty slide.
   */
  slideShapes: [] as { left: number; top: number; width: number; height: number }[] | null,
  deckCharts: [] as { configJson: string; target: unknown }[],
  /** Slides the deck scan could not read — the signal Same Scale now gates on. */
  deckScanUnread: 0,
  selectionCharts: [] as { configJson: string; target: unknown }[],
  loadSelectionResult: null as null | { configJson: string; target: unknown },
  // When set, insertSceneIntoSlide awaits this before resolving — lets a test
  // observe the pane's mid-flight state (buttons disabled) before the action ends.
  gate: null as null | Promise<void>,
  // insertSceneIntoSlide throws this once, if set — drives the guard's catch path.
  failInsertOnce: false,
  // When set, updateChartInSlide awaits this — holds an in-place update open
  // so a test can act while it is still running.
  updateGate: null as null | Promise<void>,
  // The selection-change listener app.ts registers via addHandlerAsync, captured
  // so a test can fire it the way PowerPoint would.
  selectionListener: null as null | (() => unknown),
  /** The budget each loadChartFromSelection call was given, in order. */
  selectionBudgets: [] as (number | undefined)[],
  /** When set, loadChartFromSelection awaits this — holds a selection read open. */
  selectionGate: null as null | Promise<void>,
  /** What `readDeckStyle` answers — the style this deck carries, or none. */
  deckStyle: null as null | Record<string, unknown>,
  /** When true, the deck-style READ fails — not the same as a deck carrying none. */
  deckStyleUnreadable: false,
  /** Every style handed to `writeDeckStyle`, in order. */
  deckStyleWrites: [] as unknown[],
  /** False for a host below PowerPointApi 1.7, which cannot store one at all. */
  deckStyleWritable: true,
  agendaSlides: [] as unknown[][],
  demoRuns: 0,
  demoDeckCalls: [] as unknown[][],
  /** The runOpts each insertDemoDeck call was given, in order. */
  demoDeckOpts: [] as ({ reconcile?: boolean } | undefined)[],
  /** insertDemoDeck awaits this on the Nth call — lets a test act mid-run. */
  demoDeckGateOn: 0,
  demoDeckGate: null as null | Promise<void>,
  demoDeckPageFailures: new Set<number>(),
  demoDeckStatusOverride: null as null | "rendered" | "failed" | "skipped",
  /** What canInsertPicture() reports — false models a host below PowerPointApi 1.8. */
  canPicture: true,
  /**
   * What updateChartInSlide resolves to. Defaults to undefined because the
   * existing "target is gone" test depends on that; the explode tests opt in to
   * a live EditTarget, which is what a real successful update hands back.
   */
  updateResult: undefined as undefined | { slideId: string; shapeId: string; left: number; top: number; lost?: string },
  /** What a fresh insert hands back — an EditTarget now, so a recovery stays editable. */
  insertResult: null as null | { slideId: string; shapeId: string; left: number; top: number; lost?: string },
  /** Whether the host advertises insertSlidesFromBase64 — off by default. */
  canInsertFile: false,
  slideHoldsOnlyChart: false,
  /** Whether the renderer says this chart is too dense to draw as shapes. */
  autoPicture: false,
  updateChartThrows: false,
  /**
   * Which items of a multi-chart update stall, keyed by index, and the stray
   * shape ids each one left on its slide.
   *
   * The real `updateChartsInSlides` deletes a chart's old shapes before it
   * redraws them, so a stall leaves a partial chart behind and reports it via
   * `onFailed` with the debris attached to the error. Modelling that is the
   * only way a test can see whether the caller sweeps.
   */
  updateChartsStalls: new Map<number, string[]>(),
  /**
   * How many charts `updateChartsInSlides` drops WITHOUT reporting them.
   *
   * The failure the renderer's early filters produce and no callback carries:
   * a chart whose slide or shape the host would not resolve never reaches
   * `onFailed`, so the only way a caller can notice is by counting what came
   * back. A real host dropped five of six this way and the pane said
   * "Same scale applied to 6 charts".
   */
  updateChartsDrops: 0,
  /** Whether the user has pressed Stop — the flag the render loops read. */
  stopRequested: false,
  /** What slideSize() reports — 16:9 unless a test says otherwise. */
  slideSize: { width: 960, height: 540, source: "pageSetup" as const },
  demoReconcile: undefined as unknown,
  /** What `replaceSlideWithDeck` answers: "failed" | "swapped" | "duplicated". */
  swapOutcome: "failed" as "failed" | "swapped" | "duplicated",
  /** How many slides the one-shot insert reports landing; null = all of them. */
  insertFileLands: null as null | number,
  insertFileError: null as null | Error,
  buildFileError: null as null | Error,
  slideCount: 1,
  /**
   * Number of slide-count reads that succeed before the rest throw — a host
   * that answers, then wedges. `null` means it always answers.
   */
  slideCountThrowsAfter: null as null | number,
  slideCountCalls: 0,
  /** What a deck scan reports holding, when the caller asked for the inventory. */
  deckInventory: [] as { slideId: string; index: number; shapes: { id: string; name?: string }[] }[],
  /**
   * The deck's slide ids. `undefined` models a host that will not list them —
   * which is the case the round's id diff has to survive, not an edge case: it
   * is what every blind deck scan in this project's history looked like.
   */
  deckSlideIds: undefined as undefined | string[],
  deckSlideIdsBeforeFails: false,
  /** Whether the host will draw a slide — PowerPointApi 1.8, absent on plenty of hosts. */
  canRaster: true,
  /**
   * The deck scan throws rather than answering short.
   *
   * A host that will not answer and a host that raises are different things,
   * and only the second one can take a round's verdicts down with it. The
   * first version of the guard for that used the first condition and passed
   * with the protection removed.
   */
  deckScanThrows: false,
  deckSlideIdCalls: 0,
  /** Slides the round is modelled as having added, seen by the second id read. */
  roundAddsSlides: ["probe-a", "probe-b"] as string[],
  /**
   * How many charts had been drawn each time the shape selection was dropped.
   *
   * The ORDER is the whole assertion — see office-js#2775. Dropping the
   * selection after the draw would be no protection at all, and a mock that only
   * recorded "it happened" could not tell the two apart.
   */
  selectionDropped: [] as number[],
  /** Slide ids `deleteSlideById` was asked about, in order. */
  deletedSlides: [] as string[],
  /** Slide ids the host will refuse to delete — a refusal is not the same as a delete. */
  refuseSlideDelete: [] as string[],
  reconcileOutcome: undefined as unknown,
  calls: {
    insertScene: [] as { tagData?: string; left?: number; top?: number; slideId?: string; pictureBase64?: string }[],
    updateChart: [] as { target: unknown; opts: { tagData?: string; pictureBase64?: string } }[],
    updateCharts: [] as { scene: unknown; target: unknown; opts?: { tagData?: string } }[][],
    insertFile: [] as { b64: string; expected: number }[],
    deselected: [] as string[][],
    /** Sweeps of the litter a stalled redraw left behind, per recovery. */
    swept: [] as { slideId: string; ids: string[] }[],
  },
}));

vi.mock("../src/render/powerpoint", () => ({
  // FLIPPABLE, to reproduce the ordering the real host has. app.ts asks this at
  // MODULE SCOPE, long before `Office.context` exists, and asked it only once.
  isPowerPointHost: () => hostReadiness.answer(),
  // The round-start line evaluates this as an argument, so a mock without it
  // throws and the line vanishes — which is how eight tests in this file went
  // red at once when it was added. Kept minimal on purpose: the pane must not
  // depend on any field of it.
  roundEnvironment: () => ({ requirementSets: [] }),
  // Same reason: the battery stamps every scenario with the host-friction
  // delta, so it is on the round's critical path now. (`deckSlideIds` is
  // already mocked further down, with the round's own growth modelled.)
  // FOUR NAMES HERE, EIGHT IN THE REAL THING. This listed the counters that
  // existed when it was written, so a caller reading any of the four added
  // since got `undefined` from the double and a number from production — a
  // fake that cannot fail, which this repo has been bitten by before.
  //
  // A `vi.mock` factory is hoisted and cannot read the real module, so the list
  // stays literal. `the pane's friction double carries every counter` below
  // compares it against `vi.importActual` and goes red the day a ninth is added.
  hostFrictionCounts: () => ({
    errors: 0,
    idRefusals: 0,
    generalExceptions: 0,
    emptyReReads: 0,
    shortReReads: 0,
    unmatchedReReads: 0,
    reReadsRepaired: 0,
  }),
  canInsertPicture: vi.fn(() => host.canPicture),
  getSelectionBounds: vi.fn(async () => host.selectionBounds),
  dropShapeSelection: vi.fn(async () => {
    host.selectionDropped.push(host.calls.insertScene.length);
    return true;
  }),
  getSlideShapeBounds: vi.fn(async () => host.slideShapes),
  insertSceneIntoSlide: vi.fn(
    async (
      scene: { nodes: unknown[] },
      opts: { tagData?: string; left?: number; top?: number },
      onPhase?: (phase: string, detail?: string) => void,
    ) => {
      // Report phases the way the real renderer does. The mock used to ignore
      // `onPhase` entirely, and that silence is what hid a live bug for good:
      // with no phase notes the pane's note stayed exactly "Working…", which is
      // the one string the old `guard` recognised as "nothing was said", so
      // every test saw the "Done." a real PowerPoint never printed.
      onPhase?.("context");
      onPhase?.("queue", `${scene.nodes.length} nodes`);
      onPhase?.("commit", `${scene.nodes.length} of ${scene.nodes.length} shapes`);
      onPhase?.("group");
      if (host.gate) await host.gate;
      if (host.failInsertOnce) {
        host.failInsertOnce = false;
        throw new Error("host refused the insert");
      }
      host.calls.insertScene.push(opts);
      // The last thing a successful insert ever says — and it is still "busy".
      onPhase?.("done");
      return host.insertResult;
    },
  ),
  updateChartInSlide: vi.fn(
    async (_scene: unknown, target: unknown, opts: { tagData?: string; pictureBase64?: string }) => {
      host.calls.updateChart.push({ target, opts });
      if (host.updateGate) await host.updateGate;
      // A stop taken while this was in flight surfaces the way the renderer
      // surfaces it: a marked error carrying whatever the halted redraw had
      // already committed. Stopping mid-batch leaves the same debris a stall
      // does, so the wreckage is not optional.
      if (host.stopRequested) {
        const err = new Error("Stopped.") as Error & Record<string, unknown>;
        err.__powerchartStopped = true;
        err.__powerchartWreckage = { slideId: "s1", at: { left: 40, top: 50 }, strayIds: ["stray-1"] };
        throw err;
      }
      // Models the live-canvas stall: the in-place redraw refuses, a picture
      // update (which draws one shape) does not.
      if (host.updateChartThrows && !opts.pictureBase64) {
        // Carrying wreckage, because the real one always does: an update deletes
        // the old chart and COMMITS that before it draws a single new shape, so
        // a stall mid-redraw is never a no-op. A double that threw a bare error
        // modelled a rollback the host does not have, and every fallback below
        // was written against that fiction.
        const err = new Error("did not respond while drawing shapes 1-10") as Error & Record<string, unknown>;
        err.__powerchartWreckage = { slideId: "s1", at: { left: 40, top: 50 }, strayIds: ["stray-1", "stray-2"] };
        throw err;
      }
      return host.updateResult;
    },
  ),
  wreckageOf: (err: unknown) =>
    err && typeof err === "object" && Object.prototype.hasOwnProperty.call(err, "__powerchartWreckage")
      ? (err as Record<string, unknown>).__powerchartWreckage
      : undefined,
  deleteShapesById: vi.fn(async (slideId: string, ids: string[]) => {
    host.calls.swept.push({ slideId, ids });
    return ids.length;
  }),
  // The cooperative stop. Modelled with the real semantics — a flag the render
  // loops read — so a test can press Stop mid-action and see what the pane does
  // with it, rather than only that the button exists.
  requestStop: vi.fn(() => {
    host.stopRequested = true;
  }),
  resetStop: vi.fn(() => {
    host.stopRequested = false;
  }),
  isStopRequested: vi.fn(() => host.stopRequested),
  // The destination deck's slide size, which the deck builder needs so the
  // generated file declares the size it is being inserted into.
  slideSize: vi.fn(async () => host.slideSize),
  isStopped: (err: unknown) =>
    !!err && typeof err === "object" && Object.prototype.hasOwnProperty.call(err, "__powerchartStopped"),
  updateChartsInSlides: vi.fn(
    async (
      items: { scene: unknown; target: { slideId?: string }; opts?: { tagData?: string } }[],
      onFailed?: (item: unknown, err: unknown) => void,
    ) => {
      host.calls.updateCharts.push(items);
      // Report the stalls the test asked for, exactly as the renderer does:
      // the error carries the wreckage, because that is the only channel the
      // caller has for finding out what was left on the slide.
      for (const [i, strayIds] of host.updateChartsStalls) {
        const item = items[i];
        if (!item) continue;
        const err = new Error("host stalled") as Error & Record<string, unknown>;
        err.__powerchartWreckage = {
          slideId: item.target?.slideId ?? `s${i}`,
          at: { left: 0, top: 0 },
          strayIds,
        };
        onFailed?.(item, err);
      }
      // One target per chart the host RESOLVED — the real contract, and the
      // only channel a caller has for "this chart was dropped without a word".
      // The renderer's `live`/`alive` filters drop a chart whose slide or shape
      // the host declines to answer for, and they are early returns rather than
      // throws, so `onFailed` never fires for them. A stalled chart is still in
      // this array (it comes back carrying its OLD target), which is why the two
      // signals do not double-count. `updateChartsDrops` is how a test arms the
      // silent case.
      return items.slice(host.updateChartsDrops).map((it) => it.target);
    },
  ),
  listChartsInDeck: vi.fn(async (opts: { withInventory?: boolean } = {}) => {
    if (host.deckScanThrows) throw new Error("the host refused to describe the deck");
    return {
      charts: host.deckCharts,
      unread: host.deckScanUnread,
      short: 0,
      tagsUnread: 0,
      slides: host.slideCount,
      // Only when asked, exactly like the real one — a mock that always returns
      // the inventory would let a caller that forgot to ask still find it.
      ...(opts.withInventory ? { inventory: host.deckInventory } : {}),
    };
  }),
  // The deck GROWS across a round, because that is the only thing the id diff
  // is measuring. A double that answered the same list twice would make the
  // diff empty and every assertion about it vacuous.
  deckSlideIds: vi.fn(async () => {
    const first = host.deckSlideIdCalls++ === 0;
    // ONLY THE BEFORE-READ FAILS, which is the state that matters and the one a
    // double answering `undefined` to BOTH calls cannot express. With both
    // unreadable the old code and the fixed code agree on an empty list, so a
    // test built that way passes with the fix deleted — it did, before this
    // existed.
    if (host.deckSlideIdsBeforeFails && first) return undefined;
    if (!host.deckSlideIds) return undefined;
    return first ? host.deckSlideIds : [...host.deckSlideIds, ...host.roundAddsSlides];
  }),
  slideShots: vi.fn(async (ids: string[], opts: { max?: number } = {}) =>
    ids.map((slideId, i) => (i < (opts.max ?? 12) && host.canRaster ? { slideId, png: "UE5H" } : { slideId })),
  ),
  scanIsComplete: (s: { unread: number; short: number; tagsUnread: number }) =>
    s.unread === 0 && s.short === 0 && s.tagsUnread === 0,
  scanGap: (s: { unread: number; slides: number }) =>
    s.unread ? `${s.unread} of ${s.slides} slide(s) would not answer` : "",
  listChartsInSelection: vi.fn(async () => host.selectionCharts),
  loadChartFromSelection: vi.fn(async (budgetMs?: number) => {
    host.selectionBudgets.push(budgetMs);
    if (host.selectionGate) await host.selectionGate;
    return host.loadSelectionResult;
  }),
  insertAgendaSlides: vi.fn(async (scenes: unknown[][]) => {
    host.agendaSlides.push(scenes);
  }),
  insertDemoDeck: vi.fn(
    async (
      items: { title?: string; scene: { nodes: unknown[] } }[],
      _onProgress?: unknown,
      runOpts?: { reconcile?: boolean },
    ) => {
      host.demoRuns++;
      const call = host.demoRuns;
      host.demoDeckCalls.push(items);
      host.demoDeckOpts.push(runOpts);
      if (host.demoDeckGate && call === host.demoDeckGateOn) await host.demoDeckGate;
      if (host.demoDeckPageFailures.has(call)) throw new Error(`page ${call} refused`);
      // Fabricated report — each item counted as failed so the pane's results
      // path has rows to render. Callers that need a specific status set
      // demoDeckStatusOverride.
      const status = host.demoDeckStatusOverride ?? "failed";
      const results = items.map(() => ({ created: 5, status, ms: 100 }));
      return {
        // The run's identity, as the real renderer reports it — the join key
        // between the downloaded log and the .pptx it produced. Fabricated
        // per call so a test can tell two runs apart.
        run: `fake-run-${call}`,
        results,
        slidesAdded: items.length,
        addsIssued: items.length,
        blankSlides: [],
        blankItems: [],
        blanksRead: true,
        reconcile: host.demoReconcile,
        totalMs: items.length * 100,
      };
    },
  ),
  loadThemePalette: vi.fn(async () => null),
  // The deck's own style file. `null` is the ordinary answer — a deck nobody
  // has branded — and the pane must be usable before this resolves, which is
  // what the startup read being fire-and-forget is for.
  readDeckStyle: vi.fn(async () => host.deckStyle),
  // The same read, saying why it came back empty. The button awaits THIS one,
  // because "the read failed" and "the deck carries none" send the person at
  // the pane to opposite conclusions about their own document.
  // The pane warms the custom-XML surface before reading it — a call whose
  // failure is the point. Absent from this mock, app.ts throws at boot and every
  // test in this file dies, which is how 110 of them went red at once.
  // Found by `pane-mock-covers-the-renderer` on its first run: app.ts:2864 calls
  // this and the mock had no such key, so any pane test reaching that line would
  // have taken all 110 down with it. A landmine, not yet a crater.
  enrichSnapshots: vi.fn(async () => {}),
  warmCustomXmlSurface: vi.fn(async () => {}),
  readDeckStyleWithReason: vi.fn(async () => ({
    style: host.deckStyleUnreadable ? null : host.deckStyle,
    unreadable: host.deckStyleUnreadable,
  })),
  writeDeckStyle: vi.fn(async (style: unknown) => {
    host.deckStyleWrites.push(style);
    return host.deckStyleWritable;
  }),
  onLateSync: vi.fn(),
  errorText: (e: unknown) => String(e),
  // The one-shot deck path. Off by default so the existing cases keep
  // exercising the shape-by-shape renderer they were written for.
  canInsertSlidesFromBase64: vi.fn(() => host.canInsertFile),
  // EMITS, rather than doing nothing. A no-op here is why 69 rounds shipped with
  // `traceEnvironment` on the wrong side of the round's trace mark and no test
  // could see it: the defect is WHERE this is called, and a stub that traces
  // nothing cannot express where. `traceRef.fn` is pointed at the real `trace`
  // by the test that cares.
  traceEnvironment: vi.fn((build: string) => traceRef.fn("host", "environment", { build })),
  wantsAutoPicture: vi.fn(() => host.autoPicture),
  // Selection juggling around an in-place redraw. The fake host has no view,
  // so it just runs the body — with `deselected` false, which is the honest
  // answer for a host that cannot move the selection.
  withSlideDeselected: vi.fn(async (ids: string[], fn: (d: boolean) => Promise<unknown>) => {
    host.calls.deselected.push(ids);
    return fn(false);
  }),
  slideHoldsOnlyChart: vi.fn(async () => host.slideHoldsOnlyChart),
  // Three outcomes, not two: "duplicated" is the swap that inserted the new
  // slide but could not remove the old one.
  replaceSlideWithDeck: vi.fn(async () => host.swapOutcome),
  newRunId: vi.fn(() => "run-under-test"),
  OFFSCREEN_BATCH: 40,
  insertSlidesFromPptx: vi.fn(async (b64: string, expected: number) => {
    host.calls.insertFile.push({ b64, expected });
    if (host.insertFileError) throw host.insertFileError;
    return host.insertFileLands ?? expected;
  }),
  reconcileDeck: vi.fn(async () => host.reconcileOutcome),
  applyReconcilePlan: vi.fn(async () => host.reconcileOutcome),
  snapshotAddedSlides: vi.fn(async () => []),
  readAddedSlides: vi.fn(async () => ({ snapshots: [], unread: 0 })),
  slideCount: vi.fn(async () => {
    host.slideCountCalls++;
    if (host.slideCountThrowsAfter !== null && host.slideCountCalls > host.slideCountThrowsAfter) {
      throw new Error("host would not report the slide count");
    }
    return host.slideCount;
  }),
  // What the host probe needs, and only that. The probe's own answers are
  // covered exhaustively in `host-probe.test.ts` against the real fake host;
  // what this file cares about is that "Run the whole round" puts a COMPLETE
  // sheet and the self-test's verdicts into ONE file. A host that will not give
  // out a scratch slide produces a complete sheet of `no-scratch-slide`, which
  // is exactly enough to test the bundling and nothing more.
  addScratchSlide: vi.fn(async () => null),
  // No chart this host has agreed to name.
  //
  // This mock's whole point is a host that refuses a scratch slide, so every
  // question comes back `no-scratch-slide` — not `no-scratch-shape` — and the
  // round's re-ask has nothing to re-ask. Which is the honest state for a mock
  // that never draws: `namedShape` is set by a chart whose TAG WROTE, and
  // nothing here writes one.
  namedShape: vi.fn(() => null),
  // No chart to observe either. This mock refuses a scratch slide and draws
  // nothing, so a scan finds no chart and the round's re-ask has nothing to ask
  // with — which is the honest state rather than a stubbed convenience.
  refreshNamedShapeFromDeck: vi.fn(async () => null),
  deleteSlideById: vi.fn(async (id: string) => {
    host.deletedSlides.push(id);
    return !host.refuseSlideDelete.includes(id);
  }),
  isTimeout: vi.fn(() => false),
  requirementSets: vi.fn(() => ["1.1", "1.5"]),
  ScratchSlideUnavailable: class ScratchSlideUnavailable extends Error {},
  withProbeContext: vi.fn(async () => {
    throw new Error("no scratch slide in this harness");
  }),
  deadlinesFired: 0,
  // The other half of a stall report: a scenario that gives up asks whether the
  // call ever came back. This harness has no host to answer late, so the double
  // says "nothing arrived" — which is the reading every real round has produced.
  lastLateSync: null,
  lastLateSyncSeq: 0,
  waitForLateSync: vi.fn(async () => false),
}));

// The deck builder is a real pptxgenjs run; the pane's own tests care about
// which path was taken, not about the bytes (test/pptx-deck.test.ts covers those).
vi.mock("../src/render/pptx-deck", () => ({
  buildDeckBase64: vi.fn(async (items: unknown[]) => {
    if (host.buildFileError) throw host.buildFileError;
    // The real builder reports what it actually drew per slide — the pane
    // verifies against THAT, not against the Office.js shape estimate.
    return { base64: "UEsDBBQA-fake-base64", shapesPerSlide: items.map(() => 7) };
  }),
}));

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Let the clicked handler's async chain (and its busy→done note) settle. */
const settle = () => new Promise((r) => setTimeout(r, 0));

/**
 * Boot app.ts with a PowerPoint host present. app.ts wires the host buttons only
 * inside `Office.onReady`, so the stub must both look like a host (mocked
 * `isPowerPointHost`) and fire onReady synchronously at import.
 */
async function bootHostPane(opts?: { deckStyle?: Record<string, unknown> | null }) {
  host.selectionBounds = null;
  host.slideShapes = [];
  host.deckCharts = [];
  host.deckScanUnread = 0;
  host.deckInventory = [];
  host.deckSlideIds = undefined;
  host.canRaster = true;
  host.deckScanThrows = false;
  host.deletedSlides = [];
  host.selectionDropped = [];
  host.deckSlideIdCalls = 0;
  host.deckSlideIdsBeforeFails = false;
  host.roundAddsSlides = ["probe-a", "probe-b"];
  host.refuseSlideDelete = [];
  host.selectionCharts = [];
  host.loadSelectionResult = null;
  host.gate = null;
  host.failInsertOnce = false;
  host.selectionListener = null;
  host.selectionBudgets = [];
  host.selectionGate = null;
  host.agendaSlides = [];
  // Read at BOOT (the pane asks the deck on `Office.onReady`), so a test that
  // wants a branded deck has to say so here rather than after the fact. The
  // parameter is optional because `beforeEach(bootHostPane)` hands this a test
  // context, which carries no `deckStyle` and so resets it like everything else.
  host.deckStyle = opts?.deckStyle ?? null;
  host.deckStyleWrites = [];
  host.deckStyleWritable = true;
  // Reset with its siblings, or the one test that sets it leaves every later
  // test in this file reading a host that will not answer — an ordering-
  // dependent failure, which is the worst kind to debug.
  host.deckStyleUnreadable = false;
  host.demoRuns = 0;
  host.demoDeckCalls = [];
  host.demoDeckOpts = [];
  host.demoDeckGate = null;
  host.demoDeckGateOn = 0;
  host.demoDeckPageFailures = new Set();
  host.demoDeckStatusOverride = null;
  host.canInsertFile = false;
  host.slideHoldsOnlyChart = false;
  host.updateChartThrows = false;
  host.demoReconcile = undefined;
  host.swapOutcome = "failed";
  host.autoPicture = false;
  host.updateGate = null;
  host.slideCountThrowsAfter = null;
  host.slideCountCalls = 0;
  host.insertFileLands = null;
  host.insertFileError = null;
  host.buildFileError = null;
  host.slideCount = 1;
  host.reconcileOutcome = undefined;
  host.calls.insertFile.length = 0;
  host.calls.deselected.length = 0;
  host.canPicture = true;
  host.updateResult = undefined;
  host.insertResult = null;
  host.calls.swept.length = 0;
  host.calls.insertScene = [];
  host.calls.updateChart = [];
  host.calls.updateCharts = [];
  host.updateChartsStalls.clear();
  host.updateChartsDrops = 0;
  host.stopRequested = false;
  host.slideSize = { width: 960, height: 540, source: "pageSetup" };

  window.history.replaceState({}, "", "/taskpane.html");
  const parsed = new DOMParser().parseFromString(readFileSync("src/taskpane/taskpane.html", "utf8"), "text/html");
  parsed.querySelectorAll("script").forEach((s) => s.remove());
  document.body.innerHTML = parsed.body.innerHTML;

  vi.stubGlobal("Office", {
    // OFFICE BECOMES AVAILABLE HERE, not before — which is the whole ordering
    // the real host has and this harness used to paper over.
    onReady: (cb: () => void) => {
      hostReadiness.officeArrives();
      cb();
    },
    EventType: { DocumentSelectionChanged: "DocumentSelectionChanged" },
    context: {
      host: "PowerPoint",
      displayLanguage: "en-US",
      document: {
        // watchSelection() registers a selection listener; capture it so a test
        // can fire it the way PowerPoint does when the user clicks a shape.
        addHandlerAsync: (_type: string, handler: () => unknown) => {
          host.selectionListener = handler;
        },
      },
    },
  });

  vi.resetModules();
  await import("../src/taskpane/app");
  await settle();
}

/** A value-axis chart config with a known extent, as a JSON tag would carry it. */
const chartJson = (values: number[]): string =>
  JSON.stringify({
    kind: "stacked",
    width: 480,
    height: 320,
    data: { categories: values.map((_, i) => `C${i}`), series: [{ name: "S", values }] },
  } satisfies ChartConfig);

afterEach(() => vi.unstubAllGlobals());

/**
 * Stand in for the browser rasteriser. jsdom decodes no SVG, `getContext("2d")`
 * returns null and `toDataURL` returns null, so every step of `rasterizeScene`
 * needs a double. Pass `{ decode: false }` to model a browser that cannot decode
 * the SVG, or `{ encode: false }` for one whose canvas cannot produce a PNG.
 * Returns a restore function.
 */
function stubRaster({ decode = true, encode = true } = {}) {
  class FakeImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    set src(_v: string) {
      Promise.resolve().then(() => (decode ? this.onload?.() : this.onerror?.()));
    }
  }
  vi.stubGlobal("Image", FakeImage);
  const ctx = vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue({
    scale() {},
    drawImage() {},
  } as unknown as CanvasRenderingContext2D);
  const url = vi
    .spyOn(HTMLCanvasElement.prototype, "toDataURL")
    .mockReturnValue(encode ? "data:image/png;base64,RASTER" : (null as unknown as string));
  // Capture the SVG each rasterisation was handed. Without this a test can only
  // see THAT a payload was produced, never OF WHAT — and the stub returns the
  // same bytes whatever it is given, so "a payload exists" passes even when the
  // wrong scene was rasterised.
  const svgs: Blob[] = [];
  const c = vi.spyOn(URL, "createObjectURL").mockImplementation((blob: Blob | MediaSource) => {
    if (blob instanceof Blob) svgs.push(blob);
    return "blob:x";
  });
  const r = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
  return {
    rasterizedSvgs: () => Promise.all(svgs.map((b) => b.text())),
    restore: () => {
      for (const s of [ctx, url, c, r]) s.mockRestore();
    },
  };
}

/** Tick "Insert as picture", the way a user does. */
const tickPicture = () => {
  const box = $("render-image") as HTMLInputElement;
  box.checked = true;
  box.dispatchEvent(new Event("change"));
};

/** Capture whatever `downloadJson` hands to the browser. */
function captureDownloads() {
  const blobs: Blob[] = [];
  const c = vi.spyOn(URL, "createObjectURL").mockImplementation((b: Blob | MediaSource) => {
    if (b instanceof Blob) blobs.push(b);
    return "blob:x";
  });
  const r = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
  return {
    lastJson: async () => JSON.parse(await blobs[blobs.length - 1].text()),
    count: () => blobs.length,
    restore: () => {
      c.mockRestore();
      r.mockRestore();
    },
  };
}

describe("image render mode", () => {
  beforeEach(bootHostPane);

  it("hands the insert a rasterised PNG when Insert as picture is ticked", async () => {
    const raster = stubRaster();
    try {
      tickPicture();
      $("insert").click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(1);
      // The payload reached the renderer, bare — the pane strips the data: URI.
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBe("RASTER");
      // And the config it tagged says image mode, so a re-edit knows what it is.
      expect(JSON.parse(host.calls.insertScene[0].tagData!).render).toBe("image");
    } finally {
      raster.restore();
    }
  });

  it("sends no payload at all in shapes mode, so nothing pays for the canvas", async () => {
    const raster = stubRaster();
    try {
      $("insert").click();
      await settle();
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
      expect(JSON.parse(host.calls.insertScene[0].tagData!).render).toBeUndefined();
    } finally {
      raster.restore();
    }
  });

  it("skips the raster and SAYS SO on a host below PowerPointApi 1.8", async () => {
    // The note must survive to the end: insertSceneIntoSlide's first act is
    // onPhase("context"), which overwrites the pane note, so a warning posted
    // BEFORE the insert would be wiped. It is posted after, on purpose.
    const raster = stubRaster();
    try {
      host.canPicture = false;
      tickPicture();
      $("insert").click();
      await settle();
      expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
      expect($("host-note").textContent).toMatch(/can't insert pictures/i);
    } finally {
      raster.restore();
    }
  });

  it("degrades to native shapes, visibly, when the browser cannot rasterise", async () => {
    for (const mode of [{ decode: false }, { encode: false }] as const) {
      await bootHostPane();
      const raster = stubRaster(mode);
      try {
        tickPicture();
        $("insert").click();
        await settle();
        // The chart still landed — as shapes, with no payload.
        expect(host.calls.insertScene, JSON.stringify(mode)).toHaveLength(1);
        expect((host.calls.insertScene[0] as { pictureBase64?: string }).pictureBase64).toBeUndefined();
        expect($("host-note").textContent, JSON.stringify(mode)).toMatch(/rasterize/i);
      } finally {
        raster.restore();
      }
    }
  });

  it("batch insert honours render per config, not the checkbox", async () => {
    // #json-insert-batch calls insertSceneIntoSlide directly and used to ignore
    // the mode outright, so a pasted image-mode array silently drew shapes.
    const raster = stubRaster();
    try {
      ($("json-io") as HTMLTextAreaElement).value = JSON.stringify([
        { kind: "stacked", render: "image", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } },
        { kind: "stacked", data: { categories: ["A"], series: [{ name: "S", values: [1] }] } },
      ]);
      $("json-insert-batch").click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(2);
      const payloads = host.calls.insertScene.map((o) => (o as { pictureBase64?: string }).pictureBase64);
      expect(payloads).toEqual(["RASTER", undefined]);
    } finally {
      raster.restore();
    }
  });

  it("Same scale rasterises the RESCALED chart, not the one it loaded", async () => {
    // The scale mutation happens inside the map that builds each scene, so
    // rasterising in the same pass captures the PRE-scale chart: the picture
    // shows the old axis while the pane reports success.
    //
    // Asserting "a payload exists" does NOT catch that — the stub returns the
    // same bytes for any input, and tagData reads the same mutated object either
    // way. So compare the SVG actually handed to the rasteriser against the two
    // candidate renders. A chart whose own extent is [1,5] rescaled to max 90
    // draws visibly different bars, so the two SVGs differ.
    const raster = stubRaster();
    try {
      const own = JSON.parse(chartJson([1, 5])) as ChartConfig;
      host.deckCharts = [
        { configJson: JSON.stringify({ ...own, render: "image" }), target: { shapeId: "a" } },
        { configJson: chartJson([2, 90]), target: { shapeId: "b" } },
      ];
      $("same-scale").click();
      await settle();

      const [rasterized] = await raster.rasterizedSvgs();
      expect(rasterized, "exactly one chart was rasterised").toBeTruthy();
      const svgOf = (cfg: ChartConfig) => sceneToSvg(buildChart(cfg));
      const preScale = svgOf({ ...own, render: "image" });
      const postScale = svgOf({ ...own, render: "image", scale: { min: undefined, max: 90 } });
      expect(preScale, "the two candidates must actually differ").not.toBe(postScale);
      expect(rasterized).toBe(postScale);
      expect(rasterized).not.toBe(preScale);
    } finally {
      raster.restore();
    }
  });

  it("does not report a successful auto-picture rescue as a fallback to shapes", async () => {
    // `chartPicture` returns a warn WITH a png on its SUCCESS path: a chart too
    // dense for the web host, rasterised and inserted as a picture, the warn
    // being the explanation. Treating every warn as a failure printed, in red,
    // "N image chart(s) fell back to native shapes" — over charts that went in
    // as PICTURES, overwriting the true success line, claiming the dangerous
    // thing had happened when the guard against it had just worked, and
    // steering the user away from "Explode to native shapes".
    const raster = stubRaster();
    try {
      host.autoPicture = true;
      host.canPicture = true;
      host.deckCharts = [
        { configJson: chartJson([1, 5]), target: { shapeId: "a" } },
        { configJson: chartJson([2, 90]), target: { shapeId: "b" } },
      ];
      $("same-scale").click();
      await settle();
      // Both really are pictures — the premise of the complaint.
      const payloads = host.calls.updateCharts.at(-1) ?? [];
      expect(payloads.length, "no update was issued").toBeGreaterThan(0);
      expect(
        payloads.every((it) => !!(it.opts as { pictureBase64?: string } | undefined)?.pictureBase64),
        "the rescue did not produce pictures, so this proves nothing",
      ).toBe(true);
      expect($("host-note").textContent ?? "").not.toMatch(/fell back to native shapes/i);
      expect($("host-note").className, "a successful rescue was reported in red").not.toMatch(/status-err/);
    } finally {
      raster.restore();
    }
  });
});

describe("Explode to native shapes", () => {
  beforeEach(bootHostPane);

  it("re-draws the selected picture chart as shapes by OMITTING the payload", async () => {
    // The omission is the explode: no pictureBase64 means wantsPicture is false
    // in the renderer, so the node loop runs. Nothing else about the update
    // differs, which is why this needs no renderer code.
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20 };
    $("explode").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(1);
    const { opts } = host.calls.updateChart[0];
    expect((opts as { pictureBase64?: string }).pictureBase64).toBeUndefined();
    // The re-saved config says shapes, so the chart does not turn back into a
    // picture on the next ordinary update.
    expect(JSON.parse(opts.tagData!).render).toBe("shapes");
    expect($("host-note").textContent).toMatch(/native shapes/i);
    // The checkbox follows the exploded config — the pane must not still claim
    // the next insert will be a picture.
    expect(($("render-image") as HTMLInputElement).checked).toBe(false);
  });

  it("refuses politely when the selection is not an SSF chart", async () => {
    host.loadSelectionResult = null;
    $("explode").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(0);
    expect($("host-note").textContent).toMatch(/select an inserted chart first/i);
  });

  it("does not claim success when the host would not save the config back", async () => {
    // The same host answer the ordinary update path handles carefully — and on
    // the web it is ordinary: five `no-config` results in a single recorded
    // round. Explode used to print GREEN over a chart that could no longer be
    // opened, and then adopt the lost target as the live edit target, so the
    // next push resolved a dead id and did nothing at all.
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20, lost: "no-config" };
    $("explode").click();
    await settle();
    expect($("host-note").className, "a lost chart was reported as a success").toMatch(/status-err/);
    expect($("host-note").textContent).toMatch(/no longer editable/i);
    // And the dead target is not kept: the Insert button must not offer Update
    // against a chart the pane cannot resolve.
    expect($("insert").textContent).not.toMatch(/update/i);
  });

  it("does not claim success when the host would not say where the shapes landed", async () => {
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20, lost: "unknown-shape" };
    $("explode").click();
    await settle();
    expect($("host-note").className).toMatch(/status-err/);
    expect($("host-note").textContent).toMatch(/lost track/i);
  });

  it("says so when the picture is already gone from the slide", async () => {
    host.loadSelectionResult = {
      configJson: JSON.stringify({ ...JSON.parse(chartJson([1, 2])), render: "image" }),
      target: { slideId: "s1", shapeId: "pic-1", left: 10, top: 20 },
    };
    host.updateResult = undefined; // the update found nothing to replace
    $("explode").click();
    await settle();
    expect($("host-note").textContent).toMatch(/no longer on the slide/i);
  });
});

describe("Insert", () => {
  beforeEach(bootHostPane);

  it("tiles a new chart at the default offset when nothing is selected", async () => {
    host.selectionBounds = null;
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    const opts = host.calls.insertScene[0];
    // Default drop point, and the config round-trips as the shape's tag.
    expect(opts.left).toBeGreaterThanOrEqual(60);
    expect(opts.top).toBeGreaterThanOrEqual(90);
    expect(JSON.parse(opts.tagData!)).toMatchObject({ kind: expect.any(String) });
    expect(host.calls.updateChart).toHaveLength(0);
  });

  /**
   * Two charts onto one slide. The insert path never looked at what was
   * already there — its whole answer to "where does this go" was a 14pt
   * cascade, and against a 480x300 chart that is a 90%-plus overlap. The
   * second chart landed on the first, which is what a user sees as one chart
   * drawn over another.
   */
  it("puts a second chart clear of the first instead of on top of it", async () => {
    const first = { left: 60, top: 90, width: 480, height: 300 };
    host.slideShapes = [first];
    $("insert").click();
    await settle();
    const at = host.calls.insertScene.at(-1)!;
    const cfg = JSON.parse(at.tagData!) as { width?: number; height?: number };
    const box = { left: at.left!, top: at.top!, width: cfg.width ?? 480, height: cfg.height ?? 300 };
    const hits =
      box.left < first.left + first.width &&
      first.left < box.left + box.width &&
      box.top < first.top + first.height &&
      first.top < box.top + box.height;
    expect(hits).toBe(false);
  });

  it("places the second chart BESIDE the first on a 16:9 deck", async () => {
    // The slide WIDTH decides this, and it is the number the pane spent its
    // life assuming. On 16:9 there is a 390pt band beside the first chart.
    const first = { left: 60, top: 90, width: 480, height: 300 };
    host.slideShapes = [first];
    host.slideSize = { width: 960, height: 540, source: "pageSetup" };
    $("insert").click();
    await settle();
    const at = host.calls.insertScene.at(-1)!;
    expect(at.left, "did not sit beside the first chart").toBeGreaterThanOrEqual(first.left + first.width);
    expect(at.top).toBe(first.top);
    expect(at.left! + (JSON.parse(at.tagData!) as { width: number }).width).toBeLessThanOrEqual(960);
    expect($("host-note").textContent?.toLowerCase()).toContain("beside");
  });

  it("drops the second chart BELOW instead on a 4:3 deck", async () => {
    // Same two charts, 240pt less width. Beside leaves 150pt, which scales the
    // chart under the readability floor — so down it goes. A pane that assumed
    // 16:9 here would run a chart off the right edge of every 4:3 deck.
    //
    // A fresh pane, not a second insert into the previous test's: doInsert
    // feeds the placed size back into the config it tags, so a follow-on insert
    // starts from the SHRUNK size and would be measuring two variables at once.
    const first = { left: 60, top: 90, width: 480, height: 300 };
    host.slideShapes = [first];
    host.slideSize = { width: 720, height: 540, source: "pageSetup" };
    $("insert").click();
    await settle();
    const at = host.calls.insertScene.at(-1)!;
    expect(at.left, "squeezed a chart beside on a 4:3 deck").toBe(first.left);
    expect(at.top, "did not drop below the first chart").toBeGreaterThanOrEqual(first.top + first.height);
    // And it stays on the slide, which is the risk horizontal placement adds.
    expect(at.left! + (JSON.parse(at.tagData!) as { width: number }).width).toBeLessThanOrEqual(720);
  });

  /**
   * A host that will not describe the slide must get the CASCADE, and the
   * cascade has to actually move.
   *
   * `getSlideShapeBounds` used to swallow a refusal into `[]`, with a comment
   * claiming placement then "falls back to the cascade, which is what it always
   * did". It does not: `placeChart` reads an empty `occupied` as "there is room
   * everywhere", so `placeBeside` succeeds on its first pass and returns the
   * origin UNMOVED — its `fallback` argument is unreachable. Two inserts onto a
   * slide the host would not read therefore landed on exactly the same point,
   * which is the pile the placement rule exists to prevent and worse than the
   * fixed cascade it replaced. A real host refused every shape read on a whole
   * deck (`unread=8 slides=8`), so this is its ordinary behaviour.
   */
  it("cascades instead of stacking when the host will not say what is on the slide", async () => {
    host.slideShapes = null;
    $("insert").click();
    await settle();
    const first = host.calls.insertScene.at(-1)!;
    $("insert").click();
    await settle();
    const second = host.calls.insertScene.at(-1)!;
    expect(host.calls.insertScene).toHaveLength(2);
    expect({ left: second.left, top: second.top }, "two charts landed on exactly the same point").not.toEqual({
      left: first.left,
      top: first.top,
    });
  });

  it("fits the chart to a selected placeholder's bounds", async () => {
    host.selectionBounds = { left: 200, top: 150, width: 360, height: 240 };
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    const opts = host.calls.insertScene[0];
    expect(opts.left).toBe(200);
    expect(opts.top).toBe(150);
    // The placeholder's size overrides the config's own dimensions.
    const cfg = JSON.parse(opts.tagData!) as ChartConfig;
    expect(cfg.width).toBe(360);
    expect(cfg.height).toBe(240);
  });

  /**
   * The user's own shape, not destroyed by inserting a chart beside it.
   *
   * office-js#2775: on PowerPoint on the web, adding a text box deletes the
   * shape that was selected. Every chart drawn here contains text boxes, and
   * this path deliberately leaves the selection alone because that is how it
   * learns where to put the chart — so a picture selected to position a chart
   * against would be silently destroyed by the insert. office-js#3698 is the
   * same setup failing the other way: a picture cannot be inserted while a
   * shape is selected, and this path inserts a picture for a dense chart.
   */
  it("lets go of the user's selected shape before it draws anything", async () => {
    host.selectionBounds = { left: 200, top: 150, width: 360, height: 240 };
    $("insert").click();
    await settle();
    // The bounds still decide the placement — reading the selection and holding
    // onto it are different things, and only the second one is the hazard.
    expect(host.calls.insertScene[0].left).toBe(200);
    // BEFORE the draw, which is the entire protection. `[0]` is "no charts had
    // been drawn yet when the selection was dropped"; a `[1]` here would be a
    // pane that let go after the damage.
    expect(host.selectionDropped, "drew with the user's shape still selected").toEqual([0]);
  });

  it("ignores a selection too small to be a real placeholder", async () => {
    host.selectionBounds = { left: 5, top: 5, width: 20, height: 20 };
    $("insert").click();
    await settle();
    // Below the 40pt threshold → falls back to the tiled offset, not the bounds.
    expect(host.calls.insertScene[0].left).toBeGreaterThanOrEqual(60);
  });
});

describe("Insert updates in place after loading a chart", () => {
  beforeEach(bootHostPane);

  it("routes Insert to an update once a chart is loaded from the selection", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();

    // Now the primary Insert edits in place rather than dropping a new chart.
    $("insert").click();
    await settle();
    expect(host.calls.updateChart).toHaveLength(1);
    expect(host.calls.updateChart[0].target).toMatchObject({ shapeId: "shape-9" });
    expect(host.calls.insertScene).toHaveLength(0);
  });

  /**
   * `updateChartInSlide` returns null when the target slide or shape is gone —
   * deliberately, since "a chart whose slide is gone is not an error, it is
   * nothing to do". Nothing consumed the null: the guard saw an unchanged note
   * and printed "Done." in green, and the stale target kept the button reading
   * "Update chart", so every later push no-opped just as silently.
   */
  it("says so, instead of Done., when the target chart is gone", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    expect($("insert").textContent).toBe("Update chart");

    // The mocked update returns undefined — the "nothing to do" answer.
    $("insert").click();
    await settle();
    const note = $("host-note");
    expect(note.textContent).not.toBe("Done.");
    expect(note.className).toContain("status-err");
    // And the pane falls back to inserting, rather than pushing into thin air.
    expect($("insert").textContent).toBe("Insert into slide");
    $("insert").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
  });

  it("Insert-new always drops a fresh chart even with a chart loaded", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([4, 5]),
      target: { slideId: "s1", shapeId: "shape-3", left: 0, top: 0 },
    };
    $("load-selection").click();
    await settle();

    $("insert-new").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(1);
    expect(host.calls.updateChart).toHaveLength(0);
  });
});

describe("Load selection", () => {
  beforeEach(bootHostPane);

  it("loads an SSF chart and reveals the in-place edit affordance", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([7, 8, 9]),
      target: { slideId: "s1", shapeId: "shape-1", left: 0, top: 0 },
    };
    $("load-selection").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("loaded");
    // The "edit selected chart" banner is stale once loaded, so it hides.
    expect($("selection-banner").style.display).toBe("none");
  });

  it("reports a non-SSF-chart selection without touching the deck", async () => {
    host.loadSelectionResult = null;
    $("load-selection").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("not an ssf chart");
    expect(host.calls.updateChart).toHaveLength(0);
  });
});

describe("Same scale", () => {
  beforeEach(bootHostPane);

  it("pins every value-axis chart in the deck to the union extent", async () => {
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
    ];
    $("same-scale").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(1);
    const batch = host.calls.updateCharts[0];
    expect(batch).toHaveLength(2);
    // Both charts are re-tagged with the SAME max — the union of the two extents.
    const maxima = batch.map((b) => (JSON.parse(b.opts!.tagData!) as ChartConfig).scale?.max);
    expect(new Set(maxima).size).toBe(1);
    expect(maxima[0]).toBe(90);
  });

  it("refuses to apply with fewer than two value-axis charts", async () => {
    host.deckCharts = [{ configJson: chartJson([1, 2, 3]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } }];
    $("same-scale").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(0);
    expect($("host-note").textContent?.toLowerCase()).toContain("two");
  });

  /**
   * "The deck now shares one scale" is the whole feature, and it was being
   * printed in green over a scan the code already knew was short.
   *
   * `listChartsInDeck` computes `unread` carefully — it has a comment saying so
   * — and then sent it to a trace and handed back a bare array. Tracing is off
   * by default, so in ordinary use that distinction reached nobody, and the max
   * was taken over whatever slides happened to answer. A real host produced
   * `unread=8 slides=8` and, in the same run, `unread=7 slides=8`.
   */
  it("will not rescale a deck it could not fully read", async () => {
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
    ];
    host.slideCount = 8;
    host.deckScanUnread = 6;
    $("same-scale").click();
    await settle();
    expect(host.calls.updateCharts, "rescaled a deck it could not see").toHaveLength(0);
    const said = $("host-note").textContent?.toLowerCase() ?? "";
    expect(said, "did not say the scan was short").toContain("whole deck");
    expect($("host-note").className).toContain("status-err");
  });

  /**
   * The count in the success note has to be what the host TOOK.
   *
   * `updateChartsInSlides` drops a chart whose slide or shape the host declines
   * to resolve, and it drops it silently — those are early filters, not throws,
   * so `onFailed` never fires and the pane's `stalled` list stays empty. The
   * note said `parsed.length`, the charts requested. A real host applied the
   * shared scale to one chart in six and the self-test caught it only by
   * re-reading the deck afterwards (`1 of 6 charts carry the shared scale`);
   * this path does no such re-read and would have reported six of six.
   */
  it("reports the charts the host took, not the charts it was asked for", async () => {
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
      { configJson: chartJson([1, 40]), target: { slideId: "s3", shapeId: "c", left: 0, top: 0 } },
    ];
    host.updateChartsDrops = 2; // two charts the host would not resolve, silently
    $("same-scale").click();
    await settle();
    const said = $("host-note").textContent ?? "";
    expect(said, "claimed the whole deck").not.toContain("applied to 3");
    expect(said).toContain("1 of 3");
    expect($("host-note").className, "reported a partial rescale in green").toContain("status-err");
  });

  it("scopes to the selection and guides the user when too few are selected", async () => {
    host.selectionCharts = [
      { configJson: chartJson([1, 2]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
    ];
    $("same-scale-sel").click();
    await settle();
    expect(host.calls.updateCharts).toHaveLength(0);
    // The selection-scoped guard names Ctrl-click, not the deck message.
    expect($("host-note").textContent?.toLowerCase()).toContain("ctrl-click");
  });

  it("sweeps the debris every stalled redraw left behind, not just the first", async () => {
    // Same Scale deletes each chart's old shapes and redraws them. A stall
    // leaves whatever committed on the slide, and the single-chart path has
    // swept that since updateChartResilient learned to — this path never did.
    // It reported the charts "now empty" while half a chart sat on each one.
    //
    // TWO stalls on purpose. The renderer used to annotate only the FIRST
    // failure with its wreckage and hand `onFailed` the raw error regardless,
    // so a one-stall test would pass against a caller that swept nothing and a
    // renderer that reported nothing. The second chart is what proves both
    // halves are fixed.
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
      { configJson: chartJson([1, 40]), target: { slideId: "s3", shapeId: "c", left: 0, top: 0 } },
    ];
    host.updateChartsStalls.set(0, ["stray-a1", "stray-a2"]);
    host.updateChartsStalls.set(2, ["stray-c1"]);
    $("same-scale").click();
    await settle();
    // Both wrecks cleared, each against its OWN slide.
    expect(host.calls.swept).toEqual([
      { slideId: "s1", ids: ["stray-a1", "stray-a2"] },
      { slideId: "s3", ids: ["stray-c1"] },
    ]);
    // And the user is still told which charts went blank.
    expect($("host-note").textContent?.toLowerCase()).toContain("would not redraw");
  });

  it("sweeps nothing when every chart redraws", async () => {
    // The other half of the contract: a clean run must not go poking at the
    // slide. A sweep here would be deleting shapes that are the new chart.
    host.deckCharts = [
      { configJson: chartJson([10, 20, 30]), target: { slideId: "s1", shapeId: "a", left: 0, top: 0 } },
      { configJson: chartJson([5, 90]), target: { slideId: "s2", shapeId: "b", left: 0, top: 0 } },
    ];
    $("same-scale").click();
    await settle();
    expect(host.calls.swept).toHaveLength(0);
    expect($("host-note").textContent?.toLowerCase()).toContain("same scale applied");
  });
});

describe("Elements and batch insert", () => {
  beforeEach(bootHostPane);

  // The five element buttons (harvey balls, checkboxes, process flow, KPI row,
  // table) all drop a compact scene at the same fixed offset through the guard.
  for (const id of ["harvey-insert", "check-insert", "flow-insert", "kpi-insert", "table-insert"]) {
    it(`${id} inserts a compact element scene at the element offset`, async () => {
      $(id).click();
      await settle();
      expect(host.calls.insertScene).toHaveLength(1);
      expect(host.calls.insertScene[0].left).toBe(120);
      expect(host.calls.insertScene[0].top).toBe(160);
    });
  }

  it("batch-inserts every config in the JSON box onto the current slide", async () => {
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify([
      { kind: "pie", data: { categories: ["A", "B"], series: [{ name: "S", values: [1, 1] }] } },
      { kind: "stacked", data: { categories: ["A"], series: [{ name: "S", values: [3] }] } },
    ]);
    $("json-insert-batch").click();
    await settle();
    expect(host.calls.insertScene).toHaveLength(2);
    expect($("host-note").textContent?.toLowerCase()).toContain("2 chart");
  });
});

describe("guard — busy lockout and error surfacing", () => {
  beforeEach(bootHostPane);

  it("disables the button while the action runs and re-enables it after", async () => {
    // Hold the insert open so the mid-flight state is observable.
    let release!: () => void;
    host.gate = new Promise<void>((r) => (release = r));
    const insertBtn = $<HTMLButtonElement>("insert");

    insertBtn.click();
    await settle();
    expect(insertBtn.disabled).toBe(true); // locked out mid-action

    release();
    await settle();
    expect(insertBtn.disabled).toBe(false); // restored once the action settles
    expect(host.calls.insertScene).toHaveLength(1);
  });

  /**
   * "Done." in green over a chart that is no longer an SSF chart.
   *
   * The renderer has always known — `groupAndTagAll` returns `tagged` — and the
   * insert path discarded its whole return value, so nothing reached the user.
   * The chart is on the slide, clicking it says "the selection is not an
   * SSF chart", and reopening the deck loses the settings permanently. A real
   * host produced it four times in one run.
   */
  it("says when an inserted chart carries no config, instead of Done.", async () => {
    host.insertResult = { slideId: "s1", shapeId: "grp-1", left: 40, top: 50, lost: "no-config" };
    $("insert").click();
    await settle();
    const said = $("host-note").textContent ?? "";
    expect(said, "reported an unusable chart as Done.").not.toBe("Done.");
    expect(said.toLowerCase()).toContain("settings");
    expect($("host-note").className).toContain("status-err");
  });

  /**
   * An update that redrew the chart and lost track of it must not keep the
   * target — `shapeId` names the shape that same update deleted, so the next
   * push resolves a dead id (or a group member, and draws a second chart).
   */
  it("drops the edit target when an update loses track of the chart", async () => {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "grp-1", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    host.updateResult = { slideId: "s1", shapeId: "grp-1", left: 10, top: 20, lost: "unknown-shape" };
    $("insert").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase(), "kept quiet about losing the chart").toContain("lost track");
    // The button is back to Insert, so the next press cannot push to a dead id.
    expect($("insert").textContent?.toLowerCase()).not.toContain("update");
  });

  it("surfaces a host failure as a Failed note instead of throwing", async () => {
    host.failInsertOnce = true;
    $("insert").click();
    await settle();
    expect($("host-note").textContent?.toLowerCase()).toContain("failed");
    expect(host.calls.insertScene).toHaveLength(0);
  });

  /**
   * The progress bar exists to say "the host is still working", so it has to
   * stop when the host stops. An insert's last phase note is "Working… done" —
   * still busy — and `guard` used to decide whether to print "Done." by asking
   * whether the note text had changed since it posted "Working…". It had, so
   * "Done." was skipped, nothing ever posted a settled note, and the bar kept
   * its `indeterminate` class: a finished insert under an animation that slid
   * forever. Asserting the bar, not just the words, is the point — the note
   * could read "done" while the strip below it still claimed to be busy.
   */
  it("stops the progress bar when the insert finishes, not just the phase notes", async () => {
    let release!: () => void;
    host.gate = new Promise<void>((r) => (release = r));
    $("insert").click();
    await settle();
    // Mid-flight the bar must actually be running, or the assertion after the
    // release would pass against a bar that never showed at all.
    expect($("status-bar").hasAttribute("hidden")).toBe(false);
    expect($("status-bar").classList.contains("indeterminate")).toBe(true);

    release();
    await settle();
    expect($("host-note").textContent).toBe("Done.");
    expect($("host-note").className).toContain("status-ok");
    expect($("status-bar").hasAttribute("hidden")).toBe(true);
    expect($("status-bar").classList.contains("indeterminate")).toBe(false);
  });

  it("closes out an element insert too — every phase-reporting action settles", async () => {
    $("harvey-insert").click();
    await settle();
    expect($("host-note").textContent).toBe("Done.");
    expect($("status-bar").classList.contains("indeterminate")).toBe(false);
  });

  /**
   * The flip side: an action that DID report an end state keeps it. "Done." is
   * the fallback for silence, not an overwrite — a settlement counted per
   * action is what keeps those two apart.
   */
  it("leaves an action's own closing note alone instead of overwriting it with Done.", async () => {
    host.loadSelectionResult = null;
    $("load-selection").click();
    await settle();
    expect($("host-note").textContent).not.toBe("Done.");
    expect($("host-note").textContent?.toLowerCase()).toContain("not an ssf chart");
    expect($("status-bar").classList.contains("indeterminate")).toBe(false);
  });
});

describe("demo-insert results pages", () => {
  beforeEach(bootHostPane);

  it("attempts each results page in its own insertDemoDeck call — a failing page does not drop the rest", async () => {
    // Presentation_3.pptx: one failing results page threw insertDemoDeck's
    // "everything failed" guard and NO pages landed. Fix: pane loops over
    // pages, per-page try/catch, so page 2 still gets attempted after page 1
    // throws.
    //
    // With the fake mock returning "failed" for every item and demoItems()'s
    // 32-ish failure-count on the full deck, buildResultsScenes emits multiple
    // pages (≥2). Fail the FIRST results page's insertDemoDeck call; the
    // pane must still call insertDemoDeck at least once more for page 2.
    host.demoDeckPageFailures.add(2); // main deck = call #1, results page 1 = call #2
    $("demo-insert").click();
    await settle();
    // At least 3 calls: main deck + results page 1 (thrown) + results page 2.
    // If the pane batched all results into one call (the old behavior), only
    // 2 calls would happen and the failure at #2 would drop the rest.
    expect(host.demoRuns).toBeGreaterThanOrEqual(3);
    // Each results-page call carries exactly ONE item (per-page loop), NOT
    // the whole batch — the load-bearing property of the fix.
    const resultsCalls = host.demoDeckCalls.slice(1); // drop the main deck call
    for (const call of resultsCalls) expect(call).toHaveLength(1);
    // The pane's note reflects that some pages landed and some did not.
    const noteText = $("host-note").textContent ?? "";
    expect(noteText).toMatch(/results page/i);
  });

  it("reconciles each results page, so a half-landed one cannot be orphaned", async () => {
    // The results insert runs at the worst possible moment — right after a run
    // that has just finished exhausting the host — and used to get none of the
    // run's own protections. A page whose add landed but whose shapes did not
    // left a stamped, untagged slide at the END of the deck that nothing ever
    // cleaned: the main run's repair had already finished, and its range
    // stopped short of that slide. A real 38-item web run ended exactly so.
    $("demo-insert").click();
    await settle();
    const resultsOpts = host.demoDeckOpts.slice(1); // drop the main deck call
    expect(resultsOpts.length).toBeGreaterThan(0);
    for (const o of resultsOpts) expect(o?.reconcile).toBe(true);
  });

  it("writes the run log after the results pages, not before them", async () => {
    // Taken before, the log's trace ended at the repair read — so when a
    // results page then failed, the file said "(results slide not added)" with
    // nothing in it about why. The log has to outlast the run it describes.
    let release!: () => void;
    host.demoDeckGate = new Promise<void>((r) => (release = r));
    host.demoDeckGateOn = 2; // the first results page
    $("demo-insert").click();
    await settle();
    // Mid-run, held inside the results insert: no log yet.
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(true);
    release();
    host.demoDeckGate = null;
    await settle();
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(false);
  });

  it("all results pages land ⇒ no 'results slide not added' warning", async () => {
    $("demo-insert").click();
    await settle();
    const noteText = $("host-note").textContent ?? "";
    expect(noteText).not.toMatch(/results slide not added/i);
    // At least the main deck + at least one results page.
    expect(host.demoRuns).toBeGreaterThanOrEqual(2);
  });
});

describe("reporting what a repair actually removed", () => {
  beforeEach(bootHostPane);

  it("counts duplicates and empty strays apart", async () => {
    // A real run said "1 duplicate slide removed · 8 orphan slides" and then,
    // one sentence later, "Repaired 9 duplicate slide(s)". Both numbers came
    // from the same pass; only the second was wrong.
    host.demoDeckStatusOverride = "rendered";
    host.demoReconcile = {
      snapshots: [],
      plan: {
        actions: [
          { kind: "delete", index: 5, slot: 3, reason: "dup" },
          { kind: "delete", index: 4, slot: null, reason: "empty" },
          { kind: "delete", index: 2, slot: null, reason: "empty" },
        ],
        verdicts: [],
        orphans: [],
        summary: {
          items: 1,
          rendered: 1,
          partial: 0,
          lost: 0,
          skipped: 0,
          wreckage: 0,
          empty: 0,
          duplicates: 1,
          falseStamps: 0,
          untagged: 0,
          orphans: 2,
        },
      },
      applied: { unstamped: 0, regrouped: 0, deleted: 3 },
      refused: 0,
    };
    $("demo-insert").click();
    await settle();
    const text = $("host-note").textContent ?? "";
    expect(text).toMatch(/removed 3 slide\(s\) \(1 duplicate, 2 empty\)/);
    expect(text).not.toMatch(/3 duplicate slide/);
  });
});

describe("updating a chart the live canvas will not redraw", () => {
  beforeEach(bootHostPane);

  /** Load a chart so Insert becomes Update, then push an edit. */
  async function loadThenUpdate() {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    $("insert").click();
    await settle();
  }

  it("does not start a second write to the same chart while one is in flight", async () => {
    // The auto-update timer calls doInsert DIRECTLY, so it never sees the
    // disabled buttons a click would. One resilient update can legitimately
    // take tens of seconds — a 45s stall, then a slide swap, then a bounded
    // raster — and a user editing through that could start a second update
    // against the SAME stale target. Whichever finished last won; if that was
    // the one started from the already-superseded target, the pane reported
    // "that chart is no longer on the slide" about a chart just written fine.
    vi.useFakeTimers();
    try {
      host.loadSelectionResult = {
        configJson: chartJson([1, 2, 3]),
        target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
      };
      host.updateResult = { slideId: "s1", shapeId: "shape-10", left: 10, top: 20 };
      $("load-selection").click();
      await vi.advanceTimersByTimeAsync(10);
      ($("auto-update") as HTMLInputElement).checked = true;

      // Hold the first update open, then let the debounce fire underneath it.
      let release!: () => void;
      host.updateGate = new Promise<void>((r) => (release = r));
      $("insert").click();
      await vi.advanceTimersByTimeAsync(10);
      const duringFirst = host.calls.updateChart.length;
      // Two edits land while the first write is still open.
      ($("chart-title") as HTMLInputElement).value = "edited";
      $("chart-title").dispatchEvent(new Event("input", { bubbles: true }));
      await vi.advanceTimersByTimeAsync(3000);
      // Nothing new was issued — the timer re-armed instead of racing.
      expect(host.calls.updateChart).toHaveLength(duringFirst);
      release();
      host.updateGate = null;
      await vi.advanceTimersByTimeAsync(3000);
      // …and once the first finished, the queued edit did go out.
      expect(host.calls.updateChart.length).toBeGreaterThan(duringFirst);
    } finally {
      host.updateGate = null;
      vi.useRealTimers();
    }
  });

  it("redraws with the slide deselected, at the off-screen batch size", async () => {
    // The whole point: a redraw is the add-in's worst case on the web, and it
    // is only bad because the slide is on screen. Looking away is free.
    await loadThenUpdate();
    expect(host.calls.deselected).toEqual([["s1"]]);
  });

  it("swaps the slide when the redraw stalls and the slide holds only the chart", async () => {
    host.updateChartThrows = true;
    host.canInsertFile = true;
    host.slideHoldsOnlyChart = true;
    host.swapOutcome = "swapped";
    await loadThenUpdate();
    expect(host.calls.insertFile).toHaveLength(0); // a slide swap, not a deck insert
    expect($("host-note").textContent).toMatch(/rebuilt that slide/i);
    // And it says what the rebuild COST. A swap replaces the slide, so the
    // chart comes back and anything on it that is not a shape does not:
    // speaker notes, transition, animations. The guard in front of the swap
    // can only see shapes — Office.js exposes no way to read notes at all
    // (office-js#3269) — so the add-in cannot ask first and cannot avoid it.
    // Telling the user afterwards is the difference between a loss they can
    // undo and one they discover in front of an audience.
    expect($("host-note").textContent, "a swap discarded the slide's notes without saying so").toMatch(
      /speaker notes|replaced/i,
    );
  });

  it("stops, and says so, when the swap left the original slide behind", async () => {
    // The swap inserts the new slide and THEN deletes the old one. When only
    // the delete fails, the chart is on two slides — and answering the caller
    // the same "false" as a swap that did nothing sent it on to the picture
    // layer, which rasterized the chart onto the surviving original. One
    // stall, the chart three times over, and a message saying it went fine.
    host.updateChartThrows = true;
    host.canInsertFile = true;
    host.slideHoldsOnlyChart = true;
    host.swapOutcome = "duplicated";
    await loadThenUpdate();
    // No picture update was attempted on the original.
    expect(host.calls.updateChart.filter((c) => c.opts.pictureBase64)).toHaveLength(0);
    expect($("host-note").textContent).toMatch(/two slides/i);
  });

  it("gives up on a canvas that never answers instead of hanging the pane", async () => {
    // jsdom never fires `img.onload`, which is exactly the failure being
    // guarded: a wedged image decode. Unbounded, the await never settles —
    // the button stays disabled, no error, no timeout, nothing to do but
    // reload the pane. Same Scale awaits one of these PER CHART, so one
    // wedged decode froze the whole deck-wide operation.
    vi.useFakeTimers();
    try {
      host.autoPicture = true; // the too-dense branch, which rasterizes
      $("insert").click();
      await vi.advanceTimersByTimeAsync(30_000);
      // The action finished: the chart went in as shapes, and the pane says why.
      expect(($("insert") as HTMLButtonElement).disabled).toBe(false);
      expect(host.calls.insertScene).toHaveLength(1);
      expect($("host-note").textContent).toMatch(/could not be turned into a picture/i);
    } finally {
      vi.useRealTimers();
    }
  });

  it("will not swap a slide that holds anything besides the chart", async () => {
    // The replacement is a NEW slide: notes, transitions and any other shape
    // on the old one do not come with it. So it is only ever offered where
    // there is demonstrably nothing else to lose.
    host.updateChartThrows = true;
    host.canInsertFile = true;
    host.swapOutcome = "swapped"; // the swap WOULD work — the guard is the only thing stopping it
    host.slideHoldsOnlyChart = false;
    await loadThenUpdate();
    expect($("host-note").textContent).not.toMatch(/rebuilt that slide/i);
  });

  /**
   * The floor of the ladder, on the failure it exists for.
   *
   * Layer 1 deletes the old chart, COMMITS that, and only then draws — so a
   * stall leaves half a chart on the slide and no shape at the target's id.
   * Every layer below assumed the chart was still there: layer 2 checks the
   * slide holds nothing but the chart and saw the litter, layer 3 re-resolved
   * the id layer 1 had just destroyed, got nothing back, and the pane announced
   * "that chart is no longer on the slide" over a half-drawn one. Three
   * fallbacks, all disabled by the damage they were meant to repair.
   */
  it("sweeps the wreckage and draws the chart back, rather than calling it gone", async () => {
    const raster = stubRaster();
    try {
      host.updateChartThrows = true; // layer 1 stalls — after the delete committed
      host.slideHoldsOnlyChart = false; // layer 2 out of the picture, so this is the floor
      host.insertResult = { slideId: "s1", shapeId: "grp-new", left: 40, top: 50 };
      await loadThenUpdate();

      // The half-drawn chart went first — otherwise the replacement lands on
      // top of it and the user keeps both.
      expect(host.calls.swept).toEqual([{ slideId: "s1", ids: ["stray-1", "stray-2"] }]);

      // Then the chart was drawn back: onto the slide it came from, at the
      // position it held, as one picture no live canvas can stall on.
      const drawn = host.calls.insertScene.at(-1);
      expect(drawn).toMatchObject({ slideId: "s1", left: 40, top: 50 });
      expect(drawn?.pictureBase64).toBeTruthy();

      // And NOT as a picture update against the shape layer 1 deleted — that
      // call can only ever resolve nothing.
      expect(host.calls.updateChart.filter((c) => c.opts.pictureBase64)).toHaveLength(0);
      expect($("host-note").textContent).not.toMatch(/no longer on the slide/i);
    } finally {
      raster.restore();
    }
  });

  it("keeps the chart editable after that recovery, so the next edit still lands", async () => {
    const raster = stubRaster();
    try {
      host.updateChartThrows = true;
      host.slideHoldsOnlyChart = false;
      host.insertResult = { slideId: "s1", shapeId: "grp-new", left: 40, top: 50 };
      await loadThenUpdate();
      // The button still reads Update, against the shape the recovery created —
      // not the dead one, and not "Insert into slide".
      expect(($("insert") as HTMLButtonElement).textContent).toMatch(/update/i);
      host.updateChartThrows = false;
      host.updateResult = { slideId: "s1", shapeId: "grp-newer", left: 40, top: 50 };
      $("insert").click();
      await settle();
      expect(host.calls.updateChart.at(-1)?.target).toMatchObject({ shapeId: "grp-new" });
    } finally {
      raster.restore();
    }
  });

  it("still says 'gone' when the chart really is gone — nothing was destroyed to recover", async () => {
    // The other side of the same coin: an update that found nothing to replace
    // destroyed nothing, so there is no wreckage, no sweep, and no chart to
    // draw back. Telling this user to insert again is right; doing it for the
    // case above would have given them the chart twice.
    host.updateResult = undefined; // resolved no shape — the user deleted it
    await loadThenUpdate();
    expect(host.calls.swept).toHaveLength(0);
    expect(host.calls.insertScene).toHaveLength(0);
    expect($("host-note").textContent).toMatch(/no longer on the slide/i);
  });
});

describe("demo-insert one-shot deck insert", () => {
  beforeEach(bootHostPane);

  it("hands the host a generated file and never draws the deck shape by shape", async () => {
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.calls.insertFile[0].b64).toMatch(/^UEsDBBQA/);
    // The whole point: not one slide is drawn through the shape renderer.
    expect(host.demoRuns).toBe(0);
    expect($("host-note").textContent).toMatch(/one file/i);
  });

  it("says WHY when it could not verify, instead of quietly reporting a raw count", async () => {
    // A real run reported "Inserted 12 of 12 slides as one file" — the
    // fallback wording — because the verification pass returned null and the
    // pane shrugged. Three different failures produced that same null and the
    // message named none of them.
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    // The fake reads no slot tags, so the pass cannot identify a single slide.
    expect($("host-note").textContent).toMatch(/not verified:/i);
    expect($("host-note").textContent).toMatch(/slot tag|no slides/i);
  });

  it("falls back to shapes when the host has no insertSlidesFromBase64", async () => {
    host.canInsertFile = false;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("falls back when the file could not be built at all", async () => {
    host.canInsertFile = true;
    host.buildFileError = new Error("pptxgenjs blew up");
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("falls back when the insert threw and nothing landed", async () => {
    host.canInsertFile = true;
    host.insertFileError = new Error("host refused the deck");
    host.slideCount = 1; // unchanged before and after — nothing landed
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("does NOT fall back after a partial insert — that would draw the deck twice", async () => {
    // The failure this exists to prevent: some slides landed, the pane decides
    // the fast path "failed", and the shape renderer then appends the whole
    // deck a second time on top of them.
    host.canInsertFile = true;
    host.insertFileLands = 3;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBe(0);
    expect($("host-note").textContent).toMatch(/host took 3 of/i);
  });

  /**
   * The run token has to reach the FILE, not just the run.
   *
   * It is the join key between the downloaded log and the .pptx the run
   * produced, and every diagnosis in this project's history is that join. Done
   * by hand it started with a guess at which slides were the run's, because a
   * deck holds whatever earlier runs left in it — one real file carried 30
   * slides from the run under investigation and one from another. A log
   * missing the token sends `npm run triage` straight back to guessing.
   */
  it("puts the run's identity in the log the fast path writes", async () => {
    const dl = captureDownloads();
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    $("demo-log").click();
    const log = await dl.lastJson();
    expect(log.runs).toHaveLength(1);
    expect(log.runs[0].run).toBe("run-under-test");
    // …and which items were ever meant to carry a config. Without this the
    // title and contents pages — `PowerChart` objects with no config by
    // design — read as charts that lost their tag, and a clean run triages as
    // a broken one. Seven false alarms on the first real deck this was run on.
    const items = log.runs[0].items;
    const chartFlags: boolean[] = items.map((i: { chart: boolean }) => i.chart);
    expect(chartFlags).toContain(true);
    expect(chartFlags).toContain(false);
    expect(items.find((i: { title: string }) => i.title === "Title").chart).toBe(false);
    dl.restore();
  });

  it("puts the run's identity in the log the shape path writes", async () => {
    // A different token from the fast path's, and taken from the renderer's
    // own report rather than minted again by the pane — the pane minting its
    // own would produce a log that joins to nothing in the deck.
    const dl = captureDownloads();
    host.canInsertFile = false;
    $("demo-insert").click();
    await settle();
    $("demo-log").click();
    expect((await dl.lastJson()).runs[0].run).toBe("fake-run-1");
    dl.restore();
  });

  it("leaves a downloadable run log behind — including when the run went badly", async () => {
    // The whole point of the log is diagnosing real host failures after the
    // fact, and the fast path is what a real run TAKES: the checkbox ships
    // checked and every current host advertises insertSlidesFromBase64. Yet
    // this path returned before `lastRunLog` was ever set, so the button that
    // downloads it stayed disabled — success or failure. The failing run is
    // exactly the one worth having a file for.
    host.canInsertFile = true;
    host.insertFileLands = 3; // a partial insert: the case most worth logging
    $("demo-insert").click();
    await settle();
    expect($("host-note").textContent).toMatch(/host took 3 of/i);
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(false);
  });

  it("does not offer an older run's log after a run that produced none", async () => {
    // The button used to stay enabled from a previous run, handing out a file
    // about a completely different insert with nothing on screen to say so.
    host.canInsertFile = false;
    $("demo-insert").click();
    await settle();
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(false); // shape path logs
    // Now a run that produces no log of its own: the deck build fails, and the
    // shape path it falls back to throws on its very first call.
    host.canInsertFile = true;
    host.buildFileError = new Error("pptxgenjs blew up");
    host.demoRuns = 0; // the failure indices below count from this run
    host.demoDeckPageFailures = new Set([1]);
    $("demo-insert").click();
    await settle();
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(true);
  });

  it("does not redraw the whole deck when it could not measure what landed", async () => {
    // The insert throws, and the rescue count read fails WITH it — the same
    // wedged host, milliseconds apart. Treating that as "nothing landed" sent
    // the shape renderer off to draw all 38 slides a second time, on top of
    // however many did arrive. When it cannot tell, it must not guess.
    host.canInsertFile = true;
    host.insertFileError = new Error("host refused the deck");
    host.slideCountThrowsAfter = 1; // the opening read answers; the rescue does not
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBe(0); // the deck was NOT drawn again
    expect($("host-note").textContent).toMatch(/twice/i);
  });

  it("appends the generated deck instead of letting the host front it", async () => {
    // The first real run put 37 generated slides AHEAD of the user's own title
    // slide, because insertSlidesFromBase64 inserts at the front unless it is
    // given a target. The pane must ask for the tail.
    host.canInsertFile = true;
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.calls.insertFile[0].expected).toBeGreaterThan(0);
  });

  it("respects the fast-path opt-out", async () => {
    host.canInsertFile = true;
    ($("demo-path") as HTMLSelectElement).value = "shapes";
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
  });

  it("takes both paths on one click, and logs them as two runs", async () => {
    // Every renderer change touches both paths, and they fail in completely
    // different ways — so testing them meant two runs an hour apart with a
    // deploy in between. One click, one deck, one log, same session.
    const dl = captureDownloads();
    host.canInsertFile = true;
    ($("demo-path") as HTMLSelectElement).value = "both";
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns).toBeGreaterThanOrEqual(1);
    $("demo-log").click();
    const log = await dl.lastJson();
    expect(log.runs.map((r: { path: string }) => r.path)).toEqual(["file", "shapes"]);
    // Separate identities, or the deck they share cannot be read apart.
    expect(log.runs[0].run).not.toBe(log.runs[1].run);
    dl.restore();
  });

  /**
   * One click, one file.
   *
   * A round used to be three clicks producing three downloads, uploaded
   * separately and joined at the other end — and most probe runs establish
   * nothing, so a good share of that traffic existed to learn that the answers
   * had not changed.
   *
   * The demo deck is deliberately NOT chained in: its two halves have to run on
   * different decks, and a button cannot open a fresh one.
   */
  it("says how loaded the deck was before it starts", async () => {
    // A round that dies leaves only its steps, and on 2026-08-09 one died
    // sixteen seconds in — two 8-second stalls and then the tab — with nothing
    // in the file to say whether PowerPoint had been fresh or already carrying
    // a self-test's worth of shapes. This project has documented since
    // 2026-08-06 that heavy work on an already-large deck is what kills the
    // tab, and `docs/PUBLISHING.md` splits the demo halves across two decks for
    // exactly that reason. Both readings fitted the crash file and neither
    // could be checked.
    const { setTracing, traceLog } = await import("../src/core/trace");
    host.deckSlideIds = ["a", "b", "c"];
    setTracing(true);
    try {
      $("demo-round").click();
      await settle();
      const said = traceLog().entries.filter((e) => e.message === "round starting");
      expect(said, "a round that dies cannot say what deck it died on").toHaveLength(1);
      // The number, not merely the line. A line carrying nothing would satisfy
      // a check for its own existence and answer the question no better.
      expect(said[0].data?.deckSlides).toBe(3);
    } finally {
      setTracing(false);
      host.deckSlideIds = undefined;
    }
  });

  it("puts the environment INTO the round file, not into a window it slices off", async () => {
    // ZERO OF 69 ARCHIVED ROUNDS carry a `host`/`environment` line, or a
    // `slide size` line, on an archive where every single build contains the
    // code that emits them. `traceEnvironment` was called at pane-wiring time,
    // and the round's `traceLog(traceFrom)` slices off everything traced before
    // its mark — so the host, the platform, the Office version and the slide
    // size have never once reached a round file.
    //
    // The pre-mark window is not small: round 092's first entry sits at
    // ms 48245 with `dropped: 0`, so 48 seconds of trace was cut away.
    //
    // Asserted against the DOWNLOADED BUNDLE rather than `traceLog()`, because
    // the whole defect is the difference between the two. A test reading the
    // live buffer would have passed for all 69 rounds.
    const dl = captureDownloads();
    const { setTracing, trace } = await import("../src/core/trace");
    traceRef.fn = trace;
    setTracing(true);
    try {
      $("demo-round").click();
      await settle();
      const bundle = (await dl.lastJson()) as { trace?: { entries?: { scope?: string; message?: string }[] } };
      const entries = bundle?.trace?.entries ?? [];
      expect(
        entries.filter((e) => e.scope === "host" && e.message === "environment"),
        "the round file cannot say which host it ran on",
      ).toHaveLength(1);
    } finally {
      setTracing(false);
    }
  });

  it("runs the probe and the self-test on one click, and saves them as one file", async () => {
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    expect(bundle.hostAnswers?.answers?.length, "the round's file carries no probe answers").toBeGreaterThan(0);
    expect(bundle.selftest?.length, "the round's file carries no self-test verdicts").toBeGreaterThan(0);
    // One file, not two. The whole point is that there is one thing to send.
    expect(dl.count()).toBe(1);
    dl.restore();
  });

  /**
   * The two uploads that used to be a person's job.
   *
   * Every diagnosis in this project's history has taken three things — the run
   * log, the deck, and a screenshot — and the owner has been producing all
   * three by hand, once per round. Two of them the add-in can read for itself.
   */
  it("carries what landed on the slides, and pictures of the slides it added", async () => {
    host.deckSlideIds = ["s1"];
    host.deckInventory = [
      { slideId: "s1", index: 0, shapes: [{ id: "sh1", name: "PowerChart" }] },
      { slideId: "s2", index: 1, shapes: [{ id: "sh2", name: "bar 1" }] },
    ];
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    // Shapes that are NOT charts are the point. The scan has always had them
    // and always dropped them, and they are what "41 shapes became 79" and
    // "the slide still holds what was there before" are questions about.
    expect(bundle.deck?.inventory?.[1]?.shapes?.[0]?.name).toBe("bar 1");
    dl.restore();
  });

  it("says what it left in the deck, and points at the button that clears it", async () => {
    // A round leaves its slides on purpose — they ARE the evidence, and
    // `docs/REGRESSION.md` is written around a deck someone can open. What it
    // never did was say so. The 2026-08-09 evening round added 43 slides, 36 of
    // which read back empty, reported "Saved as one file", and left the owner
    // to discover a 44-slide deck by opening it. `Clean up the last round` had
    // existed the whole time and nothing named it at the moment it mattered.
    //
    // The same gap cost a round at the other end of this one: the host probe
    // left 21 blank slides and reported nothing, which is why it carries a row
    // in the answer sheet now. The self-test never got the equivalent.
    host.deckSlideIds = ["s1"];
    host.roundAddsSlides = ["added-full", "added-empty"];
    // `s1` is EMPTY on purpose, and it is the whole point of the fixture: it is
    // the user's own blank slide, sitting in the deck before the round started.
    // With it full, dropping the added-slides filter changed no number and the
    // assertion below passed against the bug it names.
    host.deckInventory = [
      { slideId: "s1", index: 0, shapes: [] },
      { slideId: "added-full", index: 1, shapes: [{ id: "sh2", name: "bar 1" }] },
      { slideId: "added-empty", index: 2, shapes: [] },
    ];
    $("demo-round").click();
    await settle();
    const said = $("host-note").textContent ?? "";
    expect(said, "the round never said it left anything behind").toMatch(/left 2 slides in this deck/);
    // ONE of the two, and not the slide that was already there: an empty count
    // that swept in the deck the round landed in would report a user's own
    // blank slides as our litter.
    expect(said, "counted the wrong slides as empty").toMatch(/1 of which read back empty/);
    expect(said, "named no way to clear them").toContain("Clean up the last round");
  });

  it("photographs only the slides the round added, never the whole deck", async () => {
    // A picture of a forty-slide deck is mostly slides nobody touched. The id
    // diff is what makes the pictures worth the bytes — and ids rather than
    // counts, because a deck's own id list is the stronger question about it.
    host.deckSlideIds = ["s1"];
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const before = await dl.lastJson();
    expect(before.deck?.newSlides, "photographed a slide that was there before the round").not.toContain("s1");
    dl.restore();
  });

  it("does not claim it added the whole deck when the before-list could not be read", async () => {
    // THE FALLBACK WAS `idsAfter`, which says every slide the deck holds was
    // added by this round. A round that added four slides to the owner's
    // forty-slide deck reported forty — and `describeLitter` then told them it
    // had left forty behind, some of which "read back empty". Those were their
    // own blank slides. `triage.mjs` builds its added-set from this field too.
    //
    // Unknown is its own answer, the same lesson as reading a host's silence as
    // a "no": `newSlides` stays empty and `beforeUnknown` says why.
    host.deckSlideIds = ["s1", "s2", "s3"];
    host.deckSlideIdsBeforeFails = true;
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    expect(bundle.deck?.newSlides, "claimed the user's own slides as this round's litter").toEqual([]);
    expect(
      bundle.deck?.beforeUnknown,
      "an empty list that cannot say why is indistinguishable from adding nothing",
    ).toBe(true);
    dl.restore();
  });

  /**
   * Putting the deck back.
   *
   * A round leaves slides behind on purpose, and clearing them has been a manual
   * chore once per round — in a deck that also grows and skews the next round's
   * timings. What makes it safe to automate is that it deletes an ID LIST the
   * round watched appear, never a rule for recognising a test slide: a rule can
   * match a slide the owner made, an id list cannot.
   */
  it("cleans up exactly the slides the round added, and nothing that was there before", async () => {
    host.deckSlideIds = ["s1"];
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    dl.restore();
    const added = bundle.deck.newSlides as string[];
    expect(added.length, "the round added no slides, so this proves nothing").toBeGreaterThan(0);
    $("demo-tidy").click();
    await settle();
    expect(host.deletedSlides).toEqual(added);
    expect(host.deletedSlides, "deleted a slide that was in the deck before the round").not.toContain("s1");
  });

  it("stays disabled until a round has named something to clean up", async () => {
    // With no id list there is nothing to delete but a guess, and guessing about
    // deletion in someone's own deck is not a trade this pane makes.
    expect(($("demo-tidy") as HTMLButtonElement).disabled).toBe(true);
    $("demo-tidy").click();
    await settle();
    expect(host.deletedSlides).toEqual([]);
  });

  it("says how many the host refused rather than reporting a clean sweep", async () => {
    // `deleteSlideById` has a whole comment about why a host saying "gone" is
    // not proof. A cleanup that reported success for slides still sitting in the
    // deck would be the same class of claim.
    host.deckSlideIds = ["s1"];
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    dl.restore();
    host.refuseSlideDelete = [(bundle.deck.newSlides as string[])[0]];
    $("demo-tidy").click();
    await settle();
    expect($("host-note").textContent ?? "").toMatch(/Removed \d+ of \d+/);

    // And it STAYS dead. The handler empties `tidyable` and disables itself on
    // purpose — after a partial sweep there is no longer a list it trusts — but
    // `guard()` captured the clicked button before running the handler and its
    // finally brought it back. A second press then printed a green
    // "Cleaned up — 0 slide(s) removed." over the slides the host had just
    // refused, which docs/MANUAL.md and the button's own title both deny.
    expect(($("demo-tidy") as HTMLButtonElement).disabled, "guard resurrected a button meant to stay dead").toBe(true);
    const deletedSoFar = [...host.deletedSlides];
    $("demo-tidy").click();
    await settle();
    expect(host.deletedSlides, "a second press asked the host again").toEqual(deletedSoFar);
    expect($("host-note").textContent ?? "").not.toMatch(/Cleaned up/);
  });

  it("still writes the file when the host will not describe its own deck", async () => {
    // The tail of a round runs on a host that has just been through the
    // self-test, and may well be the reason the round is worth reading. Losing
    // the verdicts to the diagnostic's own evidence-gathering would be the
    // worst trade in the file.
    host.deckScanThrows = true;
    host.canRaster = false;
    const dl = captureDownloads();
    $("demo-round").click();
    await settle();
    const bundle = await dl.lastJson();
    expect(bundle.selftest?.length, "lost the verdicts because the deck would not answer").toBeGreaterThan(0);
    expect(bundle.hostAnswers?.answers?.length).toBeGreaterThan(0);
    dl.restore();
  });

  /**
   * "Both, one after the other" has never survived a deck that was not empty.
   *
   * Four attempts, four crash dialogs, and every one has the same shape: the
   * file half fills the deck, the shape half then draws onto that larger deck,
   * and the host stops answering. The last one waited 45 seconds for a batch of
   * FIVE shapes on 40 slides. `docs/PUBLISHING.md` splits the two into separate
   * tests on separate decks for exactly this reason, and the dropdown quietly
   * puts them back together.
   *
   * So the option degrades instead of obeying: run the half that works, and say
   * what the other one needs.
   */
  it("refuses the shape half on a deck that is already full, and says why", async () => {
    const dl = captureDownloads();
    host.canInsertFile = true;
    host.slideCount = 40;
    ($("demo-path") as HTMLSelectElement).value = "both";
    $("demo-insert").click();
    await settle();
    // The file half still runs — it is the half that works on a big deck, and
    // refusing both would cost the measurement entirely.
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns, "drew the deck shape by shape onto a 40-slide deck").toBe(0);
    $("demo-log").click();
    const log = await dl.lastJson();
    expect(log.runs.map((r: { path: string }) => r.path)).toEqual(["file"]);
    // And the LOG says the difference was deliberate. The on-screen note is
    // overwritten by this run's own summary within seconds; a reader of the log
    // would otherwise see a run asked for both halves that quietly did one, and
    // read it as the shape half having crashed.
    expect(log.refusedShapeHalf?.slides, "the log does not say the shape half was refused").toBe(40);
    expect(log.refusedShapeHalf?.why).toMatch(/fresh deck/i);
    dl.restore();
  });

  it("still takes both halves when the deck is empty enough to be worth it", async () => {
    // The gate is a threshold, not a ban. A fresh deck is exactly where the
    // measurement is wanted, and that case has never been the one that crashes.
    host.canInsertFile = true;
    host.slideCount = 1;
    ($("demo-path") as HTMLSelectElement).value = "both";
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(1);
    expect(host.demoRuns, "the shape half was refused on an empty deck").toBeGreaterThanOrEqual(1);
  });

  it("does not run the shape path twice when the fast path falls back", async () => {
    // "Both" continues to the shape path after a SUCCESSFUL file insert, which
    // is the one case where drawing the deck twice is the intent. A fast path
    // that landed nothing already falls back to shapes on its own — running
    // the fall-back and then the second leg would draw it three times.
    const dl = captureDownloads();
    host.canInsertFile = true;
    host.buildFileError = new Error("pptxgenjs blew up");
    ($("demo-path") as HTMLSelectElement).value = "both";
    $("demo-insert").click();
    await settle();
    expect(host.calls.insertFile).toHaveLength(0);
    // Counted by whole-deck calls, not by `demoRuns` — the shape path also
    // calls insertDemoDeck once per results page at the end of a run.
    expect(host.demoDeckCalls.filter((c) => c.length > 5)).toHaveLength(1);
    $("demo-log").click();
    expect((await dl.lastJson()).runs.map((r: { path: string }) => r.path)).toEqual(["shapes"]);
    dl.restore();
  });

  it("keeps the fast path's log when the shape path then throws", async () => {
    // The failing run is the one worth having a file for. Banking the log at
    // the end of the click would discard the file run along with the throw.
    const dl = captureDownloads();
    host.canInsertFile = true;
    ($("demo-path") as HTMLSelectElement).value = "both";
    host.demoRuns = 0;
    host.demoDeckPageFailures = new Set([1]);
    $("demo-insert").click();
    await settle();
    expect(($("demo-log") as HTMLButtonElement).disabled).toBe(false);
    $("demo-log").click();
    const log = await dl.lastJson();
    expect(log.runs.map((r: { path: string }) => r.path)).toEqual(["file"]);
    dl.restore();
  });
});

describe("watchSelection", () => {
  beforeEach(bootHostPane);

  it("offers the edit banner when an SSF chart is selected", async () => {
    expect(host.selectionListener).toBeTypeOf("function");
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "sel-1", left: 0, top: 0 },
    };
    await host.selectionListener!();
    await settle();
    expect($("selection-banner").style.display).toBe(""); // shown
  });

  it("hides the banner when the selection is not an SSF chart", async () => {
    host.loadSelectionResult = null;
    await host.selectionListener!();
    await settle();
    expect($("selection-banner").style.display).toBe("none");
  });

  /**
   * A real host answered two of these ninety seconds later — the run log reads
   * `gave up waiting what=reading the selected chart afterMs=90000`, twice,
   * during Same Scale. 90s is `READBACK_TIMEOUT_MS`, the budget sized for a
   * twenty-slide repair page; a background listener that decides whether one
   * banner is visible inherited it by passing no budget at all.
   */
  it("gives the host a deadline of its own rather than the readback budget", async () => {
    await host.selectionListener!();
    await settle();
    expect(host.selectionBudgets).toHaveLength(1);
    const budget = host.selectionBudgets[0];
    expect(typeof budget).toBe("number");
    expect(budget!).toBeLessThan(10_000);
  });

  /**
   * The pane moves the selection itself — `showSlide` on every deck pass,
   * `withSlideDeselected` before every off-screen redraw, a slide swap on every
   * resilient update. Each of those fires this handler, so a long run used to
   * queue a selection read per slide behind the work the user actually asked
   * for. The banner is worth exactly one read, and only the last one is current.
   */
  it("asks nothing while an action is running, then catches up once", async () => {
    let release!: () => void;
    host.gate = new Promise<void>((r) => (release = r));
    $("insert").click();
    await settle();

    // The pane's own navigation, three slides' worth.
    await host.selectionListener!();
    await host.selectionListener!();
    await host.selectionListener!();
    await settle();
    expect(host.selectionBudgets, "queued host work while the pane was busy").toHaveLength(0);

    release();
    await settle();
    expect(host.selectionBudgets, "did not catch up after the action ended").toHaveLength(1);
  });

  /** Nothing to catch up on means nothing to ask — an action that never moved
   *  the selection must not provoke a read merely by finishing. */
  it("does not read the selection after an action that never moved it", async () => {
    $("insert").click();
    await settle();
    expect(host.selectionBudgets).toHaveLength(0);
  });

  /**
   * Two clicks in the time one read takes are still one question.
   *
   * Fired without awaiting, which is what PowerPoint does: `addHandlerAsync`
   * takes a callback and never looks at what it returns. Awaiting it here would
   * serialise the two reads by hand and test the harness rather than the pane.
   */
  it("keeps only one selection read outstanding", async () => {
    let release!: () => void;
    host.selectionGate = new Promise<void>((r) => (release = r));
    void host.selectionListener!();
    await settle();
    void host.selectionListener!();
    await settle();
    expect(host.selectionBudgets).toHaveLength(1);
    release();
    await settle();
  });
});

describe("Stop", () => {
  beforeEach(bootHostPane);

  /**
   * Load a chart from the selection, then press Insert — which makes Insert an
   * in-place UPDATE of that chart rather than a fresh insert. The update path
   * is the one that is slow enough to need a stop, and the only one that goes
   * through the gate these tests hold open.
   */
  async function startGatedUpdate(): Promise<() => void> {
    host.loadSelectionResult = {
      configJson: chartJson([1, 2, 3]),
      target: { slideId: "s1", shapeId: "shape-9", left: 10, top: 20 },
    };
    $("load-selection").click();
    await settle();
    let release!: () => void;
    host.updateGate = new Promise<void>((r) => (release = r));
    $("insert").click();
    await settle();
    return () => {
      release();
      host.updateGate = null;
    };
  }

  it("is offered only while an action is in flight", async () => {
    // The button lives in the status strip and must be absent the rest of the
    // time — a permanent Stop on an idle pane is a button that does nothing.
    const stop = $("status-stop") as HTMLButtonElement;
    expect(stop.hidden).toBe(true);
    const release = await startGatedUpdate();
    expect(stop.hidden, "no way out of a long action").toBe(false);
    release();
    await settle();
    expect(stop.hidden, "left offering a stop after the action ended").toBe(true);
  });

  it("says it is stopping before the host has come back", async () => {
    // The batch already handed to PowerPoint still has to return — up to
    // BATCH_TIMEOUT_MS. Without immediate feedback the pane looks like it
    // ignored the click for as long as that takes.
    const stop = $("status-stop") as HTMLButtonElement;
    const release = await startGatedUpdate();
    stop.click();
    await settle();
    expect(stop.disabled).toBe(true);
    expect(stop.textContent?.toLowerCase()).toContain("stopping");
    release();
    await settle();
  });

  it("reports a stop as a stop, not as a failure", async () => {
    // "Failed: Stopped." reads like the add-in broke. The user pressed the
    // button; the pane should say what happened, and say the work already
    // drawn was kept.
    const release = await startGatedUpdate();
    ($("status-stop") as HTMLButtonElement).click();
    release();
    await settle();
    const said = $("host-note").textContent?.toLowerCase() ?? "";
    expect(said).toContain("stopped");
    expect(said, "blamed the host for the user's own stop").not.toContain("failed");
  });

  it("still sweeps what the stopped redraw left on the slide", async () => {
    // A stop mid-batch leaves a partial chart exactly as a stall does. Leaving
    // it there would make Stop destructive — the one thing a cancel must not be.
    const release = await startGatedUpdate();
    ($("status-stop") as HTMLButtonElement).click();
    release();
    await settle();
    expect(host.calls.swept).toEqual([{ slideId: "s1", ids: ["stray-1"] }]);
  });

  it("does not run the recovery ladder for a stop", async () => {
    // Every rung below the in-place redraw exists to get the chart drawn some
    // other way — rebuild the slide, rasterize it. That is precisely what a
    // user who just pressed Stop has said they do not want, and the slowest
    // rungs would run AFTER the cancel.
    host.canInsertFile = true;
    host.slideHoldsOnlyChart = true;
    // The swap RUNG FAILS, deliberately: it is what makes the picture rung
    // reachable, and the picture rung is the one that leaves a trace here (an
    // update carrying pictureBase64). Asserting against a swap that succeeds
    // proves nothing — a successful swap records nothing either way.
    host.swapOutcome = "failed";
    const release = await startGatedUpdate();
    ($("status-stop") as HTMLButtonElement).click();
    release();
    await settle();
    // No rasterized redraw was attempted after the cancel.
    expect(host.calls.updateChart.filter((c) => c.opts.pictureBase64)).toHaveLength(0);
    // And the pane did not claim to have rebuilt anything.
    const said = $("host-note").textContent?.toLowerCase() ?? "";
    expect(said).toContain("stopped");
    expect(said).not.toMatch(/rebuil|picture/i);
  });
});

describe("the live step list", () => {
  beforeEach(bootHostPane);

  /**
   * The trace module the PANE is using, not the one this file imported.
   *
   * `bootHostPane` calls `vi.resetModules()` before importing the app, so a
   * static import at the top of this file is a different instance: writing to
   * it would leave the pane's subscriber untouched and every assertion below
   * would pass or fail on an empty box for the wrong reason.
   */
  const paneTrace = async () => {
    const mod = await import("../src/core/trace");
    mod.setTracing(true);
    return mod.trace;
  };

  it("puts the newest step at the TOP, where a crash leaves it visible", async () => {
    // Ordering is the whole feature, not a preference. When PowerPoint dies you
    // get whatever pixels were on screen — no scrolling, no clicking, often a
    // dialog over half the pane. A list that grows downwards puts the last
    // thing that happened at the bottom of a small scrolled box, which is
    // exactly where it cannot be relied on to be visible. Growing upwards puts
    // it one line under the header, always.
    const trace = await paneTrace();
    const steps = document.getElementById("demo-steps")!;
    trace("selftest", "scenario starting", { name: "first" });
    trace("selftest", "scenario starting", { name: "second" });
    const lines = steps.textContent!.trim().split("\n");
    expect(lines[0], `newest was not first — got:\n${steps.textContent}`).toContain("name=second");
    expect(lines[1]).toContain("name=first");
  });

  it("carries the data payload, which is where the phase and the verdict live", async () => {
    const trace = await paneTrace();
    const steps = document.getElementById("demo-steps")!;
    trace("error", "drawing the chart's shapes", { error: "PowerPoint did not respond | at=drawing" });
    expect(steps.textContent).toContain("error");
    expect(steps.textContent).toContain("at=drawing");
  });

  it("keeps the newest when it reaches its cap, not the oldest", async () => {
    // The tail is what is read. Truncating the wrong end would leave a list
    // that is full of the start of a run and silent about how it ended.
    const trace = await paneTrace();
    const steps = document.getElementById("demo-steps")!;
    for (let i = 0; i < 320; i++) trace("draw", "batch issued", { n: i });
    const lines = steps.textContent!.trim().split("\n");
    expect(lines.length).toBeLessThanOrEqual(300);
    expect(lines[0], "the newest step fell off the cap").toContain("n=319");
  });

  it("puts the whole box ABOVE the buttons that start a run", () => {
    // Newest-first settles where in the box the last line is. It says nothing
    // about whether the box is on screen, and a real-host round proved the
    // difference: PowerPoint died, the pane came back, and the log was sitting
    // under nine controls and a paragraph of prose — reachable only by
    // scrolling, which is the one thing a crash does not leave you.
    //
    // Asserted as document order rather than pixels, because jsdom has no
    // layout: `compareDocumentPosition` is what actually decides which of two
    // blocks a scrolled-to-top pane shows first.
    // ONE control is allowed above it, and only one.
    //
    // "Probe, then self-test" sits at the top of the section at the owner's
    // request: it is the starred row of the standing test run and the thing he
    // opens the pane to press, and it was three groups down behind two buttons
    // it supersedes. The rule this test defends is about a log buried under
    // "nine controls and a paragraph", and one button with no prose does not
    // bury anything — so the exception is measured rather than argued, and the
    // count is what stops it growing back into the thing the rule forbids.
    const steps = document.getElementById("demo-steps")!;
    const section = steps.closest("section")!;
    expect(section.querySelector("h2")!.textContent).toContain("Testing");
    const above = [...section.querySelectorAll(".actions")].filter(
      (a) => steps.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_PRECEDING,
    );
    expect(
      above.flatMap((a) => [...a.querySelectorAll("button, select, input")]).map((el) => el.id),
      "more than one control now sits above the live step list — a crash leaves the log off-screen",
    ).toEqual(["demo-round"]);
    // Everything else still follows it. FOLLOWING means `actions` comes after
    // `steps`, which is the order a scrolled-to-top pane draws them in.
    const rest = [...section.querySelectorAll(".actions")].filter((a) => !above.includes(a));
    expect(rest.length, "the run controls vanished from the Testing section").toBeGreaterThan(2);
    for (const a of rest)
      expect(
        steps.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING,
        `${a.querySelector("button")?.id ?? "a control group"} comes BEFORE the live step list`,
      ).toBeTruthy();
    // And the header goes with it, or the box arrives with no Copy button and
    // nothing saying which end is newest.
    const head = section.querySelector(".steps-head")!;
    expect(head.compareDocumentPosition(steps) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(head.compareDocumentPosition(rest[0]) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
  });
});

describe("the style a deck carries", () => {
  beforeEach(bootHostPane);

  /**
   * The palette the chart is actually INSERTED with — read off the config the
   * pane hands the renderer, which is the thing that ends up in the deck.
   *
   * Not off `style-export`, which was the first version of this and measured
   * nothing: that button reads the same resolver the chart does, so a test
   * probing it passed with the chart's own wiring reverted. A probe that cannot
   * fail is the shape this repo keeps finding in its own gates.
   */
  const insertedPalette = async () => {
    host.calls.insertScene = [];
    $("insert").click();
    await settle();
    const tag = (host.calls.insertScene[0] as { tagData?: string }).tagData ?? "{}";
    return (JSON.parse(tag) as ChartConfig).style?.palette;
  };
  /** What `style-export` puts in the box — the file a user would send on. */
  const exported = () => JSON.parse(($("json-io") as HTMLTextAreaElement).value) as { palette?: string[] };

  it("draws with the DECK's style when the deck carries one", async () => {
    await bootHostPane({ deckStyle: { palette: ["#dd1122", "#dd3344"] } });
    await settle();
    expect(await insertedPalette()).toEqual(["#dd1122", "#dd3344"]);
    // …and the file the user would send on says the same thing.
    $("style-export").click();
    expect(exported().palette).toEqual(["#dd1122", "#dd3344"]);
  });

  it("hands the browser's style back when one is imported, and the deck's back on request", async () => {
    await bootHostPane({ deckStyle: { palette: ["#dd1122"] } });
    await settle();
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"] });
    $("style-import").click();
    // An import is an explicit act by the person at the pane; the deck must not
    // silently undo it on the chart they are looking at.
    expect(await insertedPalette()).toEqual(["#22aa33"]);

    $("style-from-deck").click();
    await settle();
    // The note first: inserting a chart writes its own ("Done."), and asserting
    // after that would be asserting the insert's message, not this button's.
    expect($("host-note").textContent).toMatch(/deck/i);
    expect(await insertedPalette()).toEqual(["#dd1122"]);
  });

  it("enables the deck-style buttons when Office arrives AFTER the module ran", async () => {
    // PROVEN ON THE REAL HOST AFTER ROUND 091, in the live pane:
    //
    //     Office.context.host  "PowerPoint"     <- would answer true NOW
    //     style-from-deck      disabled: true
    //     style-to-deck        disabled: true
    //
    // app.ts asked `isPowerPointHost()` at MODULE SCOPE, where `Office.context`
    // does not exist yet, so it answered false for every load inside PowerPoint
    // and disabled both buttons for the whole session. Nothing asked again, so
    // #583's entire user-facing feature was unreachable — and no round could see
    // it, because the thirteen scenarios drive charts and none clicks these.
    hostReadiness.lateOffice();
    await bootHostPane();
    await settle();
    expect(($("style-from-deck") as HTMLButtonElement).disabled, "unreachable on a real host").toBe(false);
    expect(($("style-to-deck") as HTMLButtonElement).disabled).toBe(false);
    hostReadiness.reset();
  });

  it("does not claim the deck is unbranded when the read simply failed", async () => {
    // ROUND 089: `reading the deck's style` hung for its full 90s budget on the
    // real host. The button reported that as "This deck carries no style" and
    // switched the person to the browser's copy — an absence it had not
    // established, about a deck that may well carry one.
    await bootHostPane({ deckStyle: { palette: ["#dd1122"] } });
    await settle();
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"] });
    $("style-import").click();
    expect(await insertedPalette()).toEqual(["#22aa33"]);

    host.deckStyleUnreadable = true;
    $("style-from-deck").click();
    await settle();
    expect($("host-note").textContent, "reported a failed read as an absence").not.toMatch(/carries no style/i);
    expect($("host-note").textContent).toMatch(/could not read/i);
    // AND IT LEFT THE PERSON'S OWN STYLE ALONE. Switching them to the deck on a
    // read that never answered is the half of this that changes their chart.
    expect(await insertedPalette(), "switched away from the imported style on a failed read").toEqual(["#22aa33"]);
  });

  it("saves the style the pane is actually drawing with", async () => {
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"], fontSize: 12 });
    $("style-import").click();
    $("style-to-deck").click();
    await settle();
    expect(host.deckStyleWrites).toEqual([{ palette: ["#22aa33"], fontSize: 12 }]);
    expect($("host-note").textContent).toMatch(/travels with the file/i);
  });

  it("says the deck was NOT updated on a host that cannot store one", async () => {
    host.deckStyleWritable = false;
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"] });
    $("style-import").click();
    $("style-to-deck").click();
    await settle();
    expect($("host-note").textContent).toMatch(/cannot store a deck style/i);
  });

  it("leaves the browser's style in force when the deck carries none", async () => {
    await bootHostPane({ deckStyle: null });
    ($("json-io") as HTMLTextAreaElement).value = JSON.stringify({ palette: ["#22aa33"] });
    $("style-import").click();
    await settle();
    expect(await insertedPalette()).toEqual(["#22aa33"]);
  });
});

describe("the pane's host-friction double", () => {
  it("carries every counter the real module has", async () => {
    // The double named four of eight. Comparing KEY SETS rather than a list of
    // names is the point: a list is what went stale in the first place.
    const actual = (await vi.importActual("../src/render/powerpoint")) as typeof import("../src/render/powerpoint");
    const mocked = (await import("../src/render/powerpoint")) as unknown as {
      hostFrictionCounts: () => Record<string, number>;
    };
    const real = Object.keys(actual.hostFrictionCounts()).sort();
    const fake = Object.keys(mocked.hostFrictionCounts()).sort();
    expect(fake, "the double has drifted from the real counter set").toEqual(real);
  });
});
